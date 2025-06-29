import asyncio
from textual.app import App, ComposeResult
from textual.containers import Container
from textual.widgets import Header, Footer, Static, RichLog, Button, Input
from textual.reactive import reactive
from ib_insync import IB, util, MarketOrder, Option, Contract # Using * can cause issues later on if updates happen
from ta.momentum import RSIIndicator
from ta.trend import SMAIndicator
import yfinance as yf
from datetime import datetime, time as dtime
import pytz
import time
import os
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
import math
import logging
from dataclasses import dataclass, asdict

@dataclass
class StrategyConfig:
    BASE_ALLOCATION_PERCENT: float = 0.05
    BASE_TRAILING_PERCENT: float = 15.0
    VIX_HIGH_SIGNAL_THRESHOLD: float = 24.0
    VIX_HIGH_RISK_THRESHOLD: float = 24.0
    VIX_LOW_RISK_THRESHOLD: float = 17.0
    RSI_HIGH_VIX_OVERSOLD: float = 25.0
    RSI_HIGH_VIX_OVERBOUGHT: float = 75.0
    RSI_STD_OVERSOLD: float = 30.0
    RSI_STD_OVERBOUGHT: float = 70.0
    HIGH_VIX_ALLOCATION_MULT: float = 0.5
    LOW_VIX_ALLOCATION_MULT: float = 1.15
    HIGH_VIX_TRAILING_STOP: float = 20.0
    LOW_VIX_TRAILING_STOP: float = 10.0

# =========================
# 🔐 Load ENV Variables
# =========================
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

class TextualHandler(logging.Handler):
    def __init__(self, rich_log: RichLog):
        super().__init__()
        self.rich_log = rich_log

    def emit(self, record):
        msg = self.format(record)
        self.rich_log.write(msg)

# =========================
# 📧 Email Alert
# =========================
def send_email(subject, body):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = EMAIL_USER
    msg['To'] = EMAIL_USER

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
        logging.info("✅ Email sent!")
    except Exception as e:
        logging.error(f"❌ Email error: {e}")

# =========================
# ⏰ Market Hours Check
# =========================
def is_market_open():
    eastern = pytz.timezone('US/Eastern')
    now = datetime.now(eastern).time()
    return dtime(9, 30) <= now <= dtime(16, 0)

# =========================
# 📈 Get SPY Price from Yahoo
# =========================
def get_spy_price(spy_ticker):
    data = spy_ticker.history(period="1d", interval="1m")
    if data.empty:
        logging.warning("⚠️ No price data.")
        return None, None
    latest = data.iloc[-1]
    price = latest['Close']
    timestamp = latest.name
    logging.info(f"💰 SPY Price: {price} at {timestamp}")
    return price, timestamp

# =========================
# 📊 Technical Indicators
# =========================
def get_tech_indicators(spy_ticker):
    data = spy_ticker.history(period="3d", interval="5m")
    if data.empty:
        logging.warning("⚠️ Failed to fetch historical data for indicators.")
        return None, None

    close = data['Close']
    sma = SMAIndicator(close, window=14).sma_indicator().iloc[-1]
    rsi = RSIIndicator(close, window=14).rsi().iloc[-1]

    logging.info(f"📈 SMA: {sma:.2f} | RSI: {rsi:.2f}")
    return sma, rsi

# =========================
# 💸 Option Price from Yahoo
# =========================
def get_option_price_yahoo(spy_ticker, expiry, strike, direction):
    try:
        opt_chain = spy_ticker.option_chain(expiry)
        chain = opt_chain.calls if direction == 'C' else opt_chain.puts

        opt = chain[chain['strike'] == strike]
        if opt.empty:
            logging.warning(f"⚠️ Option not found for strike {strike} and expiry {expiry}.")
            return None

        bid = opt['bid'].values[0]
        ask = opt['ask'].values[0]
        last = opt['lastPrice'].values[0]

        if (bid == 0 and ask == 0) and last > 0:
            mid = last
        elif bid == 0 and ask == 0 and last == 0:
            logging.warning("⚠️ Option has no bid, ask, or last price. Unreliable.")
            return None
        else:
            mid = round((bid + ask) / 2, 2)

        if mid <= 0:
            logging.warning(f"⚠️ Invalid option mid-price ({mid}). Bid: {bid}, Ask: {ask}, Last: {last}")
            return None

        logging.info(f"💵 Option Price ({direction} {strike} {expiry}): Bid={bid} | Ask={ask} | Mid={mid}")
        return mid

    except Exception as e:
        logging.error(f"❌ Failed to fetch option price for {direction} {strike} {expiry}: {e}")
        return None

# =========================
# 💰 Account Balance
# =========================
async def get_account_balance(ib: IB):
    if not ib.isConnected():
        logging.warning("⚠️ IB not connected. Cannot fetch account balance.")
        return 0
    try:
        # More direct and async-friendly way to get the value
        summary = await ib.accountSummaryAsync()
        net_liq_value = next((v.value for v in summary if v.tag == 'NetLiquidation' and v.currency == 'USD'), None)

        if net_liq_value:
            cash = float(net_liq_value)
            logging.info(f"💵 Account Balance (Net Liquidation): ${cash}")
            return cash
        else:
            logging.warning("⚠️ Could not fetch 'NetLiquidation' from account summary.")
            return 0
    except Exception as e:
        logging.error(f"❌ Error fetching account balance: {e}")
        return 0

# =========================
# 💸 Close Open Position
# =========================
async def close_position(ib: IB, contract: Contract):
    if not ib.isConnected():
        logging.warning("⚠️ IB not connected. Cannot close position.")
        return False

    positions = await ib.positionsAsync()
    position_closed = False
    for pos in positions:
        if pos.contract.conId == contract.conId:
            action = 'SELL' if pos.position > 0 else 'BUY'
            order = MarketOrder(action, abs(pos.position))
            trade = await ib.placeOrderAsync(contract, order)
            logging.info(f"📤 {action} order placed to close {abs(pos.position)} of {contract.localSymbol}.")

            trade = await wait_for_trade_completion(ib, trade)

            if trade.orderStatus.status == 'Filled':
                msg = f"✅ Position closed: {trade.orderStatus.avgFillPrice} x {trade.orderStatus.filled}"
                logging.info(msg)
                send_email(f"Position Closed - {contract.localSymbol}", msg)
                return True # Exit successfully
            else:
                msg = f"❌ Failed to confirm close for {contract.localSymbol} (Status: {trade.orderStatus.status})"
                logging.error(msg)
                send_email(f"Close Error - {contract.localSymbol}", msg)
                return False # Exit with failure

    logging.warning(f"⚠️ No matching open position was found for {contract.localSymbol} to close.")
    return True # If no position is found, it's effectively "closed"
    
# --- Active Trade Helper Function ---
async def wait_for_trade_completion(ib: IB, trade, max_wait_sec=60):
    """Waits for an IB-insync trade to complete or time out."""
    start_time = time.time()
    while trade.isActive() and (time.time() - start_time) < max_wait_sec:
        await ib.sleepAsync(1) # Use async sleep
    return trade

# =========================
# ⏳ Trailing Stop Monitor (Yahoo Powered)
# =========================
async def monitor_position_with_trailing(ib: IB, spy_ticker, strike, direction, expiry, entry_price, contract, dynamic_trailing_percent):
    highest_price = entry_price
    contract_display_name = contract.localSymbol
    logging.info(f"🛡️ Monitoring {contract_display_name} with entry {entry_price:.2f}, trailing stop at {dynamic_trailing_percent}%.")

    while True:
        if not is_market_open():
            logging.info(f"⏰ Market closed. Attempting to close {contract_display_name}.")
            await close_position(ib, contract)
            break

        current_price = get_option_price_yahoo(spy_ticker, expiry, strike, direction)
        if current_price is None:
            logging.warning(f"⚠️ No price data for {contract_display_name}, retrying in 10s...")
            await asyncio.sleep(10)
            continue

        if current_price > highest_price:
            highest_price = current_price

        stop_price = round(highest_price * (1 - dynamic_trailing_percent / 100), 2)
        logging.info(f"📊 {contract_display_name} - Current: {current_price:.2f} | High: {highest_price:.2f} | Stop: {stop_price:.2f}")

        if current_price <= stop_price:
            logging.warning(f"🛑 Trailing stop hit for {contract_display_name}! Attempting to close...")
            await close_position(ib, contract)
            break

        await asyncio.sleep(30)
        
# =========================
# 🧠 Trading Logic
# =========================
async def trade_spy_options(ib: IB, config: StrategyConfig):
    # --- Essential Pre-checks ---
    if not is_market_open():
        logging.info("⏰ Market is closed. Skipping trade evaluation.")
        return

    # Check for existing positions first
    positions = await ib.positionsAsync()
    if any(p.contract.symbol == 'SPY' and p.contract.secType == 'OPT' and p.position != 0 for p in positions):
        logging.info("⚠️ Active SPY option position found. Skipping new trade initiation.")
        return

    spy_ticker = yf.Ticker("SPY")
    price, _ = get_spy_price(spy_ticker)
    if price is None: return

    sma, rsi = get_tech_indicators(spy_ticker)
    if sma is None or rsi is None: return

    # --- Signal Logic (using the config object) ---
    direction = None
    # ... (The entire signal logic block remains the same, just change strategy_config["KEY"] to config.KEY)
    # For brevity, I'll show one example, apply it to all:
    # if rsi < strategy_config["RSI_HIGH_VIX_OVERSOLD"] becomes:
    # if rsi < config.RSI_HIGH_VIX_OVERSOLD
    
    # NOTE: I have refactored your logic below to use the new config object.
    # Please replace your entire function with this.
    current_vix = yf.Ticker("^VIX").history(period="1d")['Close'].iloc[-1]
    trade_rationale = ""
    if current_vix > config.VIX_HIGH_SIGNAL_THRESHOLD:
        if rsi < config.RSI_HIGH_VIX_OVERSOLD: direction = "C"; trade_rationale = "High VIX Mean Reversion"
        elif rsi > config.RSI_HIGH_VIX_OVERBOUGHT: direction = "P"; trade_rationale = "High VIX Mean Reversion"
    else:
        if price > sma and rsi < config.RSI_STD_OVERBOUGHT: direction = "C"; trade_rationale = "Standard Trend"
        elif price < sma and rsi > config.RSI_STD_OVERSOLD: direction = "P"; trade_rationale = "Standard Trend"

    if direction is None:
        logging.info("🚦 No trade signal based on current logic.")
        return

    logging.info(f"🚀 Trade Signal: {direction} | Rationale: {trade_rationale}")

    # --- Dynamic Risk Adjustment ---
    current_allocation_percentage = config.BASE_ALLOCATION_PERCENT
    current_trailing_percent = config.BASE_TRAILING_PERCENT
    if current_vix > config.VIX_HIGH_RISK_THRESHOLD:
        current_allocation_percentage *= config.HIGH_VIX_ALLOCATION_MULT
        current_trailing_percent = config.HIGH_VIX_TRAILING_STOP
    elif current_vix < config.VIX_LOW_RISK_THRESHOLD:
        current_allocation_percentage *= config.LOW_VIX_ALLOCATION_MULT
        current_trailing_percent = config.LOW_VIX_TRAILING_STOP

    # --- Option Selection ---
    expiry = spy_ticker.options[1]
    strike = round(price)
    option_price = get_option_price_yahoo(spy_ticker, expiry, strike, direction)
    if option_price is None or option_price <= 0: return

    # --- Trade Execution ---
    contract = Option('SPY', expiry.replace("-", ""), strike, direction, 'SMART', currency='USD')
    qualified_contracts = await ib.qualifyContractsAsync(contract)
    if not qualified_contracts:
        logging.warning(f"⚠️ Contract could not be qualified. Skipping trade.")
        return
    contract = qualified_contracts[0]

    balance = await get_account_balance(ib)
    if balance <= 0: return

    cost_per_contract = option_price * 100
    qty = math.floor((balance * current_allocation_percentage) / cost_per_contract)
    if qty < 1:
        logging.warning(f"⚠️ Not enough balance for allocation. Need ${cost_per_contract:.2f} for 1 contract.")
        return

    logging.info(f"🧮 Attempting to buy {qty} contract(s) of {contract.localSymbol}")
    order = MarketOrder('BUY', qty)
    trade = await ib.placeOrderAsync(contract, order)
    trade = await wait_for_trade_completion(ib, trade)

    if trade.orderStatus.status != 'Filled':
        logging.error(f"❌ Buy order failed. Status: {trade.orderStatus.status}")
        return

    entry_price_filled = trade.orderStatus.avgFillPrice
    logging.info(f"✅ Entry filled: {trade.orderStatus.filled} contract(s) at ${entry_price_filled:.2f}")
    send_email(f"Trade Executed: BUY {qty} {contract.localSymbol}", f"Filled at ${entry_price_filled:.2f}")

    await monitor_position_with_trailing(ib, spy_ticker, strike, direction, expiry, entry_price_filled, contract, current_trailing_percent)
    
class SettingsPanel(Static):
    """A widget to display and edit strategy configuration."""

    def __init__(self, config: StrategyConfig, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config = config

    def compose(self) -> ComposeResult:
        """Create the input fields and save button."""
        yield Static("Strategy Configuration", classes="title")
        # Create an input for each config value using a loop
        for key, value in asdict(self.config).items():
            yield Static(key.replace("_", " ").title(), classes="label")
            yield Input(value=str(value), id=f"input_{key}")
        yield Button("Save Changes", variant="primary", id="save_button")

class TradingApp(App):
    """A Textual app to monitor and control the trading bot."""

    TITLE = "Trading Bot"
    SUB_TITLE = "SPY Option Strategy"

    config = reactive(StrategyConfig(), layout=True)
  
    def on_mount(self) -> None:
        log_viewer = self.query_one(RichLog)

        textual_handler = TextualHandler(log_viewer)
        file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        textual_handler.setFormatter(file_formatter)
        logging.getLogger().addHandler(textual_handler)

        logging.getLogger().setLevel(logging.INFO)

        file_handler = logging.FileHandler("peyton_test_trading_bot.log")
        file_handler.setFormatter(file_formatter)
        logging.getLogger().addHandler(file_handler)

        logging.info("✅ TUI Initialized. Logging is now configured.")
        logging.info("🤖 Trading worker will start after setup...")

        self.run_worker(self.run_trading_worker, exclusive=True, group="trading_worker")
    
    def compose(self) -> ComposeResult:
        """Create child widgets for the app."""
        yield Header()
        yield RichLog(id="log_viewer", wrap=True, highlight=True)
        yield SettingsPanel(self.config, id="settings_panel")
        yield Footer()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        """Handle the save button press event."""
        if event.button.id == "save_button":
            try:
                new_config = StrategyConfig() # Create a new config object
                # Loop through the keys and update from the input widgets
                for key in asdict(new_config).keys():
                    input_widget = self.query_one(f"#input_{key}", Input)
                    # Convert the input's string value to a float
                    setattr(new_config, key, float(input_widget.value))

                self.config = new_config # Update the reactive attribute
                logging.info("✅ Strategy configuration updated successfully!")

            except ValueError:
                logging.error("❌ Invalid input. Please enter numbers only.")
            except Exception as e:
                logging.error(f"❌ Failed to save configuration: {e}")

    async def run_trading_worker(self) -> None:
        """The main trading logic loop, running as an async task."""
        ib = IB()
        try:
            logging.info("🤖 Worker connecting to IBKR...")
            await ib.connectAsync('127.0.0.1', 7497, clientId=int(time.time() % 1000) + 100)
            logging.info(f"🔗 Connected to IBKR with Client ID: {ib.clientId}.")

            while True:
                logging.info(f"\n--- 🔄 Checking for trades... ---")
                try:
                    await trade_spy_options(ib, self.config)
                except Exception as e:
                    import traceback
                    logging.error(f"❌ Unhandled error in trade logic: {e}\n{traceback.format_exc()}")

                main_loop_sleep_seconds = 300
                logging.info(f"--- 😴 Sleeping for {main_loop_sleep_seconds // 60} minutes. ---")
                await asyncio.sleep(main_loop_sleep_seconds)

        except ConnectionRefusedError:
            logging.error("❌ IBKR Connection Refused. Is TWS/Gateway running?")
        except Exception as e:
            import traceback
            logging.critical(f"❌ CRITICAL WORKER ERROR: {e}\n{traceback.format_exc()}")
        finally:
            if ib.isConnected():
                logging.info("🔗 Disconnecting from IBKR.")
                ib.disconnect()

if __name__ == "__main__":
    app = TradingApp()
    app.run()
