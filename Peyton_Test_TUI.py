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
def get_account_balance():
    if not ib.isConnected():
        logging.warning("⚠️ IB not connected. Cannot fetch account balance.")
        return 0
    try:
        account_summary = ib.accountSummary()
        df = util.df(account_summary)
        # Checks for multi-currency accounts
        cash_row = df[(df['tag'] == 'NetLiquidation') & (df['currency'] == 'USD')]
        if cash_row.empty:
            cash_row = df[df['tag'] == 'NetLiquidation'] # Try without currency constraint
            if cash_row.empty:
                 logging.warning("⚠️ Could not fetch 'NetLiquidation' from account summary.")
                 return 0
        
        cash_value_str = cash_row['value'].values[0]
        if cash_value_str:
            cash = float(cash_value_str)
            logging.info(f"💵 Account Balance (Net Liquidation): ${cash}")
            return cash
        else:
            logging.warning("⚠️ 'NetLiquidation' value is empty.")
            return 0

    except Exception as e:
        logging.error(f"❌ Error fetching account balance: {e}")
        return 0


# =========================
# 💸 Close Open Position
# =========================
def close_position(contract):
    if not ib.isConnected():
        logging.warning("⚠️ IB not connected. Cannot close position.")
        return False

    positions = ib.positions()
    position_closed = False
    for pos in positions:
        if pos.contract.conId == contract.conId:
            qty_to_close = abs(pos.position) # positive quantity for closing order
            action = 'SELL' if pos.position > 0 else 'BUY' # SELL to close long, BUY to close short

            order = MarketOrder(action, qty_to_close)
            trade = ib.placeOrder(contract, order)
            logging.info(f"📤 {action} order placed to close {qty_to_close} of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'}.")

            trade = wait_for_trade_completion(ib, trade)

            if trade.orderStatus.status == 'Filled':
                msg = f"✅ Position closed: {trade.orderStatus.avgFillPrice} x {trade.orderStatus.filled} for {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'}. Status: {trade.orderStatus.status}"
                logging.info(msg)
                send_email(f"Position Closed - {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'}", msg)
                position_closed = True
                break
            else:
                msg = f"❌ Failed to confirm close for {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'} (Status: {trade.orderStatus.status}, Reason: {trade.orderStatus.whyHeld})"
                logging.error(msg)
                send_email(f"Close Error - {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'}", msg)
                position_closed = False
                break
            
    final_position_exists = any(p.contract.conId == contract.conId and p.position != 0 for p in ib.positions())

    if not final_position_exists and not position_closed:
        # This case handles when the order didn't confirm as 'Filled' but the position is gone anyway.
        logging.info(f"ℹ️ Position for {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'} no longer found. Assuming it was closed.")
        position_closed = True # We can now confidently say it's closed.
    elif not position_closed and not final_position_exists:
        # This covers the case where the loop was never entered because the position didn't exist to begin with.
        logging.warning(f"⚠️ No matching open position was found to close.")

    return position_closed

# --- Active Trade Helper Function ---
def wait_for_trade_completion(ib_instance, trade, max_wait_sec=60):
    """Waits for an IB-insync trade to complete or time out."""
    start_time = time.time()
    while trade.isActive() and (time.time() - start_time) < max_wait_sec:
        ib_instance.sleep(1)
    return trade

# =========================
# ⏳ Trailing Stop Monitor (Yahoo Powered)
# =========================
def monitor_position_with_trailing(spy_ticker ,strike, direction, expiry, entry_price, contract, dynamic_trailing_percent):
    highest_price = entry_price
    contract_display_name = contract.localSymbol if hasattr(contract, 'localSymbol') else f"{contract.symbol} {strike}{direction} {expiry}"

    logging.info(f"🛡️ Monitoring {contract_display_name} with entry {entry_price:.2f}, initial trailing stop at {dynamic_trailing_percent}%.")

    while True:
        if not is_market_open():
            logging.info(f"⏰ Market closed while monitoring {contract_display_name}. Attempting to close position.")
            if not close_position(contract):
                error_msg = f"ALERT: Attempt to close {contract_display_name} at EOD did not confirm 'Filled'."
                logging.error(error_msg)
                send_email(f"Potential Issue: EOD Close - {contract_display_name}", error_msg)
            break # Exit monitoring loop

        current_price = get_option_price_yahoo(spy_ticker, expiry, strike, direction)

        if current_price is None:
            logging.warning(f"⚠️ No price data for {contract_display_name}, retrying in 10s...")
            time.sleep(10)
            continue

        if current_price > highest_price:
            highest_price = current_price
            logging.info(f"🔺 New High for {contract_display_name}: {highest_price:.2f}")

        stop_price = highest_price * (1 - dynamic_trailing_percent / 100)
        stop_price = round(stop_price, 2)

        logging.info(f"📊 {contract_display_name} - Current: {current_price:.2f} | High: {highest_price:.2f} | Stop: {stop_price:.2f} (Trail: {dynamic_trailing_percent}%)")

        if current_price <= stop_price:
            logging.warning(f"🛑 Trailing stop hit for {contract_display_name} at {current_price:.2f} (Stop: {stop_price:.2f})! Attempting to close...")
            if not close_position(contract):
                error_msg = f"ALERT: Attempt to close {contract_display_name} on stop-loss did not confirm 'Filled'."
                logging.error(error_msg)
                send_email(f"Potential Issue: Stop Loss Close - {contract_display_name}", error_msg)
            break

        time.sleep(30)

# =========================
# 🧠 Trading Logic
# =========================
def trade_spy_options():
    # --- Essential Pre-checks ---
    if not is_market_open():
        logging.info("⏰ Market is closed. Skipping trade evaluation.")
        return

    spy_ticker = yf.Ticker("SPY")
    
    price, _ = get_spy_price(spy_ticker)
    if price is None:
        send_email("Bot Error - SPY Price", "Failed to fetch SPY price. Trade aborted.")
        return

    sma, rsi = get_tech_indicators(spy_ticker)
    if sma is None or rsi is None or math.isnan(sma) or math.isnan(rsi):
        send_email("Bot Error - Indicators", "Failed to fetch valid indicators (SMA or RSI). Trade aborted.")
        return

    current_vix = None
    try:
        vix_ticker = yf.Ticker("^VIX")
        vix_data = vix_ticker.history(period="1d") # gets the latest closing VIX
        if not vix_data.empty:
            current_vix = vix_data['Close'].iloc[-1]
            logging.info(f"📊 Current VIX: {current_vix:.2f}")
        else:
            logging.warning("⚠️ Could not fetch VIX data. Defaulting to normal VIX logic for signals and risk.")
            send_email("Bot Warning - VIX Fetch", "Could not fetch VIX data. Using default parameters.")
    except Exception as e:
        logging.error(f"❌ Error fetching VIX: {e}. Defaulting to normal VIX logic for signals and risk.")
        send_email("Bot Error - VIX Fetch", f"Error fetching VIX: {str(e)}. Using default parameters.")

    direction = None
    trade_rationale = ""

    if current_vix is not None and current_vix > strategy_config["VIX_HIGH_SIGNAL_THRESHOLD"]:
        logging.info(f"🔶 High VIX ({current_vix:.2f}): Applying High VIX (Mean Reversion) signal logic.")
        if rsi < strategy_config["RSI_HIGH_VIX_OVERSOLD"]:
            direction = "C" # Buy Call on extreme oversold
            trade_rationale = f"High VIX Mean Reversion: RSI {rsi:.2f} < {strategy_config['RSI_HIGH_VIX_OVERSOLD']}"
        elif rsi > strategy_config["RSI_HIGH_VIX_OVERBOUGHT"]:
            direction = "P" # Buy Put on extreme overbought
            trade_rationale = f"High VIX Mean Reversion: RSI {rsi:.2f} > {strategy_config['RSI_HIGH_VIX_OVERBOUGHT']}"
    else: # Normal or Low VIX, or VIX not available will use the original trend-following logic
        vix_status_for_signal = "Normal/Low VIX"
        if current_vix is None:
            vix_status_for_signal = "VIX N/A"
        elif current_vix < strategy_config["VIX_HIGH_SIGNAL_THRESHOLD"] : # Covers low and normal
             vix_status_for_signal = f"VIX {current_vix:.2f}"

        logging.info(f"🔷 {vix_status_for_signal}: Applying Standard Trend signal logic.")
        if price > sma and rsi < strategy_config["RSI_STD_OVERBOUGHT"]:
            direction = "C"
            trade_rationale = f"Standard Trend: Price > SMA, RSI {rsi:.2f} < {strategy_config['RSI_STD_OVERBOUGHT']}"
        elif price < sma and rsi > strategy_config["RSI_STD_OVERSOLD"]:
            direction = "P"
            trade_rationale = f"Standard Trend: Price < SMA, RSI {rsi:.2f} > {strategy_config['RSI_STD_OVERSOLD']}"

    if direction is None:
        logging.info(f"🚦 No trade signal based on current logic (VIX: {current_vix if current_vix else 'N/A'}, RSI: {rsi:.2f}, Price/SMA: {price:.2f}/{sma:.2f}).")
        return

    logging.info(f"🚀 Trade Signal: {direction} | Rationale: {trade_rationale}")

    # --- Dynamic Risk Adjustment Based On VIX ---
    base_allocation_percentage = strategy_config["BASE_ALLOCATION_PERCENT"]
    base_trailing_percent = strategy_config["BASE_TRAILING_PERCENT"]
    current_allocation_percentage = base_allocation_percentage
    current_trailing_percent = base_trailing_percent
  
    vix_risk_profile = "Default"
    if current_vix is not None:
        if current_vix > strategy_config["VIX_HIGH_RISK_THRESHOLD"]:
            current_allocation_percentage = base_allocation_percentage * strategy_config["HIGH_VIX_ALLOCATION_MULT"]
            current_trailing_percent = strategy_config["HIGH_VIX_TRAILING_STOP"]
            vix_risk_profile = f"High VIX ({current_vix:.2f})"
            logging.info(f"🔶 {vix_risk_profile}: Adjusting risk - Allocation to {current_allocation_percentage*100:.1f}%, Trailing to {current_trailing_percent}%.")
        elif current_vix < strategy_config["VIX_LOW_RISK_THRESHOLD"]:
            current_allocation_percentage = base_allocation_percentage * strategy_config["LOW_VIX_ALLOCATION_MULT"]
            current_trailing_percent = strategy_config["LOW_VIX_TRAILING_STOP"]
            vix_risk_profile = f"Low VIX ({current_vix:.2f})"
            logging.info(f"🔷 {vix_risk_profile}: Adjusting risk - Allocation to {current_allocation_percentage*100:.1f}%, Trailing to {current_trailing_percent}%.")
        else: # Normal VIX range
            vix_risk_profile = f"Normal VIX ({current_vix:.2f})"
            logging.info(f"🔷 {vix_risk_profile}: Using default risk - Allocation {current_allocation_percentage*100:.1f}%, Trailing {current_trailing_percent}%.")
    else:
        logging.warning("⚠️ VIX data not available for risk adjustment, using default risk parameters.")
        vix_risk_profile = "VIX N/A (Default Risk)"


    # --- Option Selection and Trade Execution ---
    available_expiries = spy_ticker.options
    if len(available_expiries) < 2:
        logging.warning("⚠️ Not enough expiry dates available for SPY options. Skipping trade.")
        send_email("Trade Error - Expiry", "Not enough SPY option expiry dates available.")
        return
    expiry = available_expiries[1]
    strike = round(price)

    option_price = get_option_price_yahoo(spy_ticker, expiry, strike, direction)
    if option_price is None or option_price <= 0 or math.isnan(option_price):
        logging.warning(f"⚠️ Invalid or zero option price (${option_price}) for {direction} {strike} {expiry}. Skipping trade.")
        send_email("Trade Error - Option Price", f"Option price for {direction} {strike} {expiry} is invalid (${option_price}). Skipping trade.")
        return

    contract = Option(
        symbol='SPY',
        lastTradeDateOrContractMonth=expiry.replace("-", ""),
        strike=strike,
        right=direction,
        exchange='SMART',
        currency='USD'
    )
    
    logging.info(f"Qualifying contract: {contract.symbol} {contract.lastTradeDateOrContractMonth} {contract.strike} {contract.right}")
    qualified_contracts = ib.qualifyContracts(contract)
    if not qualified_contracts:
        logging.warning(f"⚠️ Contract could not be qualified: {contract}. Skipping trade.")
        send_email("Trade Error - Qualification", f"Failed to qualify contract: {contract}")
        return
    contract = qualified_contracts[0]
    logging.info(f"📄 Qualified Contract: {contract.localSymbol if hasattr(contract, 'localSymbol') else contract}")


    balance = get_account_balance()
    if balance <= 0:
        logging.warning("⚠️ Account balance is zero or could not be fetched. Skipping trade.")
        send_email("Trade Error - Balance", "Account balance is zero or could not be fetched.")
        return

    allocation_amount = balance * current_allocation_percentage
    logging.info(f"💵 Calculated allocation amount: ${allocation_amount:.2f} ({current_allocation_percentage*100:.1f}% of balance ${balance:.2f})")

    cost_per_contract = option_price * 100 # Options are typically for 100 shares
    if cost_per_contract <= 0 :
        logging.warning(f"⚠️ Cost per contract is zero or negative (${cost_per_contract:.2f}). Skipping trade.")
        send_email("Trade Error - Contract Cost", f"Cost per contract is ${cost_per_contract:.2f}. Skipping trade.")
        return

    qty = math.floor(allocation_amount / cost_per_contract)

    if qty < 1:
        logging.warning(f"⚠️ Not enough balance for {current_allocation_percentage*100:.1f}% allocation. Need ${cost_per_contract:.2f} for 1 contract (Option Price: ${option_price:.2f}), have ${allocation_amount:.2f} allocated for trade. Balance: ${balance:.2f}")
        send_email("Trade Info - Insufficient Allocation", f"Insufficient funds for 1 contract at current allocation. Need ${cost_per_contract:.2f}, allocated ${allocation_amount:.2f}.")
        return

    logging.info(f"🧮 Attempting to buy {qty} contract(s) of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'} at ~${option_price:.2f} each. (Risk Profile: {vix_risk_profile})")

    order = MarketOrder('BUY', qty)
    trade = ib.placeOrder(contract, order)
    logging.info(f"📥 Buy order placed for {qty} of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'}.")

    trade = wait_for_trade_completion(ib, trade)

    if trade.orderStatus.status != 'Filled':
        msg = f"❌ Buy order for {qty} of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'} failed or not filled. Status: {trade.orderStatus.status}, Reason: {trade.orderStatus.whyHeld}"
        logging.error(msg)
        send_email(f"Trade Error - Buy Order {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'}", msg)
        return

    entry_price_filled = trade.orderStatus.avgFillPrice
    filled_qty = trade.orderStatus.filled
    logging.info(f"✅ Entry filled: {filled_qty} contract(s) of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'} at ${entry_price_filled:.2f} each.")

    email_subject = f"Trade Executed: {direction} {qty} {contract.localSymbol if hasattr(contract, 'localSymbol') else 'SPY Option'}"
    email_body = (
        f"Strategy Signal: {trade_rationale}\n"
        f"VIX Context: {vix_risk_profile}\n"
        f"Action: BUY {direction} Option\n"
        f"Contract: {contract.localSymbol if hasattr(contract, 'localSymbol') else contract}\n"
        f"Quantity: {filled_qty}\n"
        f"Entry Price: ${entry_price_filled:.2f}\n"
        f"Allocation: {current_allocation_percentage*100:.1f}%\n"
        f"Trailing Stop: {current_trailing_percent}%"
    )
    send_email(email_subject, email_body)

    # Pass contract to the monitoring function
    monitor_position_with_trailing(spy_ticker, strike, direction, expiry, entry_price_filled, contract, current_trailing_percent)

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

if __name__ == "__main__":
    app = TradingApp()
    app.run()
