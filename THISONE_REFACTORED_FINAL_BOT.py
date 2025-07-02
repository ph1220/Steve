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
import coloredlogs
import json

# =========================
#   Load ENV Variables
# =========================
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# =========================
#   Configure Logging
# =========================
# 1. Get the root logger
logger = logging.getLogger()
logger.setLevel(logging.INFO) # Set the lowest level to capture

# 2. Create a handler for writing to the log file
file_handler = logging.FileHandler("peyton_test_trading_bot.log")
file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

# 3. Install coloredlogs for the console output
# This will automatically create and add a colored StreamHandler
coloredlogs.install(
    level='INFO',
    logger=logger,
    fmt='%(asctime)s - %(levelname)s - %(message)s'
)

# =========================
#   Email Alert
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
        logging.info("Email sent!")
    except Exception as e:
        logging.error(f"Email error: {e}")


STATE_FILE = "trade_state.json"

# --- ADJUST THESE VALUES AS NEEDED ---
strategy_config = {
    "BASE_ALLOCATION_PERCENT": 0.03, # Changed from 0.05 to 0.03
    "BASE_TRAILING_PERCENT": 11.0, # Changed from 15.0 to 11.0
    "VIX_HIGH_SIGNAL_THRESHOLD": 25.0, # Changed from 24 to 25
    "VIX_HIGH_RISK_THRESHOLD": 24.0,
    "VIX_LOW_RISK_THRESHOLD": 17.0,
    "RSI_HIGH_VIX_OVERSOLD": 26.0, # Changed from 25 to 26
    "RSI_HIGH_VIX_OVERBOUGHT": 74.0, # Changed from 75 to 74
    "RSI_STD_OVERSOLD": 28.0, # Changed from 30 to 28
    "RSI_STD_OVERBOUGHT": 72.0, # Changed from 70 to 72
    "HIGH_VIX_ALLOCATION_MULT": 0.35, # Changed from 0.5 to 0.35
    "LOW_VIX_ALLOCATION_MULT": 1.25, # Changed from 1.15 to 1.25
    "HIGH_VIX_TRAILING_STOP": 20.0,
    "LOW_VIX_TRAILING_STOP": 10.0,
    "MIN_VOLUME": 100,
    # "MIN_OPEN_INTEREST": 500,
    # --- NORMAL/LOW VIX - TREND FOLLOWING ---
    "TREND_PROFIT_TARGET_1": 25.0,  # Move to breakeven at 25%
    "TREND_PROFIT_TARGET_2": 50.0,  # Tighten stop at 50%
    "TREND_TIGHTENED_STOP": 8.0,    # The tightened stop for trends
    # --- HIGH VIX - MEAN REVERSION ---
    "REVERSION_PROFIT_TARGET_1": 15.0,  # Move to breakeven at only 15%
    "REVERSION_PROFIT_TARGET_2": 30.0,  # Tighten stop at only 30%
    "REVERSION_TIGHTENED_STOP": 5.0     # An even tighter stop for choppy markets
}

def save_trade_state(contract, entry_price, quantity, highest_price, trailing_percent, active_regime, breakeven_activated=False, profit_lock_activated=False):
    """Saves the active trade's state, including the active trading regime."""
    contract_state = {
        'conId': contract.conId,
        'symbol': contract.symbol,
        'lastTradeDateOrContractMonth': contract.lastTradeDateOrContractMonth,
        'strike': contract.strike,
        'right': contract.right,
        'exchange': contract.exchange,
        'currency': contract.currency,
        'localSymbol': contract.localSymbol
    }
    
    state = {
        "is_position_open": True,
        "contract": contract_state,
        "entry_price": entry_price,
        "quantity": quantity,
        "highest_price": highest_price,
        "trailing_percent": trailing_percent,
        "active_regime": active_regime,
        "breakeven_activated": breakeven_activated,
        "profit_lock_activated": profit_lock_activated
    }
    with open(STATE_FILE, 'w') as f:
        json.dump(state, f, indent=4)
    logging.info(f"Trade state saved for {contract.localSymbol}.")

def load_trade_state():
    """Loads the trade state, including the active trading regime."""
    if not os.path.exists(STATE_FILE):
        return {"is_position_open": False}
    try:
        with open(STATE_FILE, 'r') as f:
            state = json.load(f)
            if state.get("is_position_open"):
                state['contract'] = Contract(**state['contract'])
                # Load the flags, defaulting to False if they don't exist
                state['breakeven_activated'] = state.get('breakeven_activated', False)
                state['profit_lock_activated'] = state.get('profit_lock_activated', False)
                # Load the regime, defaulting to "TREND" for backward compatibility
                state['active_regime'] = state.get('active_regime', 'TREND')
            return state
    except (json.JSONDecodeError, IOError) as e:
        logging.error(f"Error loading state file: {e}. Resetting state.")
        return {"is_position_open": False}

def clear_trade_state():
    """Clears the trade state by resetting the file."""
    state = {"is_position_open": False}
    with open(STATE_FILE, 'w') as f:
        json.dump(state, f, indent=4)
    logging.info("Trade state has been cleared.")

# =========================
#   Market Hours Check
# =========================
def is_market_open():
    eastern = pytz.timezone('US/Eastern')
    now = datetime.now(eastern).time()
    return dtime(9, 30) <= now <= dtime(16, 0)

# =========================
#    Option Snapshot from IBKR
# =========================
def get_option_snapshot(contract):
    """
    Fetches a full market data snapshot (ticker) for a contract from IBKR.
    This includes price, volume, open interest, etc.

    Args:
        contract: The ib_insync Contract object.

    Returns:
        The Ticker object if successful, otherwise None.
    """
    
    ticker = ib.reqMktData(contract, '', True, False)
    
    ib.sleep(2) 

    if ticker is None:
        logging.warning(f"IBKR did not return a ticker object for {contract.localSymbol}.")
        return None
    
    if ticker.last != ticker.last and ticker.bid != ticker.bid: # Check for NaN on both last and bid
        logging.warning(f"Ticker for {contract.localSymbol} returned but contains no valid data.")
        return None

    return ticker

# =========================
#    Get SPY Price from Yahoo
# =========================
def get_spy_price(spy_ticker):
    data = spy_ticker.history(period="1d", interval="1m")
    if data.empty:
        logging.warning("No price data.")
        return None, None
    latest = data.iloc[-1]
    price = latest['Close']
    timestamp = latest.name
    logging.info(f"SPY Price: {price} at {timestamp}")
    return price, timestamp

# =========================
#   Technical Indicators
# =========================
def get_tech_indicators(spy_ticker):
    data = spy_ticker.history(period="3d", interval="5m")
    if data.empty:
        logging.warning("Failed to fetch historical data for indicators.")
        return None, None

    close = data['Close']
    sma = SMAIndicator(close, window=14).sma_indicator().iloc[-1]
    rsi = RSIIndicator(close, window=14).rsi().iloc[-1]

    logging.info(f"SMA: {sma:.2f} | RSI: {rsi:.2f}")
    return sma, rsi

# =========================
#    Option Price from Yahoo (OLD AND NO LONGER USED BUT NOT DELETING YET JUST IN CASE)
# =========================
# def get_option_price_yahoo(spy_ticker, expiry, strike, direction):
#    try:
#        opt_chain = spy_ticker.option_chain(expiry)
#         chain = opt_chain.calls if direction == 'C' else opt_chain.puts
# 
#         opt = chain[chain['strike'] == strike]
#         if opt.empty:
#             logging.warning(f"Option not found for strike {strike} and expiry {expiry}.")
#             return None
# 
#         bid = opt['bid'].values[0]
#         ask = opt['ask'].values[0]
#         last = opt['lastPrice'].values[0]
#
#         if (bid == 0 and ask == 0) and last > 0:
#             mid = last
#         elif bid == 0 and ask == 0 and last == 0:
#             logging.warning("Option has no bid, ask, or last price. Unreliable.")
#             return None
#         else:
#             mid = round((bid + ask) / 2, 2)
#
#         if mid <= 0:
#             logging.warning(f"Invalid option mid-price ({mid}). Bid: {bid}, Ask: {ask}, Last: {last}")
#             return None
#
#         logging.info(f"Option Price ({direction} {strike} {expiry}): Bid={bid} | Ask={ask} | Mid={mid}")
#         return mid
#
#     except Exception as e:
#         logging.error(f"Failed to fetch option price for {direction} {strike} {expiry}: {e}")
#         return None

# =========================
#    Account Balance
# =========================
def get_account_balance():
    if not ib.isConnected():
        logging.warning("IB not connected. Cannot fetch account balance.")
        return 0
    try:
        account_summary = ib.accountSummary()
        df = util.df(account_summary)
        # Checks for multi-currency accounts
        cash_row = df[(df['tag'] == 'NetLiquidation') & (df['currency'] == 'USD')]
        if cash_row.empty:
            cash_row = df[df['tag'] == 'NetLiquidation'] # Try without currency constraint
            if cash_row.empty:
                 logging.warning("Could not fetch 'NetLiquidation' from account summary.")
                 return 0
        
        cash_value_str = cash_row['value'].values[0]
        if cash_value_str:
            cash = float(cash_value_str)
            logging.info(f"Account Balance (Net Liquidation): ${cash}")
            return cash
        else:
            logging.warning("'NetLiquidation' value is empty.")
            return 0

    except Exception as e:
        logging.error(f"Error fetching account balance: {e}")
        return 0


# =========================
#    Close Open Position
# =========================
def close_position(contract):
    if not ib.isConnected():
        logging.warning("IB not connected. Cannot close position.")
        return False

    positions = ib.positions()
    position_closed = False
    for pos in positions:
        if pos.contract.conId == contract.conId:
            qty_to_close = abs(pos.position) # positive quantity for closing order
            action = 'SELL' if pos.position > 0 else 'BUY' # SELL to close long, BUY to close short

            order = MarketOrder(action, qty_to_close)
            trade = ib.placeOrder(contract, order)
            logging.info(f"{action} order placed to close {qty_to_close} of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'}.")

            trade = wait_for_trade_completion(ib, trade)

            if trade.orderStatus.status == 'Filled':
                msg = f"Position closed: {trade.orderStatus.avgFillPrice} x {trade.orderStatus.filled} for {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'}. Status: {trade.orderStatus.status}"
                logging.info(msg)
                send_email(f"Position Closed - {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'}", msg)
                position_closed = True
                break
            else:
                msg = f"Failed to confirm close for {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'} (Status: {trade.orderStatus.status}, Reason: {trade.orderStatus.whyHeld})"
                logging.error(msg)
                send_email(f"Close Error - {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'}", msg)
                position_closed = False
                break
            
    final_position_exists = any(p.contract.conId == contract.conId and p.position != 0 for p in ib.positions())

    if not final_position_exists and not position_closed:
        # This case handles when the order didn't confirm as 'Filled' but the position is gone anyway.
        logging.info(f"Position for {contract.localSymbol if hasattr(contract, 'localSymbol') else 'contract'} no longer found. Assuming it was closed.")
        position_closed = True # We can now confidently say it's closed.
    elif not position_closed and not final_position_exists:
        # This covers the case where the loop was never entered because the position didn't exist to begin with.
        logging.warning(f"No matching open position was found to close.")

    if position_closed:
        clear_trade_state()

    return position_closed

# --- Active Trade Helper Function ---
def wait_for_trade_completion(ib_instance, trade, max_wait_sec=60):
    """Waits for an IB-insync trade to complete or time out."""
    start_time = time.time()
    while trade.isActive() and (time.time() - start_time) < max_wait_sec:
        ib_instance.sleep(1)
    return trade

# =========================
#   Intelligent Trailing Stop Monitor
# =========================
def monitor_position_with_trailing(contract, entry_price, quantity, dynamic_trailing_percent, strategy_config, active_profit_targets, active_regime):
    contract_display_name = contract.localSymbol
    
    trade_state = load_trade_state()
    highest_price = trade_state.get('highest_price', entry_price)
    breakeven_activated = trade_state.get('breakeven_activated', False)
    profit_lock_activated = trade_state.get('profit_lock_activated', False)
    current_trailing_percent = trade_state.get('trailing_percent', dynamic_trailing_percent)

    logging.info(f"Monitoring {contract_display_name} in {active_regime} mode with entry {entry_price:.2f}, initial stop at {current_trailing_percent}%.")
    logging.info(f"Using profit targets: Breakeven at {active_profit_targets['target_1']}%, Tighten Stop at {active_profit_targets['target_2']}%")
    if breakeven_activated: logging.warning("RECOVERY: Breakeven stop is already active.")
    if profit_lock_activated: logging.warning("RECOVERY: Profit lock-in stop is already active.")

    while True:
        if not is_market_open(): break

        if not entry_price or entry_price <= 0:
            logging.error(f"CRITICAL MONITORING ERROR: Invalid entry price ({entry_price}) for {contract_display_name}. Cannot calculate gain. Aborting monitor.")
            send_email(f"Bot Critical Error - Invalid Entry Price", f"Monitoring for {contract_display_name} has been aborted due to an invalid entry price of {entry_price}.")
            break # Exit the monitoring loop safely

        ticker = get_option_snapshot(contract)
        if ticker is None: continue
        
        current_price = ticker.last
        if current_price != current_price: current_price = (ticker.bid + ticker.ask) / 2
        current_price = round(current_price, 2)
        if current_price <= 0: continue

        state_changed = False

        if current_price > highest_price:
            highest_price = current_price
            logging.info(f"New High for {contract_display_name}: {highest_price:.2f}")
            state_changed = True

        current_gain_percent = ((current_price - entry_price) / entry_price) * 100

        if not breakeven_activated and current_gain_percent >= active_profit_targets["target_1"]:
            breakeven_activated = True
            logging.warning(f"PROFIT TARGET 1 HIT (+{current_gain_percent:.1f}%). Stop loss is now at breakeven (${entry_price:.2f}).")
            state_changed = True

        if not profit_lock_activated and current_gain_percent >= active_profit_targets["target_2"]:
            profit_lock_activated = True
            breakeven_activated = True # Hitting target 2 automatically implies target 1 is also hit
            current_trailing_percent = active_profit_targets["tightened_stop"]
            logging.warning(f"PROFIT TARGET 2 HIT (+{current_gain_percent:.1f}%). Trailing stop tightened to {current_trailing_percent}%.")
            state_changed = True

        if state_changed:
            save_trade_state(contract, entry_price, quantity, highest_price, current_trailing_percent, active_regime, breakeven_activated, profit_lock_activated)

        trailing_stop_price = highest_price * (1 - current_trailing_percent / 100)
        final_stop_price = max(trailing_stop_price, entry_price) if breakeven_activated else trailing_stop_price
        final_stop_price = round(final_stop_price, 2)

        logging.info(f"{contract_display_name} - Gain: {current_gain_percent:+.1f}% | Current: {current_price:.2f} | High: {highest_price:.2f} | Stop: {final_stop_price:.2f} (Trail: {current_trailing_percent}%)")

        if current_price <= final_stop_price:
            logging.warning(f"Stop hit for {contract_display_name} at {current_price:.2f} (Stop: {final_stop_price:.2f})! Attempting to close...")
            if not close_position(contract):
                error_msg = f"ALERT: Attempt to close {contract_display_name} on stop-loss did not confirm 'Filled'."
                logging.error(error_msg)
                send_email(f"Potential Issue: Stop Loss Close - {contract_display_name}", error_msg)
            break

        ib.sleep(20)

# =========================
#          Trading Logic
# =========================
def trade_spy_options(spy_ticker):
    if not is_market_open():
        logging.info("Market is closed. Skipping trade evaluation.")
        return

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
        vix_data = vix_ticker.history(period="1d")
        if not vix_data.empty:
            current_vix = vix_data['Close'].iloc[-1]
            logging.info(f"Current VIX: {current_vix:.2f}")
    except Exception as e:
        logging.error(f"Error fetching VIX: {e}.")

    # --- Determine Trading Regime and Parameters ---
    is_high_vix = current_vix is not None and current_vix > strategy_config["VIX_HIGH_SIGNAL_THRESHOLD"]
    regime_name = "REVERSION" if is_high_vix else "TREND"

    direction = None
    trade_rationale = ""
    if is_high_vix:
        logging.info(f"High VIX ({current_vix:.2f}): Applying {regime_name} signal logic.")
        if rsi < strategy_config["RSI_HIGH_VIX_OVERSOLD"]: direction, trade_rationale = "C", f"High VIX Mean Reversion: RSI {rsi:.2f} < {strategy_config['RSI_HIGH_VIX_OVERSOLD']}"
        elif rsi > strategy_config["RSI_HIGH_VIX_OVERBOUGHT"]: direction, trade_rationale = "P", f"High VIX Mean Reversion: RSI {rsi:.2f} > {strategy_config['RSI_HIGH_VIX_OVERBOUGHT']}"
    else:
        vix_status = f"VIX {current_vix:.2f}" if current_vix is not None else "VIX N/A"
        logging.info(f"{vix_status}: Applying {regime_name} signal logic.")
        if price > sma and rsi < strategy_config["RSI_STD_OVERBOUGHT"]: direction, trade_rationale = "C", f"Standard Trend: Price > SMA, RSI {rsi:.2f} < {strategy_config['RSI_STD_OVERBOUGHT']}"
        elif price < sma and rsi > strategy_config["RSI_STD_OVERSOLD"]: direction, trade_rationale = "P", f"Standard Trend: Price < SMA, RSI {rsi:.2f} > {strategy_config['RSI_STD_OVERSOLD']}"

    if direction is None:
        logging.info("No trade signal based on current logic.")
        return
    logging.info(f"Trade Signal: {direction} | Rationale: {trade_rationale}")

    # --- Set Dynamic Parameters based on Regime ---
    if is_high_vix:
        current_allocation_percentage = strategy_config["BASE_ALLOCATION_PERCENT"] * strategy_config["HIGH_VIX_ALLOCATION_MULT"]
        current_trailing_percent = strategy_config["HIGH_VIX_TRAILING_STOP"]
        active_profit_targets = {"target_1": strategy_config["REVERSION_PROFIT_TARGET_1"], "target_2": strategy_config["REVERSION_PROFIT_TARGET_2"], "tightened_stop": strategy_config["REVERSION_TIGHTENED_STOP"]}
    else: # Normal or Low VIX
        current_allocation_percentage = strategy_config["BASE_ALLOCATION_PERCENT"] * strategy_config["LOW_VIX_ALLOCATION_MULT"]
        current_trailing_percent = strategy_config["LOW_VIX_TRAILING_STOP"]
        active_profit_targets = {"target_1": strategy_config["TREND_PROFIT_TARGET_1"], "target_2": strategy_config["TREND_PROFIT_TARGET_2"], "tightened_stop": strategy_config["TREND_TIGHTENED_STOP"]}

    # --- Option Selection and Execution ---
    available_expiries = spy_ticker.options
    if len(available_expiries) < 2:
        logging.error(f"Not enough expiration dates found for SPY. Found: {len(available_expiries)}. Need at least 2. Aborting trade.")
        return
    # We select the second expiry (weekly) to avoid the nearest-term options
    expiry = available_expiries[1]

    # Fetch the entire option chain for the chosen expiry
    opt_chain = spy_ticker.option_chain(expiry)
    chain_df = opt_chain.calls if direction == 'C' else opt_chain.puts

    # Ensure the chain is not empty
    if chain_df.empty:
        logging.warning(f"No {direction} options found for expiry {expiry}. Aborting trade.")
        return

    # Find the strike in the chain that is mathematically closest to the current SPY price
    closest_strike_index = (chain_df['strike'] - price).abs().idxmin()
    strike = chain_df.loc[closest_strike_index]['strike']
    logging.info(f"Target price is {price:.2f}. Closest available strike is {strike}.")

    contract = Option('SPY', expiry.replace("-", ""), strike, direction, 'SMART', currency='USD')
    
    qualified_contracts = ib.qualifyContracts(contract)
    if not qualified_contracts: return
    contract = qualified_contracts[0]
    logging.info(f"Qualified Contract: {contract.localSymbol}")

    ticker = get_option_snapshot(contract)
    if ticker is None: return

    volume = ticker.volume if ticker.volume == ticker.volume else 0
    if volume < strategy_config["MIN_VOLUME"]:
        logging.warning(f"TRADE REJECTED: {contract.localSymbol} failed liquidity check. Volume ({volume}) < MinVol ({strategy_config['MIN_VOLUME']}).")
        return
    logging.info(f"Liquidity check passed. Volume={volume}")

    price = ticker.last if ticker.last == ticker.last else (ticker.bid + ticker.ask) / 2
    option_price = round(price, 2)
    if option_price <= 0: return

    balance = get_account_balance()
    if balance <= 0: return
    
    allocation_amount = balance * current_allocation_percentage
    cost_per_contract = option_price * 100
    if cost_per_contract <= 0: return
    qty = math.floor(allocation_amount / cost_per_contract)
    if qty < 1:
        logging.warning(f"Not enough balance for 1 contract. Need ${cost_per_contract:.2f}, allocated ${allocation_amount:.2f}.")
        return

    order = MarketOrder('BUY', qty)
    trade = ib.placeOrder(contract, order)
    trade = wait_for_trade_completion(ib, trade)
    if trade.orderStatus.status != 'Filled':
        logging.error(f"Buy order for {qty} of {contract.localSymbol} failed. Status: {trade.orderStatus.status}")
        return

    entry_price_filled = trade.orderStatus.avgFillPrice
    filled_qty = trade.orderStatus.filled
    logging.info(f"Entry filled: {filled_qty} contract(s) of {contract.localSymbol} at ${entry_price_filled:.2f} each.")

    save_trade_state(contract, entry_price_filled, filled_qty, entry_price_filled, current_trailing_percent, regime_name)
    
    email_subject = f"Trade Executed: {direction} {qty} {contract.localSymbol}"
    email_body = f"Strategy Signal: {trade_rationale}\n" # ... (rest of email body)
    send_email(email_subject, email_body)

    monitor_position_with_trailing(contract, entry_price_filled, filled_qty, current_trailing_percent, strategy_config, active_profit_targets, regime_name)

# =========================
#         Main Loop
# =========================
ib = IB()
try:
    logging.info("Attempting to connect to IBKR...")
    ib.connect('127.0.0.1', 7497, clientId=int(time.time() % 1000) + 100)
    logging.info(f"Connected to IBKR with Client ID: {ib.client.clientId}.")
    ib.reqMarketDataType(3)
    
    

    while True:
        logging.info(f"\n--- Main Loop Iteration ({datetime.now(pytz.timezone('US/Eastern')).strftime('%Y-%m-%d %H:%M:%S %Z')}) ---")
        spy_ticker = yf.Ticker("SPY")
        trade_state = load_trade_state()

        if trade_state.get("is_position_open"):
            logging.warning("RECOVERY MODE: Active trade found in state file. Resuming monitoring.")
            
            contract = trade_state["contract"]
            entry_price = trade_state["entry_price"]
            trailing_percent = trade_state["trailing_percent"]
            active_regime = trade_state.get("active_regime", "TREND")
            
            if active_regime == "REVERSION":
                active_profit_targets = {"target_1": strategy_config["REVERSION_PROFIT_TARGET_1"], "target_2": strategy_config["REVERSION_PROFIT_TARGET_2"], "tightened_stop": strategy_config["REVERSION_TIGHTENED_STOP"]}
            else:
                active_profit_targets = {"target_1": strategy_config["TREND_PROFIT_TARGET_1"], "target_2": strategy_config["TREND_PROFIT_TARGET_2"], "tightened_stop": strategy_config["TREND_TIGHTENED_STOP"]}
            
            qualified_contracts = ib.qualifyContracts(contract)
            if not qualified_contracts:
                logging.error(f"RECOVERY FAILED: Contract {contract.localSymbol} from state file is expired or invalid. Clearing state.")
                send_email("Bot Recovery Failure", f"The contract {contract.localSymbol} could not be qualified, likely because it has expired. The position is assumed closed or worthless. The state file has been cleared.")
                clear_trade_state()
                continue # Skip to the next main loop iteration

            # If successful, use the qualified contract
            contract = qualified_contracts[0]
            quantity = trade_state["quantity"]
            monitor_position_with_trailing(contract, entry_price, quantity, trailing_percent, strategy_config, active_profit_targets, active_regime)
            logging.info("Monitoring finished. Returning to main loop.")

        elif is_market_open():
            logging.info("Market is open. Checking for new trading opportunities...")
    
            # Check for any existing SPY option positions in the account
            current_positions = ib.positions()
            spy_option_positions = [p for p in (current_positions or []) if p.contract.symbol == 'SPY' and p.contract.secType == 'OPT' and p.position != 0]

            if not spy_option_positions:
                # No position exists, clear to look for a new trade
                trade_spy_options(spy_ticker)
            else:
                # A position exists on IBKR, but our state file is clear. This is a mismatch.
                # SAFEST ACTION: Close the unknown position to prevent unexpected behavior.
                ghost_position = spy_option_positions[0] # Take the first one found
                contract_display = ghost_position.contract.localSymbol
        
                error_msg = f"STATE MISMATCH: An unknown SPY Option position for {contract_display} (Qty: {ghost_position.position}) was found in IBKR without a state file. Attempting to close it for safety."
                logging.error(error_msg)
                send_email(f"Bot Safety Alert: Closing Unknown Position", error_msg)
        
                # Attempt to close the position
                close_position(ghost_position.contract)

        else:
            logging.info("Market closed. Sleeping...")

        main_loop_sleep_seconds = 300
        logging.info(f"--- End of Loop Iteration. Sleeping for {main_loop_sleep_seconds // 60} minutes. ---")
        time.sleep(main_loop_sleep_seconds)

except ConnectionRefusedError:
    logging.error("IBKR Connection Refused. Ensure TWS/Gateway is running and API connections are enabled.")
    send_email("Bot Critical Error - Connection", "IBKR Connection Refused.")
except Exception as e:
    import traceback
    tb_str = traceback.format_exc()
    logging.error(f"Main loop or connection critical error: {e}")
    logging.info(tb_str)
    send_email("Bot Critical Error - Main", f"Main loop or connection error: {str(e)}\n\nTraceback:\n{tb_str}")
finally:
    if ib.isConnected():
        logging.info("Disconnecting from IBKR.")
        ib.disconnect()
    else:
        logging.info("Was not connected to IBKR or already disconnected.")
