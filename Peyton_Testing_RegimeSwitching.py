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

def save_trade_state(contract, entry_price, quantity, trailing_percent):
    """Saves the active trade's state to a file."""
    state = {
        "is_position_open": True,
        # util.contract_to_dict converts the IB contract object into a savable format
        "contract": util.contract_to_dict(contract),
        "entry_price": entry_price,
        "quantity": quantity,
        "trailing_percent": trailing_percent
    }
    with open(STATE_FILE, 'w') as f:
        json.dump(state, f, indent=4)
    logging.info(f"Trade state saved for {contract.localSymbol}.")

def load_trade_state():
    """Loads the trade state from a file."""
    if not os.path.exists(STATE_FILE):
        return {"is_position_open": False} # Default state if no file exists
    try:
        with open(STATE_FILE, 'r') as f:
            state = json.load(f)
            # Convert the dictionary back into an ib_insync Contract object
            if state.get("is_position_open"):
                state['contract'] = Contract(**state['contract'])
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
                clear_trade_state()
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

    return position_closed

# --- Active Trade Helper Function ---
def wait_for_trade_completion(ib_instance, trade, max_wait_sec=60):
    """Waits for an IB-insync trade to complete or time out."""
    start_time = time.time()
    while trade.isActive() and (time.time() - start_time) < max_wait_sec:
        ib_instance.sleep(1)
    return trade

# =========================
#   Trailing Stop Monitor (IBKR Paper-Trading Powered)
# =========================
def monitor_position_with_trailing(contract, entry_price, dynamic_trailing_percent):
    highest_price = entry_price
    contract_display_name = contract.localSymbol

    logging.info(f"Monitoring {contract_display_name} with entry {entry_price:.2f}, initial trailing stop at {dynamic_trailing_percent}%.")

    while True:
        if not is_market_open():
            logging.info(f"Market closed while monitoring {contract_display_name}. Attempting to close position.")
            if not close_position(contract):
                error_msg = f"ALERT: Attempt to close {contract_display_name} at EOD did not confirm 'Filled'."
                logging.error(error_msg)
                send_email(f"Potential Issue: EOD Close - {contract_display_name}", error_msg)
            break

        ticker = get_option_snapshot(contract)

        if ticker is None:
            logging.warning(f"Could not get a valid snapshot for {contract_display_name}. Retrying in 15s...")
            ib.sleep(15) 
            continue
        
        # Extract the price from the ticker
        current_price = ticker.last
        if current_price != current_price: # Check for NaN
            current_price = (ticker.bid + ticker.ask) / 2
        
        current_price = round(current_price, 2)

        if current_price <= 0:
            logging.warning(f"Invalid price ({current_price}) from snapshot for {contract_display_name}. Retrying in 15s...")
            ib.sleep(15)
            continue

        if current_price > highest_price:
            highest_price = current_price
            logging.info(f"New High for {contract_display_name}: {highest_price:.2f}")

        stop_price = highest_price * (1 - dynamic_trailing_percent / 100)
        stop_price = round(stop_price, 2)

        logging.info(f"{contract_display_name} - Current: {current_price:.2f} | High: {highest_price:.2f} | Stop: {stop_price:.2f} (Trail: {dynamic_trailing_percent}%)")

        if current_price <= stop_price:
            logging.warning(f"Trailing stop hit for {contract_display_name} at {current_price:.2f} (Stop: {stop_price:.2f})! Attempting to close...")
            if not close_position(contract):
                error_msg = f"ALERT: Attempt to close {contract_display_name} on stop-loss did not confirm 'Filled'."
                logging.error(error_msg)
                send_email(f"Potential Issue: Stop Loss Close - {contract_display_name}", error_msg)
            break

        ib.sleep(20)

# =========================
#          Trading Logic
# =========================
def trade_spy_options():
    # --- Essential Pre-checks ---
    if not is_market_open():
        logging.info("Market is closed. Skipping trade evaluation.")
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
            logging.info(f"Current VIX: {current_vix:.2f}")
        else:
            logging.warning("Could not fetch VIX data. Defaulting to normal VIX logic for signals and risk.")
            send_email("Bot Warning - VIX Fetch", "Could not fetch VIX data. Using default parameters.")
    except Exception as e:
        logging.error(f"Error fetching VIX: {e}. Defaulting to normal VIX logic for signals and risk.")
        send_email("Bot Error - VIX Fetch", f"Error fetching VIX: {str(e)}. Using default parameters.")

# --- CHANGE THESE VALUES HERE ONLY ---
    strategy_config = {
        "BASE_ALLOCATION_PERCENT": 0.05,
        "BASE_TRAILING_PERCENT": 15.0,
        "VIX_HIGH_SIGNAL_THRESHOLD": 24.0,
        "VIX_HIGH_RISK_THRESHOLD": 24.0,
        "VIX_LOW_RISK_THRESHOLD": 17.0,
        "RSI_HIGH_VIX_OVERSOLD": 25.0,
        "RSI_HIGH_VIX_OVERBOUGHT": 75.0,
        "RSI_STD_OVERSOLD": 30.0,
        "RSI_STD_OVERBOUGHT": 70.0,
        "HIGH_VIX_ALLOCATION_MULT": 0.5,
        "LOW_VIX_ALLOCATION_MULT": 1.15,
        "HIGH_VIX_TRAILING_STOP": 20.0,
        "LOW_VIX_TRAILING_STOP": 10.0,
        "MIN_VOLUME": 100,
        "MIN_OPEN_INTEREST": 500,
    }
  
    direction = None
    trade_rationale = ""

    if current_vix is not None and current_vix > strategy_config["VIX_HIGH_SIGNAL_THRESHOLD"]:
        logging.info(f"High VIX ({current_vix:.2f}): Applying High VIX (Mean Reversion) signal logic.")
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

        logging.info(f"{vix_status_for_signal}: Applying Standard Trend signal logic.")
        if price > sma and rsi < strategy_config["RSI_STD_OVERBOUGHT"]:
            direction = "C"
            trade_rationale = f"Standard Trend: Price > SMA, RSI {rsi:.2f} < {strategy_config['RSI_STD_OVERBOUGHT']}"
        elif price < sma and rsi > strategy_config["RSI_STD_OVERSOLD"]:
            direction = "P"
            trade_rationale = f"Standard Trend: Price < SMA, RSI {rsi:.2f} > {strategy_config['RSI_STD_OVERSOLD']}"

    if direction is None:
        logging.info(f"No trade signal based on current logic (VIX: {current_vix if current_vix else 'N/A'}, RSI: {rsi:.2f}, Price/SMA: {price:.2f}/{sma:.2f}).")
        return

    logging.info(f"Trade Signal: {direction} | Rationale: {trade_rationale}")

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
            logging.info(f"{vix_risk_profile}: Adjusting risk - Allocation to {current_allocation_percentage*100:.1f}%, Trailing to {current_trailing_percent}%.")
        elif current_vix < strategy_config["VIX_LOW_RISK_THRESHOLD"]:
            current_allocation_percentage = base_allocation_percentage * strategy_config["LOW_VIX_ALLOCATION_MULT"]
            current_trailing_percent = strategy_config["LOW_VIX_TRAILING_STOP"]
            vix_risk_profile = f"Low VIX ({current_vix:.2f})"
            logging.info(f"{vix_risk_profile}: Adjusting risk - Allocation to {current_allocation_percentage*100:.1f}%, Trailing to {current_trailing_percent}%.")
        else: # Normal VIX range
            vix_risk_profile = f"Normal VIX ({current_vix:.2f})"
            logging.info(f"{vix_risk_profile}: Using default risk - Allocation {current_allocation_percentage*100:.1f}%, Trailing {current_trailing_percent}%.")
    else:
        logging.warning("VIX data not available for risk adjustment, using default risk parameters.")
        vix_risk_profile = "VIX N/A (Default Risk)"


    # --- Option Selection and Trade Execution ---
    available_expiries = spy_ticker.options
    if len(available_expiries) < 2:
        logging.warning("Not enough expiry dates available for SPY options. Skipping trade.")
        send_email("Trade Error - Expiry", "Not enough SPY option expiry dates available.")
        return
    expiry = available_expiries[1]
    strike = round(price)

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
        logging.warning(f"Contract could not be qualified: {contract}. Skipping trade.")
        send_email("Trade Error - Qualification", f"Failed to qualify contract: {contract}")
        return
    contract = qualified_contracts[0]
    logging.info(f"Qualified Contract: {contract.localSymbol}")

    ticker = get_option_snapshot(contract)
    if ticker is None:
        logging.warning(f"Could not get market data for {contract.localSymbol}. Skipping trade.")
        return

    volume = ticker.volume if ticker.volume == ticker.volume else 0 # Handle NaN volume
    open_interest = ticker.openInterest if ticker.openInterest == ticker.openInterest else 0 # Handle NaN OI
    min_volume = strategy_config["MIN_VOLUME"]
    min_open_interest = strategy_config["MIN_OPEN_INTEREST"]

    logging.info(f"Liquidity Check for {contract.localSymbol}: Volume={volume}, Open Interest={open_interest}")

    if volume < min_volume or open_interest < min_open_interest:
        logging.warning(f"TRADE REJECTED: {contract.localSymbol} failed liquidity check. "
                        f"Vol ({volume}) < MinVol ({min_volume}) or "
                        f"OI ({open_interest}) < MinOI ({min_open_interest}).")
        send_email(f"Trade Rejected - Illiquid", f"Contract {contract.localSymbol} was rejected due to low liquidity.")
        return

    logging.info("Liquidity check passed.")

    price = ticker.last
    if price != price: # Check for NaN
        price = (ticker.bid + ticker.ask) / 2
    
    option_price = round(price, 2)

    if option_price <= 0:
        logging.warning(f"Invalid or zero option price (${option_price}) from IBKR for {contract.localSymbol}. Skipping trade.")
        send_email("Trade Error - Option Price", f"IBKR option price for {contract.localSymbol} is invalid (${option_price}). Skipping trade.")
        return


    balance = get_account_balance()
    if balance <= 0:
        logging.warning("Account balance is zero or could not be fetched. Skipping trade.")
        send_email("Trade Error - Balance", "Account balance is zero or could not be fetched.")
        return

    allocation_amount = balance * current_allocation_percentage
    logging.info(f"Calculated allocation amount: ${allocation_amount:.2f} ({current_allocation_percentage*100:.1f}% of balance ${balance:.2f})")

    cost_per_contract = option_price * 100 # Options are typically for 100 shares
    if cost_per_contract <= 0 :
        logging.warning(f"Cost per contract is zero or negative (${cost_per_contract:.2f}). Skipping trade.")
        send_email("Trade Error - Contract Cost", f"Cost per contract is ${cost_per_contract:.2f}. Skipping trade.")
        return

    qty = math.floor(allocation_amount / cost_per_contract)

    if qty < 1:
        logging.warning(f"Not enough balance for {current_allocation_percentage*100:.1f}% allocation. Need ${cost_per_contract:.2f} for 1 contract (Option Price: ${option_price:.2f}), have ${allocation_amount:.2f} allocated for trade. Balance: ${balance:.2f}")
        send_email("Trade Info - Insufficient Allocation", f"Insufficient funds for 1 contract at current allocation. Need ${cost_per_contract:.2f}, allocated ${allocation_amount:.2f}.")
        return

    logging.info(f"Attempting to buy {qty} contract(s) of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'} at ~${option_price:.2f} each. (Risk Profile: {vix_risk_profile})")

    order = MarketOrder('BUY', qty)
    trade = ib.placeOrder(contract, order)
    logging.info(f"Buy order placed for {qty} of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'}.")

    trade = wait_for_trade_completion(ib, trade)

    if trade.orderStatus.status != 'Filled':
        msg = f"Buy order for {qty} of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'} failed or not filled. Status: {trade.orderStatus.status}, Reason: {trade.orderStatus.whyHeld}"
        logging.error(msg)
        send_email(f"Trade Error - Buy Order {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'}", msg)
        return

    entry_price_filled = trade.orderStatus.avgFillPrice
    filled_qty = trade.orderStatus.filled
    logging.info(f"Entry filled: {filled_qty} contract(s) of {contract.localSymbol if hasattr(contract, 'localSymbol') else 'option'} at ${entry_price_filled:.2f} each.")

    save_trade_state(contract, entry_price_filled, filled_qty, current_trailing_percent)
    
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
    monitor_position_with_trailing(contract, entry_price_filled, current_trailing_percent)

# =========================
#         Main Loop
# =========================
ib = IB()
try:
    logging.info("Attempting to connect to IBKR...")
    ib.connect('127.0.0.1', 7497, clientId=int(time.time() % 1000) + 100)
    logging.info(f"Connected to IBKR with Client ID: {ib.client.clientId}.")

    ib.reqMarketDataType(3)
    
    # Create a Ticker object for SPY to be reused
    spy_ticker = yf.Ticker("SPY")

    while True:
        logging.info(f"\n--- Main Loop Iteration ({datetime.now(pytz.timezone('US/Eastern')).strftime('%Y-%m-%d %H:%M:%S %Z')}) ---")
        
        # Load the state at the beginning of every loop
        trade_state = load_trade_state()

        if trade_state["is_position_open"]:
            logging.warning("RECOVERY MODE: Active trade found in state file. Resuming monitoring.")
            
            # Extract details from the loaded state
            contract = trade_state["contract"]
            entry_price = trade_state["entry_price"]
            trailing_percent = trade_state["trailing_percent"]
            
            # We must qualify the contract again after a restart
            ib.qualifyContracts(contract)

            # Directly start monitoring. We will replace this function in the next guide.
            # Note: The arguments for the old function are still available if needed, but the new one is cleaner.
            expiry_str = contract.lastTradeDateOrContractMonth
            expiry_formatted = f"{expiry_str[:4]}-{expiry_str[4:6]}-{expiry_str[6:]}"
            monitor_position_with_trailing(contract, entry_price, trailing_percent)
            logging.info("Monitoring finished. Returning to main loop.")

        elif is_market_open():
            logging.info("Market is open. Checking for new trading opportunities...")
            # The original logic to check for a new trade only runs if no position is open.
            # We can simplify the check here, as our state file is the source of truth.
            # A failsafe check against actual positions is still good practice.
            current_positions = ib.positions()
            spy_option_position_open = any(
                p.contract.symbol == 'SPY' and p.contract.secType == 'OPT' and p.position != 0
                for p in (current_positions or [])
            )
            if not spy_option_position_open:
                 trade_spy_options()
            else:
                 logging.warning("IBKR shows an open SPY position, but state file is clear. Please check manually. Skipping new trades.")

        else:
            logging.info(f"Market closed. Sleeping...")
            # (Your EOD check logic can remain here)

        main_loop_sleep_seconds = 300 # 5 minutes
        logging.info(f"--- End of Loop Iteration. Sleeping for {main_loop_sleep_seconds // 60} minutes. ---")
        time.sleep(main_loop_sleep_seconds)

except ConnectionRefusedError:
    logging.error(f"IBKR Connection Refused. Ensure TWS/Gateway is running, API connections are enabled, and correct port/IP.")
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
