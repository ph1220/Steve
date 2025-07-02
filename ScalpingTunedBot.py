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
    "BASE_ALLOCATION_PERCENT": 0.05,
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
    "HIGH_VIX_TRAILING_STOP": 22.0, # Changed from 20.0 to 22.0
    "LOW_VIX_TRAILING_STOP": 11.0, # Changed from 10.0 to 11.0
    "MIN_VOLUME": 100,
    # "MIN_OPEN_INTEREST": 500,
    # --- NORMAL/LOW VIX - TREND FOLLOWING ---
    "TREND_PROFIT_TARGET_1": 20.0,
    "TREND_PROFIT_TARGET_2": 36.0,
    "TREND_TIGHTENED_STOP": 8.5,
    # --- HIGH VIX - MEAN REVERSION ---
    "REVERSION_PROFIT_TARGET_1": 26.0,
    "REVERSION_PROFIT_TARGET_2": 50.0,
    "REVERSION_TIGHTENED_STOP": 6.0
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
