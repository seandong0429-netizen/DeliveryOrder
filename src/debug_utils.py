import os
import datetime
import traceback
import sys

def get_log_path():
    try:
        user_dir = os.path.expanduser("~/Documents")
        return os.path.join(user_dir, "DeliveryOrder_DebugLog.txt")
    except:
        return "DeliveryOrder_DebugLog.txt"

LOG_FILE = get_log_path()

def log(message):
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        print(f"Logging failed: {e}")

def log_exception(exc_type, exc_value, exc_traceback):
    msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    log(f"EXCEPTION:\n{msg}")
