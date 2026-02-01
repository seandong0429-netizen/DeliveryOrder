import os
import datetime
import traceback
import sys

def get_log_path():
    try:
        # 1. Determine Base Dir (Logic mirrored from logic.py to avoid circular import)
        if getattr(sys, 'frozen', False):
            # Frozen (APP/EXE)
            base = os.path.dirname(sys.executable)
            if sys.platform == 'darwin' and 'Contents/MacOS' in base:
                # macOS .app bundle -> Go up to folder containing .app
                base_dir = os.path.abspath(os.path.join(base, '../../..'))
            else:
                # Windows/Linux or raw binary
                base_dir = base
        else:
            # Source Code -> Project Root
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # 2. Check Write Permission
        if os.access(base_dir, os.W_OK):
            # Special Cleanliness Rule for /Applications root
            if base_dir.rstrip(os.sep) == '/Applications':
                target_dir = os.path.join(base_dir, "DeliveryOrderData")
                if not os.path.exists(target_dir):
                     try:
                         os.makedirs(target_dir, exist_ok=True)
                     except:
                         target_dir = os.path.expanduser("~/Documents/DeliveryOrderData")
            else:
                # Portable Mode: Log next to app
                target_dir = base_dir
        else:
            # Installed/Restricted Mode: Log to Documents/Data
            target_dir = os.path.expanduser("~/Documents/DeliveryOrderData")
            if not os.path.exists(target_dir):
                os.makedirs(target_dir, exist_ok=True)
        
        return os.path.join(target_dir, "debug.log")
        
    except Exception as e:
        # Absolute fallback
        return os.path.expanduser("~/Documents/DeliveryOrder_Emergency_Log.txt")


LOG_FILE = get_log_path()

def manage_log_size():
    """Rotate log if it gets too big (e.g., > 1MB)."""
    try:
        if os.path.exists(LOG_FILE):
            size = os.path.getsize(LOG_FILE)
            if size > 1 * 1024 * 1024: # 1MB
                backup = LOG_FILE + ".bak"
                if os.path.exists(backup):
                    os.remove(backup)
                os.rename(LOG_FILE, backup)
                # print(f"Log rotated: {LOG_FILE} -> {backup}")
    except Exception as e:
        print(f"Log rotation failed: {e}")

# Check size on startup
manage_log_size()

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
