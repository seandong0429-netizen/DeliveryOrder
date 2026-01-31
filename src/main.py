import tkinter as tk
from tkinter import ttk, messagebox
import traceback
from ui import DeliveryApp
import sys
import os

import debug_utils

def show_error(exc, val, tb):
    """Global error handler to show errors in a messagebox."""
    err_msg = "".join(traceback.format_exception(exc, val, tb))
    print("Error caught:", err_msg) # Still print to console/log
    debug_utils.log(f"Global Exception caught: {err_msg}")
    
    try:
        # Ensure we have a root window if possible, though messagebox works without one mostly
        messagebox.showerror("Error / 错误", f"An error occurred:\n\n{err_msg}")
    except:
         debug_utils.log("Failed to show messagebox.")

def main():
    debug_utils.log("Application Starting...")
    root = tk.Tk()
    
    # Register global exception handlers
    root.report_callback_exception = show_error
    sys.excepthook = show_error
    
    root.title("简易出货单生成小工具 v3.7 (路径记忆版)")
    root.geometry("1000x700")

    # Set icon at runtime (Mainly for Windows title bar)
    try:
        if hasattr(sys, '_MEIPASS'):
            base = sys._MEIPASS
        else:
             base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        if sys.platform.startswith('win'):
            icon_path = os.path.join(base, 'resources', 'app_icon.ico')
            if os.path.exists(icon_path):
                root.iconbitmap(icon_path)
    except Exception:
        pass

    app = DeliveryApp(root)
    
    root.mainloop()

if __name__ == "__main__":
    main()
