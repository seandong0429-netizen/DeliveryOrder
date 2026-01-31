import tkinter as tk
from tkinter import ttk
from ui import DeliveryApp

def main():
    root = tk.Tk()
    root.title("简易出货单生成小工具 v3.7 (路径记忆版)")
    root.geometry("1000x700")

    # Set icon at runtime (Mainly for Windows title bar)
    import os, sys
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
