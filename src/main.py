import tkinter as tk
from tkinter import ttk
from ui import DeliveryApp

def main():
    root = tk.Tk()
    root.title("简易出货单生成小工具 v3.6 (路径记忆版)")
    root.geometry("1000x700")
    
    app = DeliveryApp(root)
    
    root.mainloop()

if __name__ == "__main__":
    main()
