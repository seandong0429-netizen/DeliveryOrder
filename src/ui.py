import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import datetime
import os
from logic import ProductManager, OrderNumberGenerator
from history import HistoryManager
from export_pdf import export_pdf
from export_excel import export_to_excel
import sys

class DeliveryApp:
    def __init__(self, root):
        self.root = root
        self.product_manager = ProductManager()
        self.order_generator = OrderNumberGenerator()
        self.history_manager = HistoryManager()
        
        self.current_items = []
        self.history_displayed_items = [] # For history tab checkboxes
        
        self.setup_menu()
        self.setup_ui()
        
    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件 (File)", menu=file_menu)
        
        file_menu.add_command(label="导入产品资料 (Import Products)...", command=self.import_products)
        file_menu.add_separator()
        file_menu.add_command(label="退出 (Exit)", command=self.root.quit)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助 (Help)", menu=help_menu)
        help_menu.add_command(label="关于 (About)", command=self.show_about)

    def show_about(self):
        # Create custom window
        top = tk.Toplevel(self.root)
        top.title("关于 / About")
        top.geometry("400x450")
        top.resizable(False, False)
        
        # 1. Image
        img_name = "wechat_qr.png"
        img_full_path = ""
        
        # Check PyInstaller Temp Bundle first (if bundled)
        if hasattr(sys, '_MEIPASS'):
            img_full_path = os.path.join(sys._MEIPASS, img_name)
        else:
            # Normal / External fallback
            if getattr(sys, 'frozen', False):
                base = os.path.dirname(sys.executable)
                if sys.platform == 'darwin' and 'Contents/MacOS' in base:
                   base = os.path.abspath(os.path.join(base, '../../..'))
                img_full_path = os.path.join(base, img_name)
            else:
                base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                img_full_path = os.path.join(base, img_name)
            
        try:
            if os.path.exists(img_full_path):
                img = Image.open(img_full_path)
                img = img.resize((200, 200), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                lbl_img = tk.Label(top, image=photo)
                lbl_img.image = photo # Keep reference
                lbl_img.pack(pady=20)
            else:
                tk.Label(top, text="[QR Code Not Found]", fg='gray').pack(pady=20)
        except Exception as e:
             tk.Label(top, text=f"[Error loading image: {e}]", fg='red').pack(pady=20)

        # 2. Text Info
        info_text = (
            "简易出货单生成小工具 v3.6\n"
            "Simple Delivery Order Generator\n\n"
            "作者 / Author: Sean\n"
            "联系方式 / Contact: fishis@126.com\n\n"
            "© 2026 All Rights Reserved"
        )
        tk.Label(top, text=info_text, font=('Arial', 10), justify='center').pack(pady=10)
        
        ttk.Button(top, text="确定 / OK", command=top.destroy).pack(pady=10)

    def import_products(self):
        # Show instruction first
        msg = (
            "请准备好 Excel 表格 (.xlsx)，确保包含以下列名（任选其一）：\n\n"
            "1. 商品名称 (必填)\n"
            "   (支持表头: 商品名称, Name, 品名)\n"
            "2. 规格型号\n"
            "   (支持表头: 规格型号, Model, 型号)\n"
            "3. 适用机型\n"
            "   (支持表头: 适用机型, Machine, 机型)\n"
            "4. 单位\n"
            "   (支持表头: 单位, Unit)\n"
            "5. 参考单价\n"
            "   (支持表头: 参考单价, 单价, Price)\n\n"
            "是否立即选择文件导入？"
        )
        if not messagebox.askokcancel("导入说明 / Import Help", msg):
            return

        filepath = filedialog.askopenfilename(
            title="选择产品资料表格 (Select Excel)",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not filepath: return
        
        try:
            count, error = self.product_manager.import_from_excel(filepath)
            if error:
                messagebox.showerror("Import Failed", f"Error: {error}")
            else:
                messagebox.showinfo("Success", f"成功导入 {count} 个产品！\nSuccessfully imported {count} items.")
                # Refresh combobox
                self.all_product_names = self.product_manager.get_product_names()
                self.cb_product['values'] = self.all_product_names
        except Exception as e:
             messagebox.showerror("Error", f"Import crashed: {e}")

    def setup_ui(self):
        # Notebook for Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True)

        try:
            # Configure ttk styles
            style.configure(".", font=('Microsoft YaHei', 10))
            style.configure("Treeview", font=('Microsoft YaHei', 10), rowheight=25)
            style.configure("Treeview.Heading", font=('Microsoft YaHei', 11, 'bold')) # Slightly larger for heading
            
            # Configure standard tk widgets via option database
            # Important: Font names with spaces must be wrapped in braces for Tcl
            self.root.option_add('*Font', '{Microsoft YaHei} 10')
        except Exception as e:
            print(f"Font config warning: {e}")
            pass

        # Tab 1: Generate Order
        self.tab_generate = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_generate, text='  开单 (Create Order)  ')
        self.setup_generate_tab()

        # Tab 2: History
        self.tab_history = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_history, text='  历史记录 (History)  ')
        self.setup_history_tab()

        # Tab 3: Order Summary (New)
        self.tab_summary = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_summary, text='  订单汇总 (Summary)  ')
        self.setup_summary_tab()

    def setup_generate_tab(self):
        # Top Frame 0: Seller Info (New)
        seller_frame = ttk.LabelFrame(self.tab_generate, text="销售方信息 / Seller Info", padding=10)
        seller_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(seller_frame, text="公司名称:").grid(row=0, column=0, sticky='w')
        
        # Load saved seller name and history
        saved_seller = self.order_generator.config.get("seller_name", "广州市 XX 办公设备有限公司")
        seller_history = self.order_generator.config.get("seller_history", [])
        if saved_seller and saved_seller not in seller_history:
            seller_history.insert(0, saved_seller)
            
        self.entry_seller_name = ttk.Combobox(seller_frame, values=seller_history, width=38)
        self.entry_seller_name.grid(row=0, column=1, padx=5, pady=5)
        self.entry_seller_name.set(saved_seller)

        # Top Frame 1: Customer Info
        top_frame = ttk.LabelFrame(self.tab_generate, text="客户信息 / Client Info", padding=10)
        top_frame.pack(fill='x', padx=10, pady=5)
        
        # Row 1
        ttk.Label(top_frame, text="客户名称:").grid(row=0, column=0, sticky='w')
        
        self.all_customers = self.history_manager.get_unique_customers()
        self.entry_customer = ttk.Combobox(top_frame, values=self.all_customers, width=28)
        self.entry_customer.grid(row=0, column=1, padx=5, pady=5)
        self.entry_customer.bind("<KeyRelease>", self.on_customer_search)
        
        ttk.Label(top_frame, text="客户地址:").grid(row=0, column=2, sticky='w')
        self.entry_address = ttk.Entry(top_frame, width=30)
        self.entry_address.grid(row=0, column=3, padx=5, pady=5)

        # Row 2 (Auto generated)
        ttk.Label(top_frame, text="单据编号:").grid(row=1, column=0, sticky='w')
        self.lbl_order_id = ttk.Label(top_frame, text="待生成...", font=('Arial', 10, 'bold'), foreground='blue')
        self.lbl_order_id.grid(row=1, column=1, sticky='w', padx=5)
        self.update_order_id_display() # Initial load

        ttk.Label(top_frame, text="单据日期:").grid(row=1, column=2, sticky='w')
        self.lbl_date = ttk.Label(top_frame, text=datetime.datetime.now().strftime("%Y-%m-%d"))
        self.lbl_date.grid(row=1, column=3, sticky='w', padx=5)
        
        # Middle Frame: Product Entry
        mid_frame = ttk.LabelFrame(self.tab_generate, text="产品录入 / Product Entry", padding=10)
        mid_frame.pack(fill='x', padx=10, pady=5)
        
        # Row 1
        ttk.Label(mid_frame, text="商品名称:").grid(row=0, column=0, sticky='w')
        
        # Searchable Combobox Logic
        self.all_product_names = self.product_manager.get_product_names()
        self.cb_product = ttk.Combobox(mid_frame, values=self.all_product_names, width=20)
        self.cb_product.grid(row=0, column=1, padx=5, pady=5)
        
        self.cb_product.bind("<<ComboboxSelected>>", self.on_product_select)
        self.cb_product.bind("<KeyRelease>", self.on_product_search)
        
        ttk.Label(mid_frame, text="规格型号:").grid(row=0, column=2, sticky='w')
        self.entry_model = ttk.Entry(mid_frame, width=15)
        self.entry_model.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(mid_frame, text="适用机型:").grid(row=0, column=4, sticky='w')
        self.entry_machine = ttk.Entry(mid_frame, width=15)
        self.entry_machine.grid(row=0, column=5, padx=5, pady=5)

        # Row 2
        ttk.Label(mid_frame, text="备注/品牌:").grid(row=1, column=0, sticky='w')
        brand_values = ["A格", "原装", "无"]
        self.cb_remark = ttk.Combobox(mid_frame, values=brand_values, width=15)
        self.cb_remark.grid(row=1, column=1, padx=5, pady=5)
        self.cb_remark.set("A格")

        ttk.Button(mid_frame, text="添加 / Add", command=self.add_item).grid(row=1, column=2, padx=10)
        ttk.Button(mid_frame, text="批量添加 / Batch Add...", command=self.open_batch_add).grid(row=1, column=3, padx=10, sticky='w')
        ttk.Button(mid_frame, text="产品导入", command=self.import_products).grid(row=1, column=4, padx=5, sticky='w')
        ttk.Button(mid_frame, text="产品导出", command=self.export_products).grid(row=1, column=5, padx=5, sticky='w')

        # List Frame: Order Items
        list_frame = ttk.Frame(self.tab_generate)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        cols = ("勾选", "序号", "商品名称", "规格型号", "单位", "数量", "单价", "合计", "备注")
        self.tree = ttk.Treeview(list_frame, columns=cols, show='headings', height=8)
        
        # Configure columns
        self.tree.heading("勾选", text="√")
        self.tree.column("勾选", width=40, anchor='center')
        
        self.tree.heading("序号", text="序号")
        self.tree.column("序号", width=50, anchor='center')
        
        for col in cols[2:]: # Skip check and index
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80, anchor='center')
        self.tree.column("商品名称", width=120)
        
        # Bind double click for editing (keep existing)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        # Bind single click for checkboxes
        self.tree.bind("<Button-1>", self.on_main_tree_click)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Delete Button & Total
        bottom_frame = ttk.Frame(self.tab_generate)
        bottom_frame.pack(fill='x', padx=10, pady=5)
        
        self.var_select_all_main = tk.IntVar()
        ttk.Checkbutton(bottom_frame, text="全选 / Select All", variable=self.var_select_all_main, command=self.toggle_select_all_main).pack(side='left', padx=(0, 10))
        
        ttk.Button(bottom_frame, text="删除选中 / Delete Selected", command=self.delete_item).pack(side='left')
        
        self.lbl_total = ttk.Label(bottom_frame, text="总金额: 0.00", font=('Arial', 12, 'bold'), foreground='red')
        self.lbl_total.pack(side='right', padx=20)

        # Action Buttons
        action_frame = ttk.Frame(self.tab_generate, padding=10)
        action_frame.pack(fill='x')
        
        ttk.Label(action_frame, text="制单人:").pack(side='left')
        self.entry_maker = ttk.Entry(action_frame, width=15)
        self.entry_maker.pack(side='left', padx=5)
        self.entry_maker.insert(0, "管理员")
        
        ttk.Button(action_frame, text="生成出货单 (PDF) / Generate", command=self.generate_order, width=25).pack(side='right')

    def setup_history_tab(self):
        # Filter Frame
        filter_frame = ttk.Frame(self.tab_history, padding=10)
        filter_frame.pack(fill='x')
        
        # Row 1: Filters
        
        ttk.Label(filter_frame, text="客户名称:").pack(side='left', padx=5)
        # Unique customers for dropdown
        self.all_customers_history = self.history_manager.get_unique_customers()
        self.h_customer = ttk.Combobox(filter_frame, values=self.all_customers_history, width=15)
        self.h_customer.pack(side='left', padx=5)

        ttk.Label(filter_frame, text="开始日期:").pack(side='left', padx=5)
        self.h_start_date = ttk.Entry(filter_frame, width=12)
        self.h_start_date.pack(side='left')
        self.h_start_date.insert(0, datetime.datetime.now().strftime("%Y-%m-01")) 

        ttk.Label(filter_frame, text="结束日期:").pack(side='left', padx=5)
        self.h_end_date = ttk.Entry(filter_frame, width=12)
        self.h_end_date.pack(side='left')
        self.h_end_date.insert(0, datetime.date.today().strftime("%Y-%m-%d"))

        ttk.Label(filter_frame, text="关键字:").pack(side='left', padx=5)
        self.h_keyword = ttk.Entry(filter_frame, width=15)
        self.h_keyword.pack(side='left')

        ttk.Button(filter_frame, text="查询 / Search", command=self.load_history).pack(side='left', padx=10)
        ttk.Button(filter_frame, text="导出汇总 / Export Summary", command=self.export_history_summary).pack(side='right', padx=10)

        # History Tree
        h_frame = ttk.Frame(self.tab_history)
        h_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Add Checkbox column
        cols = ("勾选", "单据编号", "日期", "客户名称", "总金额", "制单人")
        self.h_tree = ttk.Treeview(h_frame, columns=cols, show='headings')
        
        self.h_tree.heading("勾选", text="√")
        self.h_tree.column("勾选", width=40, anchor='center')
        
        for col in cols[1:]:
            self.h_tree.heading(col, text=col)
            self.h_tree.column(col, anchor='center')
            
        scrollbar = ttk.Scrollbar(h_frame, orient="vertical", command=self.h_tree.yview)
        self.h_tree.configure(yscrollcommand=scrollbar.set)
        
        self.h_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Bind click for checkboxes
        self.h_tree.bind("<Button-1>", self.on_history_click)

        # Bottom Action Bar (Select All, Delete)
        action_frame = ttk.Frame(self.tab_history, padding=10)
        action_frame.pack(fill='x')
        
        self.var_select_all_history = tk.IntVar()
        ttk.Checkbutton(action_frame, text="全选 / Select All", variable=self.var_select_all_history, command=self.toggle_select_all_history).pack(side='left', padx=10)
        
        ttk.Button(action_frame, text="删除选中记录 / Delete Selected", command=self.delete_selected_history).pack(side='left', padx=10)
        
        ttk.Button(action_frame, text="导出选中单据 / Export PDF", command=self.export_selected_pdfs).pack(side='right', padx=10)


    def update_order_id_display(self):
        # Show what the NEXT ID will be (without incrementing yet)
        # Note: Our logic increments on 'generate'. So we just show "Auto-generated" or current incomplete.
        # But user wants to see it maybe? 
        # For now, let's just generate it when button is clicked to avoid gaps if they don't save.
        # OR, we can just peek. But implementation was simple increment. 
        # Let's show "YKYYYYMMDD..." placeholder or generate it on app start?
        # Requirement: "Software opens OR click Add". 
        # I'll modify logic to separate "generate_new" from "get_next". 
        # Let's show "YKYYYYMMDD..." placeholder or generate one now and hold it.
        pass

    def on_product_select(self, event):
        name = self.cb_product.get()
        p = self.product_manager.get_product_by_name(name)
        if p:
            self.entry_model.delete(0, tk.END)
            self.entry_model.insert(0, p['model'])
            # Unit and Price are no longer displayed, so no need to fill them.
            # But we will use them in add_item.
            
            # Fill Machine Model if exists
            self.entry_machine.delete(0, tk.END)
            self.entry_machine.insert(0, p.get('machine_model', ''))

    def on_product_search(self, event):
        # ... (keep existing)
        typed = self.cb_product.get()
        if typed == '':
            data = self.all_product_names
        else:
            data = [item for item in self.all_product_names if typed.lower() in item.lower()]
        self.cb_product['values'] = data

    def on_customer_search(self, event):
        typed = self.entry_customer.get()
        if typed == '':
            data = self.all_customers
        else:
            data = [item for item in self.all_customers if typed.lower() in item.lower()]
        self.entry_customer['values'] = data

    def open_batch_add(self):
        # Open dialog
        dialog = BatchSelectionDialog(self.root, self.product_manager)
        self.root.wait_window(dialog)
        
        # Process returned items
        if dialog.selected_items:
            for item in dialog.selected_items:
                self.current_items.append(item)
            self.refresh_tree()

    def add_item(self):
        name = self.cb_product.get()
        if not name: return
        
        # Get Price/Unit from DB or Default
        p = self.product_manager.get_product_by_name(name)
        if p:
            price = p.get('price', 0.0)
            unit = p.get('unit', '')
        else:
            price = 0.0
            unit = ""
        
        qty = 1 # Default
        total = qty * price
        
        # Collect Machine Model
        machine = self.entry_machine.get().strip()
        
        # Save Product (Update/Add) if it's new or Machine Model changed
        # Since we removed Price/Unit inputs, we only update Name/Model/Machine.
        # We can preserve existing Price/Unit if updating.
        
        product_data = {
            "name": name,
            "model": self.entry_model.get(),
            "machine_model": machine
        }
        if p:
             product_data['price'] = p.get('price', 0.0)
             product_data['unit'] = p.get('unit', '')
        else:
             product_data['price'] = 0.0
             product_data['unit'] = ''
             
        self.product_manager.add_product(product_data)
        
        item = {
            "name": name,
            "model": self.entry_model.get(),
            "unit": unit,
            "qty": qty,
            "price": price,
            "total": total,
            "remark": self.cb_remark.get(),
            "_checked": False # Default unchecked
        }
        self.current_items.append(item)
        self.refresh_tree()
        
        # Refresh combobox
        self.all_product_names = self.product_manager.get_product_names()
        self.cb_product['values'] = self.all_product_names

    def on_tree_double_click(self, event):
        # Identify region
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        # Identify column
        col_id = self.tree.identify_column(event.x) 
        # col_id is like '#1', '#2'...
        
        # Map Display Columns to Data Keys and Column Names
        # Data Cols: ("序号", "商品名称", "规格型号", "单位", "数量", "单价", "合计", "备注")
        # Indices:      0         1           2         3       4       5       6      7
        # Display:     #1        #2          #3        #4      #5      #6      #7     #8
        
        col_map = {
            '#6': {'key': 'qty', 'col_name': '数量', 'type': int},
            '#7': {'key': 'price', 'col_name': '单价', 'type': float},
            '#9': {'key': 'remark', 'col_name': '备注', 'type': str}
        }
        
        if col_id not in col_map:
            return
            
        target = col_map[col_id]
        
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        
        # Get index in current_items list
        all_ids = self.tree.get_children()
        try:
            list_idx = all_ids.index(item_id)
        except ValueError:
            return
            
        self.edit_cell(item_id, list_idx, target['key'], target['col_name'], target['type'])

    def edit_cell(self, item_id, list_idx, key, col_name, data_type):
        # Get cell geometry using Column Name
        try:
            x, y, w, h = self.tree.bbox(item_id, column=col_name)
        except ValueError:
            # Fallback if bbox fails
            return
            
        # Create Entry
        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=w, height=h)
        
        # Set value
        current_val = self.current_items[list_idx][key]
        entry.insert(0, str(current_val))
        entry.select_range(0, tk.END)
        entry.focus_set() # Ensure focus
        
        def save(event=None):
            try:
                val_str = entry.get()
                if not val_str: 
                    # If empty, maybe reset or keep old?
                    # Let's keep old just to be safe, or allow 0
                    if data_type is str:
                         new_val = ""
                    else:
                         new_val = 0
                else:
                    new_val = data_type(val_str)
                
                self.current_items[list_idx][key] = new_val
                
                # Recalculate total if needed
                if key in ['qty', 'price']:
                    updated_item = self.current_items[list_idx]
                    updated_item['total'] = updated_item['qty'] * updated_item['price']
                    
                self.refresh_tree()
                entry.destroy()
            except ValueError:
                messagebox.showerror("Error", "Invalid Value")
                entry.focus_set()
        
        def cancel(event=None):
            entry.destroy()
            
        entry.bind("<Return>", save)
        entry.bind("<Escape>", cancel)
        entry.bind("<FocusOut>", lambda e: save()) # Save on click away? Or cancel? Usually save is better UX



    def refresh_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        total_amt = 0.0
        for idx, item in enumerate(self.current_items):
            # Ensure _checked key exists (migration for old items if any)
            if '_checked' not in item: item['_checked'] = False
            
            check_mark = "☑" if item['_checked'] else "☐"
            
            self.tree.insert("", "end", values=(
                check_mark,
                idx + 1,
                item['name'],
                item['model'],
                item['unit'],
                item['qty'],
                f"{item['price']:.2f}",
                f"{item['total']:.2f}",
                item['remark']
            ))
            total_amt += item['total']
            
        self.lbl_total.config(text=f"总金额: {total_amt:.2f}")

    def on_main_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        col = self.tree.identify_column(event.x)
        if col == '#1': # Checkbox column
            item_id = self.tree.identify_row(event.y)
            if not item_id: return
            
            # Find index
            all_ids = self.tree.get_children()
            try:
                list_idx = all_ids.index(item_id)
            except ValueError:
                return
            
            # Toggle
            self.current_items[list_idx]['_checked'] = not self.current_items[list_idx]['_checked']
            self.refresh_tree()

    def toggle_select_all_main(self):
        is_checked = self.var_select_all_main.get()
        for item in self.current_items:
            item['_checked'] = bool(is_checked)
        self.refresh_tree()

    def delete_item(self):
        # 1. Check for checked boxes
        to_delete_indices = [i for i, item in enumerate(self.current_items) if item.get('_checked', False)]
        
        # 2. If no boxes checked, check standard highlighting (Fallback)
        if not to_delete_indices:
            selected = self.tree.selection()
            if selected:
                # Find indices from tree items
                all_ids = self.tree.get_children()
                for sel_id in selected:
                    if sel_id in all_ids:
                        to_delete_indices.append(all_ids.index(sel_id))
        
        if not to_delete_indices:
             return
             
        # Delete from end to start
        for idx in sorted(set(to_delete_indices), reverse=True):
            if idx < len(self.current_items):
                del self.current_items[idx]
        
        # Uncheck "Select All" if empty or done (optional, but good UX)
        if not self.current_items:
            self.var_select_all_main.set(0)
            
        self.refresh_tree()

    def generate_order(self):
        customer = self.entry_customer.get()
        if not customer:
            messagebox.showwarning("Warning", "请输入客户名称 / Please enter Customer Name")
            return
        if not self.current_items:
            messagebox.showwarning("Warning", "请添加产品 / Please add items")
            return

        # 1. Generate ID
        order_id = self.order_generator.generate_new_number()
        self.lbl_order_id.config(text=order_id)
        
        # 2. Build Data
        order_data = {
            "order_id": order_id,
            "date": datetime.datetime.now().strftime("%Y-%m-%d"),
            "customer": customer,
            "address": self.entry_address.get(),
            "maker": self.entry_maker.get(),
            "items": self.current_items,
            "total": sum(x['total'] for x in self.current_items)
        }
        
        # 3. Save History
        self.history_manager.save_order(order_data)
        
        # 4. Export PDF
        default_filename = f"出货单_{order_id}_{customer}.pdf"
        
        # Get last save path
        initial_dir = self.order_generator.get_last_save_path()
        if not initial_dir:
            initial_dir = os.getcwd()

        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialdir=initial_dir,
            initialfile=default_filename,
            filetypes=[("PDF Files", "*.pdf")]
        )
        
        if filepath:
            try:
                # Save seller name preference and history
                seller_name = self.entry_seller_name.get()
                self._save_seller_info(seller_name)
                
                export_pdf(order_data, filepath, seller_info={'name': seller_name})
                # Save the directory for next time
                self.order_generator.set_last_save_path(filepath)
                
                messagebox.showinfo("Success", f"成功生成出货单!\nSaved to {filepath}")
                self.load_history() # Refresh history tab
                # Clear inputs?
                # self.current_items = []
                # self.refresh_tree()
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {e}")

    def load_history(self):
        for i in self.h_tree.get_children():
            self.h_tree.delete(i)
            
        start = self.h_start_date.get()
        end = self.h_end_date.get()
        kw = self.h_keyword.get()
        customer = self.h_customer.get() # Get customer filter
        
        # Update dropdown if needed (optional)
        self.all_customers_history = self.history_manager.get_unique_customers()
        self.h_customer['values'] = self.all_customers_history
        
        orders = self.history_manager.get_orders(start, end, kw, customer)
        
        # Store for checkbox state management
        self.history_displayed_items = []
        for o in orders:
            # Shallow copy to add temporary UI state
            o_ui = o.copy()
            o_ui['_checked'] = False
            self.history_displayed_items.append(o_ui)
            
        self.refresh_history_tree()
        
    def refresh_history_tree(self):
        for i in self.h_tree.get_children():
            self.h_tree.delete(i)
            
        for idx, o in enumerate(self.history_displayed_items):
            check_mark = "☑" if o.get('_checked') else "☐"
            
            self.h_tree.insert("", "end", values=(
                check_mark,
                o.get('order_id'),
                o.get('date'),
                o.get('customer'),
                f"{o.get('total', 0):.2f}",
                o.get('maker')
            ))
            
    def on_history_click(self, event):
        region = self.h_tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        col = self.h_tree.identify_column(event.x)
        if col == '#1': # Checkbox column
            item_id = self.h_tree.identify_row(event.y)
            if not item_id: return
            
            # Find index
            all_ids = self.h_tree.get_children()
            try:
                list_idx = all_ids.index(item_id)
            except ValueError:
                return
            
            # Toggle
            self.history_displayed_items[list_idx]['_checked'] = not self.history_displayed_items[list_idx]['_checked']
            self.refresh_history_tree()

    def toggle_select_all_history(self):
        is_checked = self.var_select_all_history.get()
        for item in self.history_displayed_items:
            item['_checked'] = bool(is_checked)
        self.refresh_history_tree()

    def delete_selected_history(self):
        to_delete = [item for item in self.history_displayed_items if item.get('_checked')]
        if not to_delete:
            messagebox.showinfo("Info", "请先勾选需要删除的记录 / Please select orders to delete")
            return
            
        count = len(to_delete)
        if not messagebox.askyesno("Confirm Delete", f"确定要删除选中的 {count} 条记录吗？\n此操作不可恢复！\n\nDelete {count} orders?"):
            return
            
        # Get IDs
        ids = [item['order_id'] for item in to_delete]
        
        # Delete from backend
        deleted_count = self.history_manager.delete_orders(ids)
        
        if deleted_count > 0:
            messagebox.showinfo("Success", f"成功删除 {deleted_count} 条记录 / Deleted successfully")
            self.load_history() # Refresh
            self.var_select_all_history.set(0) # Reset select all
        else:
            messagebox.showwarning("Warning", "删除失败 / Delete failed")

    def export_products(self):
        products = self.product_manager.products
        if not products:
            messagebox.showinfo("Info", "暂无产品数据 / No products to export")
            return
            
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="产品数据导出.xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if filepath:
            try:
                from openpyxl import Workbook
                wb = Workbook()
                ws = wb.active
                ws.title = "Products"
                
                # Headers
                headers = ["商品名称", "规格型号", "适用机型", "单位", "参考单价"]
                ws.append(headers)
                
                for p in products:
                    row = [
                        p.get('name', ''),
                        p.get('model', ''),
                        p.get('machine_model', ''),
                        p.get('unit', ''),
                        p.get('price', 0)
                    ]
                    ws.append(row)
                    
                wb.save(filepath)
                messagebox.showinfo("Success", f"成功导出 {len(products)} 个产品!\nSaved to {filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"导出失败: {e}")

    def export_history_summary(self):
        orders = self.history_manager.get_orders(self.h_start_date.get(), self.h_end_date.get(), self.h_keyword.get(), self.h_customer.get())
        if not orders:
            messagebox.showinfo("Info", "No orders to export")
            return
            
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"销售汇总_{datetime.date.today()}.xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if filepath:
            self.history_manager.export_summary_to_excel(orders, filepath)
            messagebox.showinfo("Success", "Summary Exported")

    def export_selected_pdfs(self):
        to_export = [item for item in self.history_displayed_items if item.get('_checked')]
        if not to_export:
            messagebox.showinfo("Info", "请先勾选需要导出的单据 / Please select orders to export")
            return
            
        count = len(to_export)
        
        # Ask for directory
        directory = filedialog.askdirectory(title="选择导出文件夹 / Select Export Folder")
        if not directory:
            return
            
        success_count = 0
        try:
            for item in to_export:
                # Remove UI state if passed directly, but export_pdf only uses dict keys anyway.
                # Construct filename: OrderID_Customer.pdf
                safe_name = item.get('order_id', 'Unknown').replace('/', '_') # sanitize slashes
                safe_cust = item.get('customer', 'Client').replace('/', '_')
                filename = f"{safe_name}_{safe_cust}.pdf"
                filepath = os.path.join(directory, filename)
                
                filename = f"{safe_name}_{safe_cust}.pdf"
                filepath = os.path.join(directory, filename)
                
                filename = f"{safe_name}_{safe_cust}.pdf"
                filepath = os.path.join(directory, filename)
                
                seller_name = self.entry_seller_name.get() # Ensure consistent using main tab entry
                self._save_seller_info(seller_name)
                
                export_pdf(item, filepath, report_type='delivery', seller_info={'name': seller_name})
                success_count += 1
                
            messagebox.showinfo("Success", f"成功导出 {success_count} 个文件!\nSaved to {directory}")
            
        except Exception as e:
            messagebox.showerror("Error", f"导出失败 / Export failed: {e}")

    def setup_summary_tab(self):
        # Frame for controls
        frame = ttk.Frame(self.tab_summary, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Row 1: Customer Selection (Mandatory)
        row1 = ttk.Frame(frame)
        row1.pack(fill='x', pady=10)
        ttk.Label(row1, text="选择客户 (Customer):", width=20, anchor='e').pack(side='left', padx=5)
        self.all_customers_summary = self.history_manager.get_unique_customers()
        self.cb_summary_customer = ttk.Combobox(row1, values=self.all_customers_summary, width=25)
        self.cb_summary_customer.pack(side='left', padx=5)
        
        # Row 2: Date Range (Filtering)
        row2 = ttk.Frame(frame)
        row2.pack(fill='x', pady=10)
        ttk.Label(row2, text="统计周期 (Date Range):", width=20, anchor='e').pack(side='left', padx=5)
        
        self.sum_start_date = ttk.Entry(row2, width=12)
        self.sum_start_date.pack(side='left', padx=5)
        self.sum_start_date.insert(0, datetime.datetime.now().strftime("%Y-%m-01"))
        
        ttk.Label(row2, text="至").pack(side='left')
        
        self.sum_end_date = ttk.Entry(row2, width=12)
        self.sum_end_date.pack(side='left', padx=5)
        self.sum_end_date.insert(0, datetime.date.today().strftime("%Y-%m-%d"))

        # Row 3: Display Date (Manual Input)
        row3 = ttk.Frame(frame)
        row3.pack(fill='x', pady=10)
        ttk.Label(row3, text="报表显示日期 (Display Date):", width=20, anchor='e').pack(side='left', padx=5)
        self.entry_display_date = ttk.Entry(row3, width=30)
        self.entry_display_date.pack(side='left', padx=5)
        # Default value? "YYYY-MM-DD 至 YYYY-MM-DD"
        s = self.sum_start_date.get()
        e = self.sum_end_date.get()
        self.entry_display_date.insert(0, f"{s} 至 {e}")
        
        # Row 4: Export Mode (Detail vs Merged)
        row4 = ttk.Frame(frame)
        row4.pack(fill='x', pady=10)
        ttk.Label(row4, text="导出方式 (Export Mode):", width=20, anchor='e').pack(side='left', padx=5)
        
        self.var_export_mode = tk.StringVar(value="detail")
        ttk.Radiobutton(row4, text="按开单明细 (Detail)", variable=self.var_export_mode, value="detail").pack(side='left', padx=10)
        ttk.Radiobutton(row4, text="按产品合并 (Merged)", variable=self.var_export_mode, value="merged").pack(side='left', padx=10)

        # Row 5: file Format (PDF vs Excel)
        row5 = ttk.Frame(frame)
        row5.pack(fill='x', pady=10)
        ttk.Label(row5, text="导出格式 (Format):", width=20, anchor='e').pack(side='left', padx=5)
        
        self.var_export_format = tk.StringVar(value="pdf")
        ttk.Radiobutton(row5, text="PDF 文件", variable=self.var_export_format, value="pdf").pack(side='left', padx=10)
        ttk.Radiobutton(row5, text="Excel 表格", variable=self.var_export_format, value="excel").pack(side='left', padx=10)

        # Row 6: Generate Button
        row6 = ttk.Frame(frame)
        row6.pack(fill='x', pady=20)
        ttk.Button(row6, text="生成对账单 / Generate", command=self.generate_summary_statement).pack(anchor='center')
        
    def generate_summary_statement(self):
        customer = self.cb_summary_customer.get()
        if not customer:
            messagebox.showwarning("Warning", "请选择客户 / Please select a customer")
            return
            
        start = self.sum_start_date.get()
        end = self.sum_end_date.get()
        
        # Options
        mode = self.var_export_mode.get()
        fmt = self.var_export_format.get()
        
        # fetch orders
        orders = self.history_manager.get_orders(start_date=start, end_date=end, customer_name=customer)
        
        if not orders:
            messagebox.showinfo("Info", "该时间段内无此客户订单 / No orders found")
            return
        
        # Aggregate items
        all_items = []
        for o in orders:
            for item in o.get('items', []):
                 all_items.append(item)
                 
        if not all_items:
             messagebox.showinfo("Info", "订单中无商品 / No items found")
             return
        
        final_items = []
        if mode == 'merged':
            # Merge logic: Key = (Name, Model, Price)
            # What about Unit? Usually consistent with Name/Model.
            merged = {}
            for item in all_items:
                # Key
                key = (item.get('name'), item.get('model'), item.get('price'))
                if key not in merged:
                    merged[key] = {
                        'name': item.get('name'),
                        'model': item.get('model'),
                        'price': float(item.get('price', 0)),
                        'unit': item.get('unit'),
                        'qty': 0,
                        'total': 0.0,
                        'remark': item.get('remark') # Keep first remark? "A格"
                    }
                
                # Add
                qty = float(item.get('qty', 0))
                merged[key]['qty'] += qty
                merged[key]['total'] += (qty * float(item.get('price', 0)))
                # Update total
            
            final_items = list(merged.values())
        else:
            final_items = all_items

        # Build Data object
        summary_data = {
            "customer": customer,
            "date": self.entry_display_date.get(),
            "address": "", # Ignored
            "order_id": "", # Ignored
            "items": final_items,
            "maker": "管理员", 
        }
        
        # Export
        filename = f"对账单_{customer}_{start}_{end}"
        if fmt == 'excel':
            ext = ".xlsx"
            ftypes = [("Excel Files", "*.xlsx")]
        else:
            ext = ".pdf"
            ftypes = [("PDF Files", "*.pdf")]
            
        filepath = filedialog.asksaveasfilename(
            defaultextension=ext,
            initialfile=filename + ext,
            filetypes=ftypes
        )
        
        if filepath:
            try:
                seller_name = self.entry_seller_name.get() # Use main tab's entry
                self._save_seller_info(seller_name)
                
                if fmt == 'excel':
                    export_to_excel(summary_data, filepath, report_type='summary', seller_info={'name': seller_name})
                else:
                    export_pdf(summary_data, filepath, report_type='summary', seller_info={'name': seller_name})
                    
                messagebox.showinfo("Success", f"对账单已生成!\nMode: {mode}\nSaved to {filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed: {e}")

    def _save_seller_info(self, seller_name):
        """Helper to save current seller name and update history list in config."""
        if not seller_name: return
        
        # Update current default
        self.order_generator.config["seller_name"] = seller_name
        
        # Update history list
        history = self.order_generator.config.get("seller_history", [])
        if seller_name not in history:
            history.insert(0, seller_name) # Add to top
            # Limit history size? e.g. 10
            if len(history) > 10:
                history = history[:10]
            self.order_generator.config["seller_history"] = history
            
            # Update UI dropdown values immediately
            self.entry_seller_name['values'] = history
            
        self.order_generator.save_config()

class BatchSelectionDialog(tk.Toplevel):
    def __init__(self, parent, product_manager):
        super().__init__(parent)
        self.title("批量添加产品 / Batch Add Products")
        self.geometry("800x600")
        self.product_manager = product_manager
        self.selected_items = [] # List of dicts to return
        
        self.all_products = self.product_manager.products
        self.filtered_products = list(self.all_products)
        
        self.setup_ui()
        
    def setup_ui(self):
        # Custom style for larger rows and font
        style = ttk.Style()
        style.configure("Batch.Treeview", font=('Arial', 12), rowheight=30)
        style.configure("Batch.Treeview.Heading", font=('Arial', 11, 'bold'))

        # Filter Frame
        filter_frame = ttk.LabelFrame(self, text="筛选 / Filter", padding=10)
        filter_frame.pack(fill='x', padx=10, pady=5)
        
        # Machine Model Filter (Combobox)
        ttk.Label(filter_frame, text="适用机型 / Machine Model:").pack(side='left')
        
        # Get unique models
        models = sorted(list(set(p.get('machine_model', '') for p in self.all_products if p.get('machine_model'))))
        models.insert(0, "") # Empty option
        
        self.cb_machine = ttk.Combobox(filter_frame, values=models, width=20)
        self.cb_machine.pack(side='left', padx=5)
        self.cb_machine.bind("<<ComboboxSelected>>", self.on_search)
        
        # Keyword Filter
        ttk.Label(filter_frame, text="关键字 / Keyword:").pack(side='left', padx=(15, 0))
        self.entry_keyword = ttk.Entry(filter_frame, width=20)
        self.entry_keyword.pack(side='left', padx=5)
        self.entry_keyword.bind("<KeyRelease>", self.on_search)
        
        ttk.Button(filter_frame, text="重置 / Reset", command=self.reset_filters).pack(side='left', padx=10)
        
        # List Frame
        list_frame = ttk.Frame(self)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        cols = ("勾选", "商品名称", "规格型号", "适用机型", "单位", "单价")
        self.tree = ttk.Treeview(list_frame, columns=cols, show='headings', selectmode='extended', style="Batch.Treeview")
        
        # Define columns
        self.tree.heading("勾选", text="状态")
        self.tree.column("勾选", width=60, anchor='center')
        
        self.tree.heading("商品名称", text="商品名称")
        self.tree.column("商品名称", width=200)
        
        self.tree.heading("规格型号", text="规格型号")
        self.tree.column("规格型号", width=150)
        
        self.tree.heading("适用机型", text="适用机型")
        self.tree.column("适用机型", width=150)
        
        self.tree.heading("单位", text="单位")
        self.tree.column("单位", width=60, anchor='center')
        
        self.tree.heading("单价", text="参考单价")
        self.tree.column("单价", width=100, anchor='e')
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Configure tags for alternating colors
        self.tree.tag_configure('odd', background='#ffffff')
        self.tree.tag_configure('even', background='#e0f2fe') # Light blue
        
        # Bottom Frame
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.pack(fill='x')
        
        self.var_select_all = tk.IntVar()
        # Align with "Status" column (centered in 60px width) -> approx 20-30px offset
        ttk.Checkbutton(btn_frame, text="全选 / Select All", variable=self.var_select_all, command=self.toggle_select_all).pack(side='left', padx=(25, 0))
        
        ttk.Button(btn_frame, text="添加选中 / Add Selected", command=self.add_selected, width=20).pack(side='right')
        
        # Bind click for checkbox toggle
        self.tree.bind("<Button-1>", self.on_click)
        
        self.refresh_list()

    def toggle_select_all(self):
        is_checked = self.var_select_all.get()
        symbol = "☑" if is_checked else "☐"
        
        for item_id in self.tree.get_children():
            self.tree.set(item_id, column="勾选", value=symbol)
        
    def reset_filters(self):
        self.cb_machine.set('')
        self.entry_keyword.delete(0, tk.END)
        self.on_search()

    def on_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        col = self.tree.identify_column(event.x)
        if col == '#1': # Checkbox column
            item_id = self.tree.identify_row(event.y)
            if not item_id: return
            
            # Toggle value
            current_vals = self.tree.item(item_id)['values']
            # values is returned as a tuple/list. We need to modify it.
            # Note: tree.item(id)['values'] returns lists as strings often in older tk, but tuple in newer?
            # It usually returns a list or tuple.
            
            current_status = current_vals[0]
            new_status = "☑" if current_status == "☐" else "☐"
            
            # Update tree
            # We must pass ALL values to update, or use set for specific column
            self.tree.set(item_id, column="勾选", value=new_status)
    
    def on_search(self, event=None):
        machine = self.cb_machine.get().lower()
        keyword = self.entry_keyword.get().lower()
        
        self.filtered_products = []
        for p in self.all_products:
            p_machine = p.get('machine_model', '').lower()
            p_name = p.get('name', '').lower()
            p_model = p.get('model', '').lower()
            
            match_machine = (machine == '') or (machine in p_machine)
            match_keyword = (keyword == '') or (keyword in p_name) or (keyword in p_model)
            
            if match_machine and match_keyword:
                self.filtered_products.append(p)
                
        self.refresh_list()
        
    def refresh_list(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
            
        for i, p in enumerate(self.filtered_products):
            tag = 'even' if i % 2 == 0 else 'odd'
            self.tree.insert("", "end", values=(
                "☐", # Checkbox unchecked
                p.get('name', ''),
                p.get('model', ''),
                p.get('machine_model', ''),
                p.get('unit', ''),
                f"{p.get('price', 0):.2f}"
            ), tags=(tag,))

    def add_selected(self):
        # Scan all items for Checked status
        selected_rows = []
        for iid in self.tree.get_children():
            vals = self.tree.item(iid)['values']
            if vals[0] == "☑":
                selected_rows.append(vals)
                
        if not selected_rows:
            messagebox.showwarning("Warning", "请先勾选产品 / Please check products")
            return
            
        for vals in selected_rows:
            # values: 0=Check, 1=Name, 2=Model, 3=Machine, 4=Unit, 5=Price
            
            item = {
                "name": vals[1],
                "model": vals[2],
                # "machine_model": vals[3], 
                "unit": vals[4],
                "price": float(vals[5]),
                "qty": 1, 
                "total": float(vals[5]) * 1,
                "remark": "A格" 
            }
            self.selected_items.append(item)
            
        self.destroy()
