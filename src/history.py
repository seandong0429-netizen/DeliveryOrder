import json
import os
import datetime
import sys
from openpyxl import Workbook

if getattr(sys, 'frozen', False):
    base = os.path.dirname(sys.executable)
    if sys.platform == 'darwin' and 'Contents/MacOS' in base:
        possible_base = os.path.abspath(os.path.join(base, '../../..'))
        if os.path.exists(os.path.join(possible_base, 'orders.json')): # Check orders.json or common dir
            BASE_DIR = possible_base
        else:
             # Fallback: check if we set it in logic? Use the same logic
             if os.path.exists(os.path.join(possible_base, 'products.json')):
                 BASE_DIR = possible_base
             else:
                 BASE_DIR = base
    else:
        BASE_DIR = base
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

ORDERS_FILE = os.path.join(BASE_DIR, 'orders.json')

class HistoryManager:
    def __init__(self):
        self.orders = []
        self.load_orders()

    def load_orders(self):
        if os.path.exists(ORDERS_FILE):
            try:
                with open(ORDERS_FILE, 'r', encoding='utf-8') as f:
                    self.orders = json.load(f)
            except Exception as e:
                print(f"Error loading orders: {e}")
                self.orders = []
        else:
            self.orders = []

    def save_order(self, order_data):
        # unique check?
        self.orders.append(order_data)
        self._persist()

    def _persist(self):
        try:
            with open(ORDERS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.orders, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving orders: {e}")

    def get_orders(self, start_date=None, end_date=None, keyword="", customer_name=""):
        # dates are YYYY-MM-DD strings
        filtered = []
        for order in self.orders:
            o_date = order.get('date', '')
            
            # Date filter
            if start_date and o_date < start_date:
                continue
            if end_date and o_date > end_date:
                continue
                
            # Customer Name filter
            if customer_name:
                # Exact match required
                if customer_name.lower() != order.get('customer', '').lower():
                    continue

            # Keyword filter (search in customer name, order id, or remark)
            if keyword:
                kw = keyword.lower()
                found = False
                if kw in order.get('order_id', '').lower(): found = True
                if kw in order.get('customer', '').lower(): found = True
                # Check items for remarks or names
                for item in order.get('items', []):
                     if kw in item.get('name', '').lower(): found = True
                     if kw in item.get('remark', '').lower(): found = True
                if not found:
                    continue
            
            filtered.append(order)
            
        # Sort by date desc, then ID desc
        filtered.sort(key=lambda x: (x.get('date', ''), x.get('order_id', '')), reverse=True)
        return filtered

    def delete_orders(self, order_ids):
        """Delete orders by a list of order_ids."""
        initial_count = len(self.orders)
        self.orders = [o for o in self.orders if o.get('order_id') not in order_ids]
        if len(self.orders) < initial_count:
            self._persist()
        return initial_count - len(self.orders)

    def export_summary_to_excel(self, orders, filename):
        wb = Workbook()
        ws = wb.active
        ws.title = "销售汇总"
        
        headers = ["单据编号", "日期", "客户名称", "商品名称", "规格型号", "数量", "单位", "单价", "金额", "备注", "制单人"]
        ws.append(headers)
        
        for order in orders:
            base_info = [
                order.get('order_id'),
                order.get('date'),
                order.get('customer'),
            ]
            maker = order.get('maker', '')
            
            for item in order.get('items', []):
                row = base_info + [
                    item.get('name'),
                    item.get('model'),
                    item.get('qty'),
                    item.get('unit'),
                    item.get('price'),
                    item.get('total'),
                    item.get('remark'),
                    maker
                ]
                ws.append(row)
                
        wb.save(filename)

    def get_unique_customers(self):
        """Return a sorted list of unique customer names from history."""
        customers = set()
        for order in self.orders:
            c = order.get('customer')
            if c:
                customers.add(c)
        return sorted(list(customers))
