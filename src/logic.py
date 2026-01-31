import json
import os
import datetime
import sys

# Paths to data files
if getattr(sys, 'frozen', False):
    # If frozen (EXE), look in the same directory as the executable
    base = os.path.dirname(sys.executable)
    # Check if we are in a macOS Bundle ( Contents/MacOS )
    if sys.platform == 'darwin' and 'Contents/MacOS' in base:
        # Move up 3 levels to get out of .app
        # .../MyApp.app/Contents/MacOS -> .../MyApp.app/Contents -> .../MyApp.app -> .../
        possible_base = os.path.abspath(os.path.join(base, '../../..'))
        # If products.json exists there, use it. Otherwise default to base (inside app? or maybe we prefer external)
        # We want to favor the external one so user can edit it.
        if os.path.exists(os.path.join(possible_base, 'products.json')) or os.path.exists(os.path.join(possible_base, 'config.json')):
            BASE_DIR = possible_base
        else:
            BASE_DIR = base
    else:
        BASE_DIR = base
else:
    # If running from source, look in parent of src
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PRODUCTS_FILE = os.path.join(BASE_DIR, 'products.json')
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')

class ProductManager:
    def __init__(self):
        self.products = []
        self.load_products()

    def load_products(self):
        if os.path.exists(PRODUCTS_FILE):
            try:
                with open(PRODUCTS_FILE, 'r', encoding='utf-8') as f:
                    self.products = json.load(f)
            except Exception as e:
                print(f"Error loading products: {e}")
                self.products = []
        else:
            self.products = []

    def get_product_names(self):
        return [p['name'] for p in self.products]

    def get_product_by_name(self, name):
        for p in self.products:
            if p['name'] == name:
                return p
        return None

    def add_product(self, product_data):
        # Check if exists
        for i, p in enumerate(self.products):
            if p['name'] == product_data['name']:
                self.products[i] = product_data # Update
                self.save_products()
                return
        self.products.append(product_data)
        self.save_products()

    def save_products(self):
        try:
            with open(PRODUCTS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.products, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving products: {e}")

    def import_from_excel(self, filepath):
        """Import products from an Excel file."""
        try:
            import openpyxl
        except ImportError:
            raise ImportError("openpyxl module is required for Excel import. Please install it.")

        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            sheet = wb.active
            
            # Identify columns
            headers = {}
            for cell in sheet[1]: # First row
                val = str(cell.value).strip().lower() if cell.value else ""
                if val in ['商品名称', 'name', '品名']: headers['name'] = cell.column - 1
                elif val in ['规格型号', 'model', '型号']: headers['model'] = cell.column - 1
                elif val in ['适用机型', 'machine', 'machine_model', '机型']: headers['machine_model'] = cell.column - 1
                elif val in ['单位', 'unit']: headers['unit'] = cell.column - 1
                elif val in ['参考单价', '单价', 'price']: headers['price'] = cell.column - 1
            
            if 'name' not in headers:
                return 0, "Excel must contain a 'Name' or '商品名称' column."

            count = 0
            for row in sheet.iter_rows(min_row=2, values_only=True):
                name = row[headers['name']]
                if not name: continue # Skip empty names
                
                # Get other values safely
                model = row[headers['model']] if 'model' in headers else ""
                machine = row[headers['machine_model']] if 'machine_model' in headers else ""
                unit = row[headers['unit']] if 'unit' in headers else ""
                price_val = row[headers['price']] if 'price' in headers else 0
                
                try:
                    price = float(price_val) if price_val else 0.0
                except:
                    price = 0.0
                
                product_data = {
                    "name": str(name).strip(),
                    "model": str(model).strip() if model else "",
                    "machine_model": str(machine).strip() if machine else "",
                    "unit": str(unit).strip() if unit else "",
                    "price": price
                }
                
                self.add_product(product_data)
                count += 1
                
            return count, None
            
        except Exception as e:
            return 0, str(e)


class OrderNumberGenerator:
    def __init__(self):
        self.load_config()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            except:
                self.config = {"last_date": "", "sequence": 0}
        else:
            self.config = {"last_date": "", "sequence": 0}

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")

    def get_last_save_path(self):
        return self.config.get("last_save_path", "")

    def set_last_save_path(self, path):
        if path:
            self.config["last_save_path"] = os.path.dirname(path)
            self.save_config()

    def generate_new_number(self):
        today_str = datetime.datetime.now().strftime("%Y%m%d")
        
        if self.config.get("last_date") == today_str:
            self.config["sequence"] += 1
        else:
            self.config["last_date"] = today_str
            self.config["sequence"] = 1
        
        self.save_config()
        
        seq_str = f"{self.config['sequence']:03d}"
        return f"YK{today_str}{seq_str}"
    
    def get_current_number(self):
        # Preview what the next number would be without incrementing, 
        # OR just return the last generated one?
        # Requirement: "Software opens OR click Add, generate NEW number". 
        # So we likely usually want to generate one.
        # But if we want to display it before saving, we might need a peek method.
        # For now, let's stick to generate_new_number being called when needed.
        pass
