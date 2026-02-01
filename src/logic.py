import json
import os
import datetime
import sys
import debug_utils

# Paths to data files
if getattr(sys, 'frozen', False):
    # Frozen (EXE/APP)
    base = os.path.dirname(sys.executable)
    if sys.platform == 'darwin' and 'Contents/MacOS' in base:
        # Standard Mac App Bundle Structure: Not sandboxed, but path is inside bundle
        # We want to break out of the bundle to the folder CONTAINING the .app
        # base is .../DeliveryOrderGen.app/Contents/MacOS
        # We need .../ (The dir containing DeliveryOrderGen.app)
        
        # ../.. -> Contents
        # ../../.. -> Directory containing .app
        
        app_containing_dir = os.path.abspath(os.path.join(base, '../../..'))
        debug_utils.log(f"Mac App Bundle detected. Containing Dir: {app_containing_dir}")
        
        # Check permissions: Can we write to the app's folder?
        # If user puts app in /Applications, this is usually False.
        # If on Desktop/Downloads, this is usually True.
        
        if os.access(app_containing_dir, os.W_OK):
             # Special Cleanliness Rule for /Applications root
             if app_containing_dir.rstrip(os.sep) == '/Applications':
                 BASE_DIR = os.path.join(app_containing_dir, "DeliveryOrderData")
                 if not os.path.exists(BASE_DIR):
                     try:
                         os.makedirs(BASE_DIR)
                         debug_utils.log(f"Created data subfolder in Applications: {BASE_DIR}")
                     except Exception as e:
                         # Fallback if we thought we could write but couldn't make dir
                         BASE_DIR = os.path.expanduser("~/Documents/DeliveryOrderData")
             else:
                 # Normal Portable Mode (Desktop, etc): Flat structure
                 BASE_DIR = app_containing_dir

             debug_utils.log(f"App dir is writable. Using Path: {BASE_DIR}")
        else:
             debug_utils.log(f"App dir {app_containing_dir} is NOT writable. Falling back to Documents.")
             # Installed Mode: Use ~/Documents/DeliveryOrderData
             user_dir = os.path.expanduser("~/Documents/DeliveryOrderData")
             if not os.path.exists(user_dir):
                 try:
                    os.makedirs(user_dir)
                 except Exception as e:
                    debug_utils.log(f"Failed to create user_dir {user_dir}: {e}")
             BASE_DIR = user_dir
             debug_utils.log(f"Using Installed Mode: {BASE_DIR}")
    else:
        # Windows/Linux EXE
        BASE_DIR = base
        debug_utils.log(f"Using Standard Frozen Mode: {BASE_DIR}")
else:
    # Source Code
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    debug_utils.log(f"Using Source Code Mode: {BASE_DIR}")

PRODUCTS_FILE = os.path.join(BASE_DIR, 'products.json')
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')

class ProductManager:
    def __init__(self):
        self.products = []
        self.load_products()

    def load_products(self):
        debug_utils.log(f"Loading products from {PRODUCTS_FILE}")
        if os.path.exists(PRODUCTS_FILE):
            try:
                with open(PRODUCTS_FILE, 'r', encoding='utf-8') as f:
                    self.products = json.load(f)
            except Exception as e:
                debug_utils.log(f"Error loading products: {e}")
                self.products = []
        else:
            debug_utils.log("No products file found, starting new.")
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
        
    def batch_add_products(self, product_list):
        """Add multiple products and save only once."""
        try:
            updates = 0
            adds = 0
            for new_p in product_list:
                found = False
                for i, p in enumerate(self.products):
                    if p['name'] == new_p['name']:
                        self.products[i] = new_p
                        found = True
                        updates += 1
                        break
                if not found:
                    self.products.append(new_p)
                    adds += 1
            
            self.save_products()
            debug_utils.log(f"Batch add finished: {adds} added, {updates} updated.")
        except Exception as e:
             debug_utils.log(f"Batch add failed: {e}")
             raise e

    def save_products(self):
        try:
            with open(PRODUCTS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.products, f, indent=2, ensure_ascii=False)
            debug_utils.log("Products saved successfully.")
        except Exception as e:
            debug_utils.log(f"Error saving products to {PRODUCTS_FILE}: {e}")
            # Re-raise so UI can handle/show error if needed, or at least we know it failed
            raise e

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

            items_to_add = []
            
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
                
                items_to_add.append(product_data)

            if items_to_add:
                self.batch_add_products(items_to_add)
                
            return len(items_to_add), None
            
        except Exception as e:
            return 0, str(e)


class OrderNumberGenerator:
    def __init__(self):
        self.load_config()

    def load_config(self):
        debug_utils.log(f"Loading config from {CONFIG_FILE}")
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                debug_utils.log("Config loaded successfully.")
            except Exception as e:
                debug_utils.log(f"Error loading config: {e}")
                self.config = {"last_date": "", "sequence": 0}
        else:
            debug_utils.log("Config file not found, creating default.")
            self.config = {"last_date": "", "sequence": 0}

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
            debug_utils.log("Config saved successfully.")
        except Exception as e:
            debug_utils.log(f"Error saving config: {e}")

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
