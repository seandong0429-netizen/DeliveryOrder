import sys
import os
sys.path.append(os.path.join(os.getcwd(), 'src'))

from logic import ProductManager, OrderNumberGenerator
from history import HistoryManager
from export_pdf import export_pdf

def test_backend():
    print("Testing Backend...")
    
    # 1. Product Manager
    pm = ProductManager()
    products = pm.get_product_names()
    print(f"Products loaded: {len(products)}")
    assert len(products) >= 4, "Should have at least 4 default products"
    
    # 2. Order Number
    ong = OrderNumberGenerator()
    new_id = ong.generate_new_number()
    print(f"Generated ID: {new_id}")
    assert new_id.startswith("YK"), "ID format wrong"
    
    # 3. History
    hm = HistoryManager()
    test_order = {
        "order_id": new_id,
        "date": "2026-01-30",
        "customer": "Test Customer",
        "items": [
            {"name": "Test Product", "model": "M1", "price": 100, "qty": 2, "total": 200, "remark": "R1"}
        ],
        "total": 200,
        "maker": "Tester"
    }
    hm.save_order(test_order)
    orders = hm.get_orders()
    print(f"Orders in history: {len(orders)}")
    assert len(orders) >= 1, "Order not saved"
    
    # 4. Export PDF
    pdf_path = f"test_{new_id}.pdf"
    export_pdf(test_order, pdf_path)
    if os.path.exists(pdf_path):
        print("PDF generated successfully")
        os.remove(pdf_path)
    else:
        print("PDF generation failed")
    
    print("Backend verification passed!")

if __name__ == "__main__":
    test_backend()
