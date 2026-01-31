from src.export_pdf import export_pdf
import os

# Mock Data: 7 items + 1 ghost item? Or just 7 items.
# User log said 7 items.
# Let's try 6 items (since Total was 1170 = 6*195).
# Maybe log said 7 because of something else?
# Let's try both cases.

def test_export():
    items_6 = []
    for i in range(6):
        items_6.append({
            "name": f"Item {i+1}",
            "qty": 1,
            "price": 100,
            "total": 100
        })
        
    items_7 = items_6 + [{"name": "Item 7", "qty": 1, "price": 100, "total": 100}]
    
    order_data_6 = {
        "customer": "TestCust",
        "date": "2026-01-30",
        "items": items_6,
        "maker": "Admin",
        "order_id": "YK006"
    }

    order_data_7 = {
        "customer": "TestCust",
        "date": "2026-01-30",
        "items": items_7,
        "maker": "Admin",
        "order_id": "YK007"
    }
    
    print("--- Testing 6 Items ---")
    export_pdf(order_data_6, "test_6.pdf")
    
    print("\n--- Testing 7 Items ---")
    export_pdf(order_data_7, "test_7.pdf")

if __name__ == "__main__":
    test_export()
