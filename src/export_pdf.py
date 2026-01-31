from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import sys

def register_fonts():
    """Register Chinese fonts based on OS."""
    font_path = None
    # Common paths for Microsoft YaHei on Mac (if Office is installed) or Windows
    font_paths = [
        # Windows
        ('c:/windows/fonts/msyh.ttf', 'c:/windows/fonts/msyhbd.ttf'),
        # Mac Office / User Installed
        (os.path.expanduser('~/Library/Fonts/Microsoft YaHei.ttf'), os.path.expanduser('~/Library/Fonts/Microsoft YaHei Bold.ttf')),
        ('/Library/Fonts/Microsoft YaHei.ttf', '/Library/Fonts/Microsoft YaHei Bold.ttf'),
        # Fallback to PingFang (Mac Standard)
        ('/System/Library/Fonts/PingFang.ttc', '/System/Library/Fonts/PingFang.ttc'),
        ('/System/Library/Fonts/STHeiti Light.ttc', '/System/Library/Fonts/STHeiti Medium.ttc'),
    ]

    regular_font = 'Helvetica'
    bold_font = 'Helvetica-Bold'
    
    for reg_path, bold_path in font_paths:
        if os.path.exists(reg_path):
            try:
                # Register Regular
                pdfmetrics.registerFont(TTFont('ChineseFont', reg_path))
                regular_font = 'ChineseFont'
                
                # Register Bold (Try to find explicit bold file, or reuse regular if desperate)
                if os.path.exists(bold_path):
                    pdfmetrics.registerFont(TTFont('ChineseFont-Bold', bold_path))
                    bold_font = 'ChineseFont-Bold'
                else:
                    # If no bold file, just use regular (no fake bold support in simple TTFont check)
                    bold_font = 'ChineseFont'
                    
                return regular_font, bold_font
            except:
                continue

    return regular_font, bold_font

def export_pdf(order_data, filepath, report_type='delivery', seller_info=None):
    # Use Landscape A4
    page_size = landscape(A4)
    width, height = page_size 
    page_size = landscape(A4)
    width, height = page_size 
    font_reg, font_bold = register_fonts()
    
    c = canvas.Canvas(filepath, pagesize=page_size)
    
    
    raw_items = order_data.get('items', [])
    # Filter out empty items (where name is empty) to prevent "ghost" rows
    items = [i for i in raw_items if i.get('name', '').strip()]
    
    # Pagination: Max 8 items per page
    max_rows_per_page = 8

    
    # Calculate chunks
    if not items:
        page_chunks = [[]]
    else:
        page_chunks = [items[i:i + max_rows_per_page] for i in range(0, len(items), max_rows_per_page)]
    
    total_pages = len(page_chunks)
    
    # Calculate Grand Total
    grand_total = sum(float(item.get('price', 0)) * int(item.get('qty', 0)) for item in items)
    
    for page_idx, page_items in enumerate(page_chunks):
        is_last_page = (page_idx == total_pages - 1)
        
        # --- Draw Header (Same for all pages) ---
        
        # Title
        title_text = "销售出货单"
        if report_type == 'summary':
             title_text = "销售汇总单" # Optional: Change title? User didn't ask but "Order Summary" implies it. 
             # User said "Generate specific customer summary statement". 
             # But "Table content rules... consistent with 'Opening Order' logic".
             # "Header... exclude Customer Address and Order ID".
             # I should probably stick to "销售出货单" unless asked, OR change to "销售汇总表".
             # Let's keep "销售出货单" or make it dynamic if user asked? User didn't explicitly ask to change Title text.
             # However, "Order Summary" (订单汇总) allows for a different title.
             # Let's use "销售汇总单" for summary mode as a polish, or keep "销售出货单".
             # "Summary Filter Conditions... Manual Modification: Document Date... (2026-01-01 to...)"
             # Let's keep simple.
        
        # Actually, let's keep title "销售出货单" unless I see "Summary" in requirements.
        # "Generate specific customer summary statement" usually has title "Statement". 
        # I'll stick to input reqs: "Exclude Address and ID".
        
        c.setFont(font_bold, 26) 
        seller_name = seller_info.get('name', "广州市 XX 办公设备有限公司") if seller_info else "广州市 XX 办公设备有限公司"
        c.drawCentredString(width / 2, height - 50, seller_name)
        
        c.setFont(font_reg, 18)
        # Dynamic title based on type? Or fixed?
        # Let's use "销售汇总表" if summary.
        display_title = "销售汇总表" if report_type == 'summary' else "销售出货单"
        c.drawCentredString(width / 2, height - 85, display_title)

        # Header Info - user requested Right Align but with labels aligned
        y_pos = height - 125
        c.setFont(font_reg, 11)
        
        # Calculate Layout Guidelines
        col_widths = [50, 160, 160, 50, 60, 70, 70, 100] 
        left_margin = 60
        # Right margin aligned strictly to table width
        right_margin = left_margin + sum(col_widths)
        
        # Left side: Customer Info
        c.drawString(left_margin, y_pos, f"客户名称: {order_data.get('customer', '')}")
        
        # Right side: Date and ID (Aligned block)
        # Use simple width-based alignment for text to match previous look, or align to table right?
        # Let's align to table right (right_margin).
        # Previous: width - 240. width is usually ~841. right_margin is ~780.
        # 60 difference.
        header_right_block_x = width - 240
        c.drawString(header_right_block_x, y_pos, f"单据日期: {order_data.get('date', '')}")
        
        if report_type == 'delivery':
            y_pos -= 15
            c.drawString(left_margin, y_pos, f"客户地址: {order_data.get('address', '')}")
            
            # Right side: Order ID
            c.drawString(header_right_block_x, y_pos, f"单据编号: {order_data.get('order_id', '')}")

        # --- Table Header ---
        y_pos -= 25 
        table_top = y_pos + 12
        # col_widths defined above
        headers = ["序号", "商品名称", "规格型号", "单位", "数量", "单价", "合计", "备注"]
        
        c.setLineWidth(1)
        
        # Margin defined above
        
        c.line(left_margin, table_top, right_margin, table_top)
        c.line(left_margin, y_pos - 8, right_margin, y_pos - 8)
        
        current_x = left_margin
        for i, h in enumerate(headers):
            w = col_widths[i]
            c.drawCentredString(current_x + w/2, y_pos, h)
            current_x += w

        # --- Table Content ---
        row_height = 35
        y_pos -= row_height
        
        # Draw Items for this page
        start_idx = page_idx * max_rows_per_page
        
        for idx, item in enumerate(page_items):
            global_idx = start_idx + idx
            
            current_x = left_margin
            text_y = y_pos + 12 
            
            c.drawCentredString(current_x + col_widths[0]/2, text_y, str(global_idx + 1))
            current_x += col_widths[0]
            
            c.drawCentredString(current_x + col_widths[1]/2, text_y, item.get('name', ''))
            current_x += col_widths[1]
            
            c.drawCentredString(current_x + col_widths[2]/2, text_y, item.get('model', ''))
            current_x += col_widths[2]
            
            c.drawCentredString(current_x + col_widths[3]/2, text_y, item.get('unit', ''))
            current_x += col_widths[3]
            
            c.drawCentredString(current_x + col_widths[4]/2, text_y, str(item.get('qty', '')))
            current_x += col_widths[4]
            
            c.drawCentredString(current_x + col_widths[5]/2, text_y, f"{float(item.get('price', 0)):.2f}")
            current_x += col_widths[5]
            
            line_total = float(item.get('price', 0)) * int(item.get('qty', 0))
            c.drawCentredString(current_x + col_widths[6]/2, text_y, f"{line_total:.2f}")
            current_x += col_widths[6]
            
            c.drawCentredString(current_x + col_widths[7]/2, text_y, item.get('remark', ''))
            current_x += col_widths[7]
            
            # Line at bottom of row
            c.line(left_margin, y_pos, right_margin, y_pos)
            y_pos -= row_height

        # --- Empty Rows (Min 5 rule) ---
        # Only pad if total items <= 5. 
        # If items > 5, we do NOT pad, even if on last page with 1 item?
        # User said: "Insufficient 5 lines automatic make up logic stays unchanged, that is normal".
        # Current logic: if len(items) > 5: min_rows = 0.
        # This means if I have 20 items. Page 3 has 4 items.
        # len(items) = 20 > 5. min_rows = 0.
        # rows_on_page = 4. 4 < 0 is False. No padding.
        # So Page 3 has 4 items and STOPS.
        # This seems to be what is intended.
        
        if len(items) > 5:
            min_rows = 0
        else:
            min_rows = 5
            
        rows_on_this_page = len(page_items)
        
        target_rows = min_rows if rows_on_this_page < min_rows else rows_on_this_page
        
        while rows_on_this_page < target_rows:
            c.line(left_margin, y_pos, right_margin, y_pos)
            y_pos -= row_height
            rows_on_this_page += 1

        # Table Bottom Position (used for vertical lines)
        # y_pos is currently "one row height" BELOW the last drawn line.
        # We want vertical lines to stop at the last drawn line.
        table_bottom = y_pos + row_height
        
        # --- Draw Vertical Lines for Table Body ---
        current_x = left_margin
        c.line(current_x, table_top, current_x, table_bottom)
        for w in col_widths:
            current_x += w
            c.line(current_x, table_top, current_x, table_bottom)

        # --- Total Row (Only on Last Page) ---
        if is_last_page:
            # Total Row
            # We append it below the current table_bottom.
            # So Top of Total Row = table_bottom
            # Bottom of Total Row = y_pos (which is table_bottom - row_height)
            
            text_y = y_pos + 12
            
            # Calculate center of the merged block (Cols 0-5)
            merged_width = sum(col_widths[:6])
            merged_center_x = left_margin + (merged_width / 2)
            c.drawCentredString(merged_center_x, text_y, "合计")
            
            # Value in "Total" column (Index 6)
            total_col_start_x = left_margin + merged_width
            c.drawCentredString(total_col_start_x + col_widths[6]/2, text_y, f"{grand_total:.2f}")
            
            # Bottom line of Total Row
            c.line(left_margin, y_pos, right_margin, y_pos)
            
            # --- Vertical Lines for Total Row ---
            # 1. Left border
            c.line(left_margin, table_bottom, left_margin, y_pos)
            
            # 2. Line before Amount (Start of col 6)
            c.line(total_col_start_x, table_bottom, total_col_start_x, y_pos)
            
            # 3. Line after Amount (End of col 6 / Start of Remark)
            amount_col_end_x = total_col_start_x + col_widths[6]
            c.line(amount_col_end_x, table_bottom, amount_col_end_x, y_pos)
            
            # 4. Right border
            c.line(right_margin, table_bottom, right_margin, y_pos)
            
            # Update table_bottom so footer is relative to this new bottom
            table_bottom = y_pos

        # --- Footer (On Every Page) ---
        footer_y = table_bottom - 40
        
        c.drawString(left_margin, footer_y, f"制单人: {order_data.get('maker', '')}")
        
        # Align Signer to Total Column (Start of Col 6)
        total_col_start_x = left_margin + sum(col_widths[:6])
        c.drawString(total_col_start_x, footer_y, "签收人:")
        
        # Page Number
        c.drawCentredString(width/2, 10, f"- 第 {page_idx + 1} 页 / 共 {total_pages} 页 -")

        c.showPage()
        
    c.save()
