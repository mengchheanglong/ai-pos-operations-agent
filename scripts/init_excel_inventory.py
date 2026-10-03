"""
Initialize and manage the local Excel inventory file (inventory.xlsx).
Creates the 4 standard sheets: Products, Orders, Customers, Logs.
Styled with headers, column widths, and proper formatting.
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(PROJECT_ROOT, "inventory.xlsx")

def create_excel_inventory(file_path=EXCEL_PATH):
    wb = Workbook()
    
    # ----------------------------------------------------
    # Styles
    # ----------------------------------------------------
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark Blue
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0")
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    # ----------------------------------------------------
    # 1. Products Sheet
    # ----------------------------------------------------
    ws_products = wb.active
    ws_products.title = "Products"
    
    prod_headers = ["SKU", "Product", "Variant", "Price", "Stock", "Low_Stock_Threshold", "Category"]
    ws_products.append(prod_headers)
    
    initial_products = [
        ["PR-QEN", "Yes I Am The Queen", "Pink 100ml", 40, 10, 2, "Beauty"],
        ["PR-KNG", "Yes I Am The King Le", "Black 100ml", 38, 10, 2, "Beauty"],
        ["PR-CHR", "Cherry Delice", "Red 100ml", 38, 10, 2, "Beauty"],
        ["PR-CHC", "Chic Oriental", "Gold 100ml", 39, 10, 2, "Beauty"],
        ["PR-BLU", "Bleu Impérial", "Blue 100ml", 38, 10, 2, "Beauty"],
        ["SH-BLK-B", "Basic Shirt", "Black", 10, 20, 5, "Apparel"],
        ["SH-WHT-B", "Basic Shirt", "White", 10, 18, 5, "Apparel"],
        ["SH-RED-B", "Basic Shirt", "Red", 10, 15, 5, "Apparel"],
        ["SH-BLK-P", "Premium Shirt", "Black", 18, 5, 2, "Apparel"],
        ["SH-WHT-P", "Premium Shirt", "White", 18, 8, 2, "Apparel"],
    ]
    
    for row in initial_products:
        ws_products.append(row)
        
    # Format Products Header & Data
    for col_idx, col in enumerate(ws_products.columns, start=1):
        for row_idx, cell in enumerate(col, start=1):
            cell.border = thin_border
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
            else:
                cell.font = data_font
                if col_idx in (4, 5, 6): # Price, Stock, Threshold
                    cell.alignment = right_align
                    if col_idx == 4:
                        cell.number_format = "$#,##0.00"
                else:
                    cell.alignment = left_align
        col_letter = get_column_letter(col_idx)
        ws_products.column_dimensions[col_letter].width = 20

    # ----------------------------------------------------
    # 2. Orders Sheet
    # ----------------------------------------------------
    ws_orders = wb.create_sheet(title="Orders")
    order_headers = ["Order_ID", "Customer_Name", "Product", "Variant", "Quantity", "Total_Price", "Status", "Timestamp", "Notes"]
    ws_orders.append(order_headers)
    
    sample_orders = [
        ["ORD-B94D74AD", "John Doe", "Basic Shirt", "Black", 2, 20, "Approved", "2026-10-03T07:10:00Z", "Auto-approved (<$100)"],
        ["ORD-1434B6E7", "Jane Smith", "Basic Shirt", "Black", 12, 120, "Escalated", "2026-10-03T07:11:00Z", "Manager Override (>= $100)"],
    ]
    for row in sample_orders:
        ws_orders.append(row)

    for col_idx, col in enumerate(ws_orders.columns, start=1):
        for row_idx, cell in enumerate(col, start=1):
            cell.border = thin_border
            if row_idx == 1:
                cell.fill = PatternFill(start_color="047857", end_color="047857", fill_type="solid") # Emerald
                cell.font = header_font
                cell.alignment = center_align
            else:
                cell.font = data_font
                cell.alignment = center_align if col_idx in (1, 5, 7, 8) else left_align
                if col_idx == 6:
                    cell.number_format = "$#,##0.00"
                    cell.alignment = right_align
        col_letter = get_column_letter(col_idx)
        ws_orders.column_dimensions[col_letter].width = 22

    # ----------------------------------------------------
    # 3. Customers Sheet
    # ----------------------------------------------------
    ws_customers = wb.create_sheet(title="Customers")
    cust_headers = ["Customer_ID", "Customer_Name", "Phone", "Email", "Address"]
    ws_customers.append(cust_headers)
    sample_cust = [
        ["CUST-001", "John Doe", "+855 12 345 678", "john.doe@example.com", "Phnom Penh, Cambodia"],
        ["CUST-002", "Jane Smith", "+855 98 765 432", "jane.smith@example.com", "Siem Reap, Cambodia"],
        ["CUST-003", "Bob Wilson", "+855 77 112 233", "bob.wilson@example.com", "Battambang, Cambodia"],
        ["CUST-004", "Alice Brown", "+855 88 445 566", "alice.brown@example.com", "Phnom Penh, Cambodia"],
    ]
    for row in sample_cust:
        ws_customers.append(row)

    for col_idx, col in enumerate(ws_customers.columns, start=1):
        for row_idx, cell in enumerate(col, start=1):
            cell.border = thin_border
            if row_idx == 1:
                cell.fill = PatternFill(start_color="374151", end_color="374151", fill_type="solid") # Slate
                cell.font = header_font
                cell.alignment = center_align
            else:
                cell.font = data_font
                cell.alignment = center_align if col_idx == 1 else left_align
        col_letter = get_column_letter(col_idx)
        ws_customers.column_dimensions[col_letter].width = 24

    # ----------------------------------------------------
    # 4. Logs Sheet
    # ----------------------------------------------------
    ws_logs = wb.create_sheet(title="Logs")
    log_headers = ["Log_ID", "Timestamp", "Action_Taken", "Order_ID", "Details", "Notification_Sent"]
    ws_logs.append(log_headers)
    sample_logs = [
        ["LOG-001", "2026-10-03T07:10:05Z", "approve", "ORD-B94D74AD", "Stock deducted 20 -> 18. Value: $20.", "Yes (Telegram)"],
        ["LOG-002", "2026-10-03T07:11:05Z", "escalate", "ORD-1434B6E7", "Order value $120 >= $100 threshold. Alerted owner.", "Yes (Telegram)"],
    ]
    for row in sample_logs:
        ws_logs.append(row)

    for col_idx, col in enumerate(ws_logs.columns, start=1):
        for row_idx, cell in enumerate(col, start=1):
            cell.border = thin_border
            if row_idx == 1:
                cell.fill = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid") # Amber
                cell.font = header_font
                cell.alignment = center_align
            else:
                cell.font = data_font
                cell.alignment = center_align if col_idx in (1, 2, 3, 4, 6) else left_align
        col_letter = get_column_letter(col_idx)
        ws_logs.column_dimensions[col_letter].width = 24

    wb.save(file_path)
    print(f"Excel inventory successfully generated at: {file_path}")

if __name__ == "__main__":
    create_excel_inventory()
