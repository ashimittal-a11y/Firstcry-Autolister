import sys
import os
import pandas as pd
import openpyxl

sys.path.append(os.path.join(os.path.dirname(__file__), "app"))
from mapping_engine import generate_listings

# Files
clothing_old_listing = r"C:\Users\Manann\Desktop\FirstcryAutolister\Listing\Clothing Old Listing.xlsm"
footwear_old_listing = r"C:\Users\Manann\Desktop\FirstcryAutolister\Listing\Footwear Old Listing.xlsm"
item_master = r"C:\Users\Manann\Desktop\FirstcryAutolister\Listing\ITEM DIRECTORY - ITEM MASTER.xlsx"
content_sheet = r"C:\Users\Manann\Desktop\FirstcryAutolister\Listing\Content Sheet (3).xlsx"

clothing_template = r"C:\Users\Manann\Desktop\FirstcryAutolister\Listing\Clothing Template.xlsm"
footwear_template = r"C:\Users\Manann\Desktop\FirstcryAutolister\Listing\Footwear Template.xlsm"

def test_engines():
    print("\n--- Dry Run: Clothing Engine ---")
    with open(item_master, 'rb') as f:
        item_bytes = f.read()
    with open(content_sheet, 'rb') as f:
        content_bytes = f.read()
    with open(clothing_template, 'rb') as f:
        temp_bytes = f.read()
        
    # Let's target style "TGCAFS003326" which has color "PINK" (or "TGCAFS003326-PINK")
    # Actually let's pass a known code from master sheet or content sheet
    item_codes = ["TGCAFS003326-PINK"] # Or TGCAFS003326-PINK
    
    # Run clothing engine
    print("Generating Clothing Listing...")
    out_bytes = generate_listings(
        item_directory_bytes=item_bytes,
        content_sheet_bytes=content_bytes,
        template_bytes=temp_bytes,
        mapping_config={"join_keys": {}, "mappings": {}},
        item_codes=item_codes,
        category="clothing"
    )
    
    # Inspect output
    import io
    wb = openpyxl.load_workbook(io.BytesIO(out_bytes), data_only=True)
    sheet = wb["Sheet1"]
    
    print("\nChecking generated clothing rows:")
    row3 = [sheet.cell(row=3, column=c).value for c in range(1, 65)]
    row4 = [sheet.cell(row=4, column=c).value for c in range(1, 65)]
    headers = [sheet.cell(row=2, column=c).value for c in range(1, 65)]
    
    for h, r3, r4 in zip(headers, row3, row4):
        if r3 is not None or r4 is not None:
            print(f"Col '{h}': Row3 -> '{r3}', Row4 -> '{r4}'")

    print("\n--- Dry Run: Footwear Engine ---")
    with open(footwear_template, 'rb') as f:
        fw_temp_bytes = f.read()
    # Let's check footwear style e.g. "TGSPOS003190" (or similar) from footwear master
    # We saw Footwear Old Listing has EAN like 8905120459528 which maps to UK 11
    # Let's pass 8905120459528
    fw_codes = ["8905120459528"]
    
    print("Generating Footwear Listing...")
    out_bytes_fw = generate_listings(
        item_directory_bytes=item_bytes,
        content_sheet_bytes=content_bytes,
        template_bytes=fw_temp_bytes,
        mapping_config={"join_keys": {}, "mappings": {}},
        item_codes=fw_codes,
        category="footwear"
    )
    
    wb_fw = openpyxl.load_workbook(io.BytesIO(out_bytes_fw), data_only=True)
    sheet_fw = wb_fw["Sheet1"]
    
    print("\nChecking generated footwear rows:")
    fw_headers = [sheet_fw.cell(row=2, column=c).value for c in range(1, 65)]
    fw_row3 = [sheet_fw.cell(row=3, column=c).value for c in range(1, 65)]
    
    for h, r3 in zip(fw_headers, fw_row3):
        if r3 is not None:
            print(f"Col '{h}': Row3 -> '{r3}'")

test_engines()
