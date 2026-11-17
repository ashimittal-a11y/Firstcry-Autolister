import os
import json
import pandas as pd
import openpyxl
import io

def clean_dataframe(df):
    """
    Cleans a dataframe by stripping whitespaces from header rows
    and parsing it appropriately if it contains the Firstcry double-header structure.
    """
    if df.shape[0] < 2:
        return df
    
    # Check if Row 1 has column names (like SrNo, Style Code) and Row 0 has group names
    # Firstcry sheets have double headers. Row 1 is usually the actual column names.
    row0 = df.iloc[0].astype(str).tolist()
    row1 = df.iloc[1].astype(str).tolist()
    
    # If Row 1 has actual columns like 'Style Code' or 'Product Name'
    if any(x in ['Style Code', 'Product Name', 'SrNo'] for x in row1):
        headers = []
        for r0, r1 in zip(row0, row1):
            s1 = str(r1).strip() if pd.notna(r1) else ''
            s0 = str(r0).strip() if pd.notna(r0) else ''
            
            h = s1 if s1 not in ['nan', ''] else s0
            headers.append(h)
        cleaned_df = df.iloc[2:].copy()
        cleaned_df.columns = headers
        return cleaned_df
    else:
        # Standard header parsing
        df.columns = [str(c).strip() for c in df.iloc[0]]
        return df.iloc[1:].copy()

def learn_mappings(old_listing_bytes: bytes, item_directory_bytes: bytes, content_sheet_bytes: bytes):
    """
    Learns column mappings by analyzing a historical listing sheet
    against the Item Directory and Content Sheet.
    """
    try:
        # Load sheets
        old_df_raw = pd.read_excel(io.BytesIO(old_listing_bytes), sheet_name="Sheet1", header=None)
        master_df_raw = pd.read_excel(io.BytesIO(item_directory_bytes), header=None)
        content_df_raw = pd.read_excel(io.BytesIO(content_sheet_bytes), header=None)
        
        old_df = clean_dataframe(old_df_raw)
        master_df = clean_dataframe(master_df_raw)
        content_df = clean_dataframe(content_df_raw)
        
        # Strip string values
        for df in [old_df, master_df, content_df]:
            for idx in range(df.shape[1]):
                col_series = df.iloc[:, idx]
                if col_series.dtype == object:
                    df.iloc[:, idx] = col_series.astype(str).str.strip()
                    
        # 1. Identify the join keys between historical listing and Master File
        join_key_old = None
        join_key_master = None
        
        # We look for overlapping values between columns
        for col_old in old_df.columns:
            if not col_old or pd.isna(col_old) or str(col_old).startswith('Unnamed:'):
                continue
            vals_old = set(old_df[col_old].dropna().astype(str).unique())
            if len(vals_old) < 3: # Ignore columns with too few distinct values
                continue
                
            for col_master in master_df.columns:
                if not col_master or pd.isna(col_master) or str(col_master).startswith('Unnamed:'):
                    continue
                vals_master = set(master_df[col_master].dropna().astype(str).unique())
                intersection = vals_old.intersection(vals_master)
                
                # If there's high overlap, we found our key!
                if len(intersection) >= 3 and len(intersection) / len(vals_old) >= 0.7:
                    join_key_old = col_old
                    join_key_master = col_master
                    break
            if join_key_old:
                break
                
        # If overlap join key detection failed, fall back to column names
        if not join_key_old:
            for k in ['Style Code', 'Barcode', 'EAN', 'Manufacturer BarCode']:
                if k in old_df.columns and k in master_df.columns:
                    join_key_old = k
                    join_key_master = k
                    break
        
        # 2. Identify the join keys between historical listing and Content Sheet
        join_key_content = None
        # Often Content Sheet is joined via Style Code / prefix or title
        for col_content in content_df.columns:
            if not col_content or pd.isna(col_content) or str(col_content).startswith('Unnamed:'):
                continue
            vals_content = set(content_df[col_content].dropna().astype(str).unique())
            if len(vals_content) < 3:
                continue
            
            # Check overlap with the old listing style/barcode col
            if join_key_old:
                vals_old = set(old_df[join_key_old].dropna().astype(str).unique())
                intersection = vals_old.intersection(vals_content)
                if len(intersection) >= 3 and len(intersection) / len(vals_old) >= 0.5:
                    join_key_content = col_content
                    break
        
        if not join_key_content:
            # Look for common product code column names in content sheet
            for k in ['Style Code', 'Style ID', 'Item Code', 'Item Color', 'Product ID']:
                if k in content_df.columns:
                    join_key_content = k
                    break
            if not join_key_content and len(content_df.columns) > 0:
                join_key_content = content_df.columns[0] # Default to first col
        
        # 3. Create mapping map
        mappings = {}
        
        # If we have join keys, align rows to compare
        if join_key_old and join_key_master:
            aligned_df = old_df.merge(master_df, left_on=join_key_old, right_on=join_key_master, suffixes=('_old', '_master'))
            if join_key_content:
                aligned_df = aligned_df.merge(content_df, left_on=join_key_old, right_on=join_key_content, suffixes=('', '_content'))
        else:
            # Fallback: compare row by row directly if sizes match, or just use naming similarity
            aligned_df = pd.concat([old_df.reset_index(drop=True), master_df.reset_index(drop=True)], axis=1)
            
        # We only map columns 0 to 108 (which are the listing columns)
        # Get active listing column headers from Row 1
        raw_headers = old_df_raw.iloc[1].tolist()
        template_cols = [str(h).strip() for h in raw_headers if pd.notna(h) and str(h).strip() != 'nan' and str(h).strip() != '']
        
        # Cap mapping analysis to the columns found in Row 1 (up to index 109)
        template_cols = template_cols[:109]
        
        for t_col in template_cols:
            if t_col not in old_df.columns:
                continue
                
            # If all values in old_listing for this column are null/nan
            if old_df[t_col].isna().all() or (old_df[t_col] == "").all() or (old_df[t_col] == "nan").all():
                mappings[t_col] = {"type": "empty"}
                continue
                
            # Check if column is constant
            non_null_vals = old_df[t_col].dropna().unique()
            non_null_vals = [v for v in non_null_vals if v != 'nan' and v != '']
            if len(non_null_vals) == 1:
                mappings[t_col] = {"type": "constant", "value": non_null_vals[0]}
                continue
                
            # Check if it maps directly to a column in master_df (Item Directory)
            found_map = False
            for m_col in master_df.columns:
                if m_col == join_key_master:
                    continue
                # Compare aligned rows
                if t_col in aligned_df.columns and m_col in aligned_df.columns:
                    match_mask = (aligned_df[t_col] == aligned_df[m_col]) | (aligned_df[t_col].isna() & aligned_df[m_col].isna())
                    match_rate = match_mask.mean()
                    if match_rate >= 0.85: # Tolerant threshold
                        mappings[t_col] = {"type": "master", "column": m_col}
                        found_map = True
                        break
            
            if found_map:
                continue
                
            # Check if it maps directly to a column in content_df
            for c_col in content_df.columns:
                if c_col == join_key_content:
                    continue
                if t_col in aligned_df.columns and c_col in aligned_df.columns:
                    match_mask = (aligned_df[t_col] == aligned_df[c_col]) | (aligned_df[t_col].isna() & aligned_df[c_col].isna())
                    match_rate = match_mask.mean()
                    if match_rate >= 0.85:
                        mappings[t_col] = {"type": "content", "column": c_col}
                        found_map = True
                        break
                        
            if not found_map:
                # If no direct match, check if it's fuzzy matching column name
                if t_col in master_df.columns:
                    mappings[t_col] = {"type": "master", "column": t_col}
                elif t_col in content_df.columns:
                    mappings[t_col] = {"type": "content", "column": t_col}
                else:
                    # Default: keep empty or manual entry
                    mappings[t_col] = {"type": "manual"}
                    
        return {
            "join_keys": {
                "old_listing": join_key_old,
                "master": join_key_master,
                "content": join_key_content
            },
            "mappings": mappings
        }
    except Exception as e:
        import traceback
        print("Error learning mappings:", traceback.format_exc())
        raise e

CLOTHING_SIZES_LOOKUP = {
    "2-3Y": {"AgeFrom": "2 Years", "AgeTo": "3 Years", "UnitForAge": "Year"},
    "3-4Y": {"AgeFrom": "3 Years", "AgeTo": "4 Years", "UnitForAge": "Year"},
    "4-5Y": {"AgeFrom": "4 Years", "AgeTo": "5 Years", "UnitForAge": "Year"},
    "5-6Y": {"AgeFrom": "5 Years", "AgeTo": "6 Years", "UnitForAge": "Year"},
    "7-8Y": {"AgeFrom": "7 Years", "AgeTo": "8 Years", "UnitForAge": "Year"},
    "9-10Y": {"AgeFrom": "9 Years", "AgeTo": "10 Years", "UnitForAge": "Year"},
    "11-12Y": {"AgeFrom": "11 Years", "AgeTo": "12 Years", "UnitForAge": "Year"},
    "13-14Y": {"AgeFrom": "13 Years", "AgeTo": "14 Years", "UnitForAge": "Year"},
    "2-6Y": {"AgeFrom": "2 Years", "AgeTo": "6 Years", "UnitForAge": "Year"},
    "7-14Y": {"AgeFrom": "7 Years", "AgeTo": "14 Years", "UnitForAge": "Year"}
}

FOOTWEAR_SIZES_LOOKUP = {
    "11": {"Size": "UK 11", "AgeFrom": "5.5 Years", "AgeTo": "6 Years", "UnitForAge": "Year", "Toe to Heel Size": "19.2 (FA 122)"},
    "12": {"Size": "UK 12", "AgeFrom": "6 Years", "AgeTo": "7 Years", "UnitForAge": "Year", "Toe to Heel Size": "20 (FA 130)"},
    "13": {"Size": "UK 13", "AgeFrom": "7 Years", "AgeTo": "8 Years", "UnitForAge": "Year", "Toe to Heel Size": "20.5 (FA 135)"},
    "1": {"Size": "UK 1", "AgeFrom": "8 Years", "AgeTo": "9 Years", "UnitForAge": "Year", "Toe to Heel Size": "21 (FA 140)"},
    "2": {"Size": "UK 2", "AgeFrom": "9 Years", "AgeTo": "10 Years", "UnitForAge": "Year", "Toe to Heel Size": "21.5 (FA 145)"},
    "3": {"Size": "UK 3", "AgeFrom": "10 Years", "AgeTo": "11 Years", "UnitForAge": "Year", "Toe to Heel Size": "22.5 (FA 155)"},
    "9": {"Size": "UK 9", "AgeFrom": "3.5 Years", "AgeTo": "4 Years", "UnitForAge": "Year", "Toe to Heel Size": "18 (FA 110)"},
    "10": {"Size": "UK 10", "AgeFrom": "4 Years", "AgeTo": "4.5 Years", "UnitForAge": "Year", "Toe to Heel Size": "18.7 (FA 117)"},
    "5": {"Size": "UK 5", "AgeFrom": "18 Month", "AgeTo": "24 Month", "UnitForAge": "Month", "Toe to Heel Size": "15.2 (FA 82)"},
    "6": {"Size": "UK 6", "AgeFrom": "2 Years", "AgeTo": "2.5 Years", "UnitForAge": "Year", "Toe to Heel Size": "15.9 (FA 89)"},
    "7": {"Size": "UK 7", "AgeFrom": "2.5 Years", "AgeTo": "3 Years", "UnitForAge": "Year", "Toe to Heel Size": "16.6 (FA 96)"},
    "8": {"Size": "UK 8", "AgeFrom": "3 Years", "AgeTo": "3.5 Years", "UnitForAge": "Year", "Toe to Heel Size": "17.3 (FA 103)"},
    "4": {"Size": "UK 4", "AgeFrom": "10 Years", "AgeTo": "11 Years", "UnitForAge": "Year", "Toe to Heel Size": "23.3 (FA 163)"},
    "5L": {"Size": "UK 5I", "AgeFrom": "11 Years", "AgeTo": "12 Years", "UnitForAge": "Year", "Toe to Heel Size": "24 (FA 170)"}
}

import re

def parse_clothing_size(size_str):
    size_str = str(size_str).strip()
    if size_str in CLOTHING_SIZES_LOOKUP:
        return CLOTHING_SIZES_LOOKUP[size_str]
    # Check if we can parse it
    m_range = re.match(r'^(\d+)-(\d+)\s*([mM])', size_str)
    if m_range:
        g1, g2 = m_range.group(1), m_range.group(2)
        return {"AgeFrom": f"{g1} Month", "AgeTo": f"{g2} Month", "UnitForAge": "Month"}
    y_range = re.match(r'^(\d+)-(\d+)\s*([yY])', size_str)
    if y_range:
        g1, g2 = y_range.group(1), y_range.group(2)
        return {"AgeFrom": f"{g1} Years", "AgeTo": f"{g2} Years", "UnitForAge": "Year"}
    y_single = re.match(r'^(\d+)\s*([yY])', size_str)
    if y_single:
        g1 = y_single.group(1)
        g2 = str(int(g1) + 1)
        return {"AgeFrom": f"{g1} Years", "AgeTo": f"{g2} Years", "UnitForAge": "Year"}
    m_single = re.match(r'^(\d+)\s*([mM])', size_str)
    if m_single:
        g1 = m_single.group(1)
        g2 = str(int(g1) + 3)
        return {"AgeFrom": f"{g1} Month", "AgeTo": f"{g2} Month", "UnitForAge": "Month"}
    return {"AgeFrom": "", "AgeTo": "", "UnitForAge": "Year"}

def parse_footwear_size(size_str):
    size_str = str(size_str).strip()
    if size_str in FOOTWEAR_SIZES_LOOKUP:
        return FOOTWEAR_SIZES_LOOKUP[size_str]
    # Clean size UK 11 -> 11
    clean_sz = size_str.replace("UK", "").strip()
    if clean_sz in FOOTWEAR_SIZES_LOOKUP:
        return FOOTWEAR_SIZES_LOOKUP[clean_sz]
    return {"Size": f"UK {size_str}", "AgeFrom": "", "AgeTo": "", "UnitForAge": "Year", "Toe to Heel Size": ""}

def get_clothing_subcategory(master_category):
    mc = str(master_category).upper().strip()
    if mc in ['DENIM', 'JEANS', 'SHORTS', 'SKIRT', 'CARGOS', 'CHINOS', 'TROUSERS', 'LOWER', 'CAPRI', 'JEGGING', 'LEGGING', 'PANTS', 'SKORT']:
        return 'Shorts, Skirts & Jeans'
    elif mc in ['T-SHIRT', 'TOP', 'POLO T-SHIRT', 'SWEATSHIRT', 'JACKET', 'SWEATER', 'VEST', 'SKIVI', 'HOODIE']:
        return 'Tops'
    elif mc in ['CLOTHING SET', 'MULTIPACK', 'JUMPSUIT', 'DUNGAREE', 'ROMPER', 'PARTYWEAR', 'UNIFORM', 'CO-ORD SET', 'GIFT PACK']:
        return 'Sets & Suits'
    elif mc in ['DRESS', 'FROCK', 'PARTY FROCK', 'PARTYWEAR GOWN']:
        return 'Frocks and Dresses'
    elif mc in ['SHIRT', 'SHACKET']:
        return 'Shirts'
    return 'Tops'

def clean_fc_color(color_name):
    col = str(color_name).upper().strip()
    if 'BLUE' in col or 'AQUA' in col or 'TEAL' in col:
        return 'Navy Blue' if 'NAVY' in col else 'Blue'
    elif 'BLACK' in col:
        return 'Black'
    elif 'WHITE' in col or 'OFF-WHITE' in col:
        return 'Offwhite' if 'OFF' in col else 'White'
    elif 'YELLOW' in col:
        return 'Yellow'
    elif 'RED' in col:
        return 'Red'
    elif 'PINK' in col:
        return 'Pink'
    elif 'BROWN' in col:
        return 'Brown'
    elif 'GREY' in col or 'GRAY' in col:
        return 'Grey'
    elif 'ORANGE' in col:
        return 'Orange'
    elif 'PEACH' in col:
        return 'Peach'
    elif 'GREEN' in col:
        return 'Green'
    elif 'PURPLE' in col:
        return 'Purple'
    elif 'MUSTARD' in col:
        return 'Mustard'
    elif 'TAN' in col:
        return 'Tan'
    return col.title()

def find_content_row(master_row, content_df):
    item_code = str(master_row.get('ITEM CODE', '')).strip().split('.')[0]
    sku_cols = [c for c in content_df.columns if str(c).upper() in ['SKU', 'ITEM COLOR', 'PRODUCT ID']]
    
    # 1. Match item code
    if item_code and item_code != 'nan':
        for col in content_df.columns:
            col_series = content_df[col].fillna('').astype(str).apply(lambda x: x.strip().split('.')[0])
            matches = content_df[col_series == item_code]
            if not matches.empty:
                return matches.iloc[0]
                
    # 2. Match Item Color
    item_color = str(master_row.get('Item Color', '')).strip()
    if item_color and item_color != 'nan':
        for col in sku_cols:
            col_series = content_df[col].fillna('').astype(str).apply(lambda x: x.strip().lower())
            matches = content_df[col_series == item_color.lower()]
            if not matches.empty:
                return matches.iloc[0]
                
    # 3. Match style prefix
    style_prefix = item_color.rsplit('-', 1)[0].strip() if '-' in item_color else item_color
    if style_prefix and style_prefix != 'nan':
        style_cols = [c for c in content_df.columns if str(c).upper() in ['ITEM NAME', 'STYLE CODE', 'STYLE ID']]
        for col in style_cols:
            col_series = content_df[col].fillna('').astype(str).apply(lambda x: x.strip().lower())
            matches = content_df[col_series == style_prefix.lower()]
            if not matches.empty:
                return matches.iloc[0]
                
        # Fallback split on SKU
        for col in sku_cols:
            col_series = content_df[col].fillna('').astype(str).apply(lambda x: x.strip().rsplit('-', 1)[0].strip().lower() if '-' in x else x.strip().lower())
            matches = content_df[col_series == style_prefix.lower()]
            if not matches.empty:
                return matches.iloc[0]
                
    return None

def normalize_item_code(code):
    # Convert letter 'O' in style code to '0' if it is followed by numbers
    # E.g. TBCABSO03322-NAVY/BEIGE -> TBCABS003322-NAVY/BEIGE
    code = str(code).strip().upper()
    code_fixed = re.sub(r'([A-Z]+)O(\d+)', r'\g<1>0\g<2>', code)
    return code_fixed

def generate_listings(item_directory_bytes: bytes, content_sheet_bytes: bytes, template_bytes: bytes, mapping_config: dict, item_codes: list = None, category: str = None) -> bytes:
    """
    Generates a new listing by populating the template Excel sheet starting at Row 2 (under double headers)
    for columns 0 to 108. Uses category-specific engines (clothing/footwear) or fallback mappings.
    """
    # 1. Read input data sheets
    master_df_raw = pd.read_excel(io.BytesIO(item_directory_bytes), header=None)
    content_df_raw = pd.read_excel(io.BytesIO(content_sheet_bytes), header=None)
    
    master_df = clean_dataframe(master_df_raw)
    content_df = clean_dataframe(content_df_raw)
    
    join_keys = mapping_config.get("join_keys", {})
    mappings = mapping_config.get("mappings", {})
    
    join_key_master = join_keys.get("master")
    
    # Standardize types for key lookup
    for df in [master_df, content_df]:
        for idx in range(df.shape[1]):
            col_series = df.iloc[:, idx]
            if col_series.dtype == object:
                df.iloc[:, idx] = col_series.astype(str).str.strip()

    # Determine what rows from master_df we need to keep based on item codes filter
    matched_rows = []
    if item_codes:
        cleaned_codes = []
        for c in item_codes:
            c = str(c).strip()
            if not c:
                continue
            # Normalize typos like TBCABSO03322 to TBCABS003322
            norm_c = normalize_item_code(c)
            cleaned_codes.append(norm_c)
            
        style_col = join_key_master
        color_col = None
        
        # Detect the color column from mappings
        for t_c, rule in mappings.items():
            if t_c in ['Pantone/Vendor Color', 'Color Name'] and rule.get('type') == 'master':
                color_col = rule.get('column')
                break
                
        if not style_col:
            for c in master_df.columns:
                if 'code' in str(c).lower() or 'style' in str(c).lower() or 'barcode' in str(c).lower():
                    style_col = c
                    break
        if not color_col:
            for c in master_df.columns:
                if 'color' in str(c).lower() or 'pantone' in str(c).lower():
                    color_col = c
                    break
                    
        for code in cleaned_codes:
            exact_match_found = False
            for col in master_df.columns:
                matches = master_df[master_df[col].astype(str).str.lower() == code.lower()]
                if not matches.empty:
                    for idx, row in matches.iterrows():
                        matched_rows.append(row)
                    exact_match_found = True
                    break
            
            if exact_match_found:
                continue
                
            if "-" in code:
                parts = code.rsplit("-", 1)
                style_part = parts[0].strip()
                color_part = parts[1].strip()
                
                matches = master_df
                if style_col and style_col in master_df.columns:
                    matches = matches[matches[style_col].astype(str).str.lower() == style_part.lower()]
                if color_col and color_col in master_df.columns:
                    matches = matches[matches[color_col].astype(str).str.lower() == color_part.lower()]
                
                if not matches.empty:
                    for idx, row in matches.iterrows():
                        matched_rows.append(row)
    else:
        for idx, row in master_df.iterrows():
            matched_rows.append(row)

    # 2. Merge master data and content data
    merged_data = []
    for row in matched_rows:
        content_row = find_content_row(row, content_df)
        merged_data.append({
            "master_row": row,
            "content_row": content_row,
            "key_val": str(row.get('Item Color', '')).strip()
        })

    # 3. Load Template Workbook with macros (keep_vba=True)
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes), keep_vba=True)
    sheet = wb["Sheet1"]
    
    # Template Row 1 (actual column headers in Python 0-indexed is row 1)
    raw_headers = [sheet.cell(row=2, column=c).value for c in range(1, 110)] # cols 1 to 109
    raw_headers_clean = [str(h).strip() for h in raw_headers if h]
    
    # 4. Fill values in rows starting from Row 2 (which is row index 3 in openpyxl, since Excel is 1-indexed)
    current_excel_row = 3
    
    # Clean inputs category name & Auto-detect Category Engine based on template headers!
    cat_engine = str(category).lower().strip() if category else None
    if not cat_engine or cat_engine not in ['clothing', 'footwear']:
        if 'Default Footwear Size' in raw_headers_clean or 'Toe to Heel Size' in raw_headers_clean:
            cat_engine = 'footwear'
        elif 'Neckline' in raw_headers_clean or 'Sleeve Length' in raw_headers_clean or 'Circumference - In Cms' in raw_headers_clean:
            cat_engine = 'clothing'
        else:
            cat_engine = 'clothing' # Default fallback
    
    for item in merged_data:
        master_row = item["master_row"]
        content_row = item["content_row"]
        key_val = item["key_val"]
        
        # Pre-compute some values if using a dedicated engine
        if cat_engine == 'clothing':
            # Subcategory helper
            sub_cat = get_clothing_subcategory(master_row.get('CATEGORY', ''))
            
            # Gender helper
            mg = str(master_row.get('GENDER', '')).upper().strip()
            if 'BOY' in mg or 'MALE' in mg:
                gender_val = 'Male'
            elif 'GIRL' in mg or 'FEMALE' in mg:
                gender_val = 'Female'
            else:
                gender_val = 'Unisex'
                
            # Age helper
            age_info = parse_clothing_size(master_row.get('SIZE', ''))
            
            # PoType helper
            if sub_cat == 'Shorts, Skirts & Jeans':
                mc_lower = str(master_row.get('CATEGORY', '')).upper().strip()
                if mc_lower in ['JEANS', 'DENIM']:
                    po_type = 'Jeans'
                elif mc_lower == 'SHORTS':
                    po_type = 'Shorts'
                elif mc_lower == 'SKIRT':
                    po_type = 'Skirt'
                else:
                    po_type = 'Trousers'
            elif sub_cat == 'Tops':
                po_type = 'Tops'
            elif sub_cat == 'Sets & Suits':
                po_type = 'Top and Bottomwear'
            elif sub_cat == 'Frocks and Dresses':
                po_type = 'One Piece Dresses / Frocks'
            elif sub_cat == 'Shirts':
                po_type = 'Printed Shirts'
            else:
                po_type = str(master_row.get('CATEGORY', '')).title()
                
            # Fabric helpers
            m_fabric = str(master_row.get('FABRIC', master_row.get('MATERIAL', ''))).strip().title()
            if m_fabric.lower() in ['nan', 'none', '']:
                m_fabric = 'Cotton'
            if sub_cat in ['Tops', 'Shirts']:
                upper_fabric = m_fabric
                lower_fabric = None
            elif sub_cat == 'Shorts, Skirts & Jeans':
                upper_fabric = None
                lower_fabric = m_fabric
            else:
                upper_fabric = m_fabric
                lower_fabric = m_fabric
                
        elif cat_engine == 'footwear':
            # Size and dimension lookups
            raw_sz = str(master_row.get('SIZE', '')).strip()
            sz_info = parse_footwear_size(raw_sz)
            
            # Gender helper
            mg = str(master_row.get('GENDER', '')).upper().strip()
            if 'BOY' in mg or 'MALE' in mg:
                gender_val = 'Male'
            elif 'GIRL' in mg or 'FEMALE' in mg:
                gender_val = 'Female'
            else:
                gender_val = 'Unisex'
                
            # PoType helper
            mc_lower = str(master_row.get('CATEGORY', '')).upper().strip()
            if mc_lower == 'BALLERINAS':
                po_type = 'Ballerina'
            elif mc_lower in ['CASUAL SHOES', 'SPORTS SHOES', 'BOOTS']:
                po_type = 'Casual Shoes'
            elif mc_lower == 'SANDALS':
                po_type = 'Sandals'
            elif mc_lower in ['FLIP FLOPS', 'SLIDES']:
                po_type = 'Flip Flops'
            else:
                po_type = mc_lower.title()
                
            if not po_type or po_type.lower() == 'nan' or po_type == '':
                po_type = 'Casual Shoes'
                
            # Fabric helpers
            upper_fabric = str(master_row.get('UPPER', '')).strip().title()
            if not upper_fabric or upper_fabric.lower() in ['nan', 'none', '']:
                upper_fabric = 'Synthetic'
                
            lower_fabric = str(master_row.get('SOLE', '')).strip().title()
            if not lower_fabric or lower_fabric.lower() in ['nan', 'none', '']:
                lower_fabric = 'Tpr'
        
        # For each column in template (cols 1 to 109 in openpyxl matching columns 0 to 108)
        for col_idx in range(109):
            col_name = raw_headers[col_idx]
            if not col_name:
                continue
            col_name = str(col_name).strip()
            
            val_to_write = None
            
            # Engine logic branches
            if cat_engine == 'clothing':
                if col_name == 'Brand Name':
                    val_to_write = 'Purple United Kids'
                elif col_name == 'Style Code':
                    val_to_write = str(master_row.get('ITEM CODE', '')).split('.')[0].strip()
                elif col_name == 'Product Name':
                    prod_name = content_row.get('Firstcry Title', master_row.get('ITEM NAME', '')) if content_row is not None else master_row.get('ITEM NAME', '')
                    color_val = str(master_row.get('COLOR', '')).strip()
                    if color_val and color_val.lower() not in ['nan', 'none', '']:
                        val_to_write = f"{prod_name}-{color_val}"
                    else:
                        val_to_write = prod_name
                elif col_name == 'Product Description':
                    val_to_write = content_row.get('Description', master_row.get('ITEM DESCRIPTION', '')) if content_row is not None else master_row.get('ITEM DESCRIPTION', '')
                elif col_name == 'CategoryName':
                    val_to_write = 'Clothes & Shoes'
                elif col_name == 'SubCategoryName':
                    val_to_write = sub_cat
                elif col_name in ['Size', 'Brand Size']:
                    val_to_write = master_row.get('SIZE', '')
                elif col_name == 'Pantone/Vendor Color':
                    val_to_write = master_row.get('COLOR', '')
                elif col_name == 'MRP':
                    val_to_write = master_row.get('MRP', '')
                elif col_name in ['HSNCode', '8DigitHSNCode']:
                    val_to_write = str(master_row.get('HS CODE', '')).split('.')[0].strip()
                elif col_name == 'MadeIn':
                    val_to_write = 'India'
                elif col_name == 'Gender':
                    val_to_write = gender_val
                elif col_name == 'UnitForAge':
                    val_to_write = age_info.get('UnitForAge', 'Year')
                elif col_name == 'AgeFrom':
                    val_to_write = age_info.get('AgeFrom', '')
                elif col_name == 'AgeTo':
                    val_to_write = age_info.get('AgeTo', '')
                elif col_name == 'Top Fabric/Upper Fabric':
                    val_to_write = upper_fabric
                elif col_name == 'Bottom Fabric/Sole Material':
                    val_to_write = lower_fabric
                elif col_name == 'PoType':
                    val_to_write = po_type
                elif col_name == 'Silhouette/Shape':
                    val_to_write = str(master_row.get('SILHOUETTE', '')).strip().title()
                elif col_name == 'Pattern/Print type':
                    title = str(content_row.get('Firstcry Title', '')).upper() if content_row is not None else ''
                    val_to_write = 'Stripes' if 'STRIPE' in title else 'Solid'
                elif col_name == 'Pattern Length':
                    val_to_write = 'Full Length' if sub_cat == 'Shorts, Skirts & Jeans' else ('Knee length' if sub_cat == 'Frocks and Dresses' else 'Regular')
                elif col_name == 'Sleeve Length':
                    title = str(content_row.get('Firstcry Title', '')).upper() if content_row is not None else ''
                    if 'FULL SLEEVE' in title:
                        val_to_write = 'Full Sleeves'
                    elif 'SLEEVELESS' in title:
                        val_to_write = 'Sleeveless'
                    else:
                        val_to_write = 'Half sleeves'
                elif col_name == 'Neckline':
                    title = str(content_row.get('Firstcry Title', '')).upper() if content_row is not None else ''
                    val_to_write = 'Collar' if 'COLLAR' in title else 'Round Neck'
                elif col_name == 'Neck Opening':
                    val_to_write = 'Front-Buttons' if sub_cat == 'Shirts' else None
                elif col_name == 'Highlight':
                    val_to_write = 'Sun Protection'
                elif col_name == 'Fabric Type':
                    mc_upper = str(master_row.get('CATEGORY', '')).upper()
                    val_to_write = 'DENIM' if mc_upper in ['JEANS', 'DENIM'] else 'Woven'
                elif col_name == 'Packaging':
                    val_to_write = 'Single'
                elif col_name == 'Utility':
                    val_to_write = 'Fashion'
                elif col_name == 'No. of Units in a pack':
                    val_to_write = '2 pcs' if sub_cat == 'Sets & Suits' else '1 pc'
                elif col_name == 'Primary FC Color':
                    val_to_write = clean_fc_color(master_row.get('COLOR', ''))
                elif col_name == 'Items Included in Packaging 1':
                    mc_upper = str(master_row.get('CATEGORY', '')).upper()
                    if mc_upper in ['SHIRT', 'POLO T-SHIRT']:
                        val_to_write = '1 Shirt'
                    elif mc_upper == 'T-SHIRT':
                        val_to_write = '1 T-Shirt'
                    elif mc_upper == 'TOP':
                        val_to_write = '1 Top'
                    elif mc_upper in ['DRESS', 'FROCK']:
                        val_to_write = '1 Frock'
                    elif mc_upper in ['JEANS', 'DENIM']:
                        val_to_write = '1 Jeans'
                    elif mc_upper == 'SHORTS':
                        val_to_write = '1 Shorts'
                    elif mc_upper == 'SKIRT':
                        val_to_write = '1 Skirt'
                    else:
                        val_to_write = f"1 {str(master_row.get('CATEGORY', '')).title()}"
                elif col_name == 'Items Included in Packaging 2':
                    val_to_write = '1 Shorts' if sub_cat == 'Sets & Suits' else None
                    
            elif cat_engine == 'footwear':
                if col_name == 'Brand Name':
                    val_to_write = 'Toothless'
                elif col_name == 'Style Code':
                    val_to_write = str(master_row.get('ITEM CODE', '')).split('.')[0].strip()
                elif col_name == 'Product Name':
                    prod_name = content_row.get('Firstcry Title', master_row.get('ITEM NAME', '')) if content_row is not None else master_row.get('ITEM NAME', '')
                    color_val = str(master_row.get('COLOR', '')).strip()
                    if color_val and color_val.lower() not in ['nan', 'none', '']:
                        val_to_write = f"{prod_name}-{color_val}"
                    else:
                        val_to_write = prod_name
                elif col_name == 'Product Description':
                    val_to_write = content_row.get('Description', master_row.get('ITEM DESCRIPTION', '')) if content_row is not None else master_row.get('ITEM DESCRIPTION', '')
                elif col_name == 'CategoryName':
                    val_to_write = 'Clothes & Shoes'
                elif col_name == 'SubCategoryName':
                    val_to_write = 'Footwear'
                elif col_name == 'Default Footwear Size':
                    val_to_write = 'uksize'
                elif col_name == 'Size':
                    val_to_write = sz_info.get('Size', f"UK {raw_sz}")
                elif col_name == 'Brand Size':
                    val_to_write = raw_sz
                elif col_name == 'Pantone/Vendor Color':
                    val_to_write = master_row.get('COLOR', '')
                elif col_name == 'MRP':
                    val_to_write = master_row.get('MRP', '')
                elif col_name in ['HSNCode', '8DigitHSNCode']:
                    val_to_write = str(master_row.get('HS CODE', '')).split('.')[0].strip()
                elif col_name == 'MadeIn':
                    val_to_write = 'India'
                elif col_name == 'Gender':
                    val_to_write = gender_val
                elif col_name in ['UnitForAge', 'AgeFrom', 'AgeTo']:
                    val_to_write = None
                elif col_name == 'Top Fabric/Upper Fabric':
                    val_to_write = str(master_row.get('UPPER', '')).strip().title()
                elif col_name == 'Bottom Fabric/Sole Material':
                    val_to_write = str(master_row.get('SOLE', '')).strip().title()
                elif col_name == 'PoType':
                    val_to_write = po_type
                elif col_name == 'Pattern/Print type':
                    val_to_write = 'Solid'
                elif col_name == 'Packaging':
                    val_to_write = 'Single Pair'
                elif col_name == 'Utility':
                    val_to_write = 'Summer'
                elif col_name == 'No. of Units in a pack':
                    val_to_write = 'Single Pair'
                elif col_name == 'Toe to Heel Size':
                    val_to_write = sz_info.get('Toe to Heel Size', '')
                elif col_name == 'Primary FC Color':
                    val_to_write = clean_fc_color(master_row.get('COLOR', ''))
                elif col_name == 'Items Included in Packaging 1':
                    if po_type == 'Ballerina':
                        val_to_write = '1 Pair of Ballerinas'
                    elif po_type in ['Casual Shoes', 'Sports Shoes']:
                        val_to_write = '1 Pair of Shoes'
                    elif po_type == 'Sandals':
                        val_to_write = '1 Pair of Sandals'
                    elif po_type == 'Flip Flops':
                        val_to_write = '1 Pair of Flip Flops'
                    else:
                        val_to_write = f"1 Pair of {po_type}s"
                        
            else:
                # Standard fallback to mapping config UI rules
                rule = mappings.get(col_name, {})
                rule_type = rule.get("type")
                
                if rule_type == "constant":
                    val_to_write = rule.get("value")
                elif rule_type == "master":
                    src_col = rule.get("column")
                    if src_col in master_row:
                        val_to_write = master_row[src_col]
                elif rule_type == "content":
                    src_col = rule.get("column")
                    if content_row is not None and src_col in content_row:
                        val_to_write = content_row[src_col]
                elif rule_type == "manual" or rule_type == "empty":
                    val_to_write = None
                else:
                    if col_name in master_row:
                        val_to_write = master_row[col_name]
                    elif content_row is not None and col_name in content_row:
                        val_to_write = content_row[col_name]
            
            # Write value if not None
            if pd.notna(val_to_write) and val_to_write is not None:
                val_str = str(val_to_write).strip()
                if val_str.lower() in ['nan', 'none', 'null', '']:
                    val_to_write = None
                else:
                    if isinstance(val_to_write, (int, float)):
                        pass
                    else:
                        val_to_write = val_str
                sheet.cell(row=current_excel_row, column=col_idx + 1, value=val_to_write)
            else:
                sheet.cell(row=current_excel_row, column=col_idx + 1, value=None)
                
        current_excel_row += 1
        
    # Save modified sheet back to byte stream
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue()
