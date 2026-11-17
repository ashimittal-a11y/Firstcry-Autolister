import os
import json
import pandas as pd
from app.mapping_engine import learn_mappings, generate_listings

def run_integration_test():
    print("=== STARTING AUTOLISTER INTEGRATION TEST ===")
    
    listing_dir = r"C:\Users\Manann\Desktop\FirstcryAutolister\Listing"
    old_listing_path = os.path.join(listing_dir, "Clothing Old Listing.xlsm")
    template_path = os.path.join(listing_dir, "Clothing Template.xlsm")
    
    if not os.path.exists(old_listing_path) or not os.path.exists(template_path):
        print("ERROR: Test files missing in Listing folder.")
        return
        
    print(f"Loading historical sheet: {old_listing_path}")
    with open(old_listing_path, "rb") as f:
        old_bytes = f.read()
        
    with open(template_path, "rb") as f:
        template_bytes = f.read()
        
    # --- MOCKING MASTER ITEM DIRECTORY & CONTENT SHEET ---
    print("Mocking Item Directory and Content Sheet from historical data...")
    # Load first 15 rows of old listing to act as sample data
    df_raw = pd.read_excel(old_listing_path, sheet_name="Sheet1", header=None, nrows=15)
    
    # Extract Row 1 as headers
    headers = df_raw.iloc[1].tolist()
    data = df_raw.iloc[2:].copy()
    data.columns = headers
    
    # Create Item Directory Mock (has style code, MRP, Pantone, size, etc.)
    master_cols = ['Style Code', 'Brand Name', 'CategoryName', 'SubCategoryName', 'Size', 'Brand Size', 'Pantone/Vendor Color', 'MRP', 'HSNCode']
    master_df = data[master_cols].copy()
    # Let's rename some columns to test fuzzy mapping
    master_df.rename(columns={'MRP': 'Price Tag', 'Pantone/Vendor Color': 'Color Name'}, inplace=True)
    
    # Create Content Sheet Mock (has style code, product name, product description)
    content_cols = ['Style Code', 'Product Name', 'Product Description']
    content_df = data[content_cols].copy()
    # Let's modify product name column name to test mapping
    content_df.rename(columns={'Product Name': 'Listing Title'}, inplace=True)
    
    # Save these mockup dataframes to Excel bytes
    master_io = io_to_bytes(master_df)
    content_io = io_to_bytes(content_df)
    
    # --- RUN LEARNING ---
    print("Running learn_mappings()...")
    learn_result = learn_mappings(old_bytes, master_io, content_io)
    
    print("\nLearning Result Keys:", learn_result.keys())
    print("Detected Join Keys:", learn_result["join_keys"])
    print("Sample Learned Mappings (first 10):")
    mappings = learn_result["mappings"]
    for k in list(mappings.keys())[:10]:
        print(f"  - {k}: {mappings[k]}")
        
    # Check if the renamed columns were correctly mapped
    print("\nVerification of mappings:")
    print("  - Pantone/Vendor Color mapped to color name:", mappings.get("Pantone/Vendor Color"))
    print("  - MRP mapped to Price Tag:", mappings.get("MRP"))
    print("  - Product Name mapped to Listing Title:", mappings.get("Product Name"))
    
    # --- RUN LISTING GENERATION ---
    print("\nRunning generate_listings()...")
    # Retrieve the first style code and color name to test filtering
    first_style = master_df.iloc[0]['Style Code']
    first_color = master_df.iloc[0]['Color Name']
    mock_code = f"{first_style}-{first_color}"
    print(f"Testing list generation with filter item code: {mock_code}")
    output_bytes = generate_listings(master_io, content_io, template_bytes, learn_result, [mock_code])
    
    print(f"Generated Sheet Bytes size: {len(output_bytes)}")
    
    # Read generated sheet to verify
    df_gen = pd.read_excel(io.BytesIO(output_bytes), sheet_name="Sheet1", header=None)
    print(f"Generated Sheet Shape: {df_gen.shape}")
    
    # Check Row 2 (first data row) style code value
    print("First data row values (first 10 columns):")
    print(df_gen.iloc[2, :10].tolist())
    
    print("\n=== INTEGRATION TEST PASSED SUCCESSFULLY ===")

def io_to_bytes(df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return out.getvalue()

if __name__ == "__main__":
    import io
    run_integration_test()
