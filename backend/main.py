import os
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import io

from app.mapping_engine import learn_mappings, generate_listings

app = FastAPI(title="Firstcry Autolister Backend")

# Enable CORS for development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mappings directory
MAPPINGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mappings")
os.makedirs(MAPPINGS_DIR, exist_ok=True)

# Mount static files
static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
os.makedirs(static_dir, exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")

def create_default_presets():
    import pandas as pd
    listing_dir = r"C:\Users\Manann\Desktop\FirstcryAutolister\Listing"
    
    files = {
        "clothing_default": os.path.join(listing_dir, "Clothing Old Listing.xlsm"),
        "footwear_default": os.path.join(listing_dir, "Footwear Old Listing.xlsm")
    }
    
    for preset_name, file_path in files.items():
        if not os.path.exists(file_path):
            continue
        try:
            df_raw = pd.read_excel(file_path, sheet_name="Sheet1", header=None, nrows=100)
            if df_raw.shape[0] < 2:
                continue
                
            row0 = df_raw.iloc[0].astype(str).tolist()
            row1 = df_raw.iloc[1].astype(str).tolist()
            
            headers = []
            for r0, r1 in zip(row0, row1):
                s1 = str(r1).strip() if pd.notna(r1) else ''
                s0 = str(r0).strip() if pd.notna(r0) else ''
                h = s1 if s1 not in ['nan', ''] else s0
                headers.append(h)
                
            data = df_raw.iloc[2:].copy()
            data.columns = headers
            
            # Limit headers to columns 0-108
            active_headers = headers[:109]
            
            mappings = {}
            for col in active_headers:
                if not col or col == 'nan' or col.startswith('Unnamed:'):
                    continue
                # Check constant values
                col_data = data[col].dropna().unique()
                col_data = [str(x).strip() for x in col_data if str(x).strip() not in ['nan', '']]
                if len(col_data) == 1:
                    mappings[col] = {"type": "constant", "value": col_data[0]}
                elif col in ['Product Name', 'Product Description']:
                    mappings[col] = {"type": "content", "column": col}
                else:
                    mappings[col] = {"type": "master", "column": col}
            
            preset_data = {
                "join_keys": {
                    "old_listing": "Style Code",
                    "master": "Style Code",
                    "content": "Style Code"
                },
                "mappings": mappings
            }
            
            preset_path = os.path.join(MAPPINGS_DIR, f"{preset_name}.json")
            with open(preset_path, "w") as f:
                json.dump(preset_data, f, indent=4)
            print(f"Created default preset: {preset_name}")
        except Exception as e:
            print(f"Error creating default preset for {preset_name}: {e}")

@app.on_event("startup")
def startup_event():
    create_default_presets()

@app.get("/")
def read_index():
    index_path = os.path.join(static_dir, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return {"message": "Autolister server is running. Please create static/index.html"}

@app.post("/api/learn")
async def api_learn(
    old_listing: UploadFile = File(...),
    item_directory: UploadFile = File(...),
    content_sheet: UploadFile = File(...)
):
    try:
        old_bytes = await old_listing.read()
        item_bytes = await item_directory.read()
        content_bytes = await content_sheet.read()
        
        result = learn_mappings(old_bytes, item_bytes, content_bytes)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to learn mappings: {str(e)}")

@app.post("/api/list")
async def api_list(
    item_directory: UploadFile = File(...),
    content_sheet: UploadFile = File(...),
    template: UploadFile = File(...),
    mapping_config: str | None = Form(default=None),
    item_codes: str | None = Form(default=None),
    category: str | None = Form(default=None)
):
    try:
        item_bytes = await item_directory.read()
        content_bytes = await content_sheet.read()
        template_bytes = await template.read()
        
        config = json.loads(mapping_config) if mapping_config else {"join_keys": {}, "mappings": {}}
        
        codes_list = None
        if item_codes:
            codes_list = [c.strip() for c in item_codes.split("\n") if c.strip()]
            
        output_bytes = generate_listings(item_bytes, content_bytes, template_bytes, config, codes_list, category)
        
        # Log run history metadata
        try:
            from app.db import db_log_run
            skus_count = len(codes_list) if codes_list else 0
            db_log_run(category, skus_count, template.filename if template else "populated_listing.xlsm", "Success")
        except Exception as db_err:
            print(f"Failed to log run metadata: {db_err}")

        # We need to return a file download response retaining the xlsm content type
        filename = template.filename
        if not filename or not filename.endswith(".xlsm"):
            filename = "populated_listing.xlsm"
            
        return StreamingResponse(
            io.BytesIO(output_bytes),
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        import traceback
        print("Error in /api/list:", traceback.format_exc())
        try:
            from app.db import db_log_run
            db_log_run(category, 0, template.filename if template else "populated_listing.xlsm", f"Failed: {str(e)[:100]}")
        except:
            pass
        raise HTTPException(status_code=500, detail=f"Failed to generate listing: {str(e)}")

@app.get("/api/presets")
def get_presets():
    try:
        from app.db import db_list_presets
        return db_list_presets()
    except Exception as e:
        print(f"Error fetching presets: {e}")
        return []

@app.get("/api/presets/{name}")
def get_preset(name: str):
    try:
        from app.db import db_load_preset
        preset = db_load_preset(name)
        return preset
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to load preset: {str(e)}")

@app.post("/api/presets/{name}")
async def save_preset(name: str, config: dict):
    try:
        from app.db import db_save_preset
        db_save_preset(name, config)
        return {"status": "success", "message": f"Preset '{name}' saved successfully."}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save preset: {str(e)}")
