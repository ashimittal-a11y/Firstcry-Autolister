import os
import json
from supabase import create_client, Client

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

supabase_client: Client | None = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
        print("Connected to Supabase successfully!")
    except Exception as e:
        print(f"Failed to connect to Supabase: {e}")
else:
    print("Supabase credentials not found in env. Falling back to local storage.")

# Local file configuration fallback
MAPPINGS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "mappings")
os.makedirs(MAPPINGS_DIR, exist_ok=True)

def db_load_preset(name: str) -> dict:
    """Loads a mapping preset from Supabase or falls back to local disk."""
    if supabase_client:
        try:
            response = supabase_client.table("presets").select("config").eq("name", name).execute()
            if response.data and len(response.data) > 0:
                return response.data[0]["config"]
        except Exception as e:
            print(f"Error loading preset from Supabase: {e}")
            
    # Fallback to local json file
    file_path = os.path.join(MAPPINGS_DIR, f"{name}.json")
    if os.path.exists(file_path):
        with open(file_path, "r") as f:
            return json.load(f)
    return {"join_keys": {}, "mappings": {}}

def db_save_preset(name: str, config: dict) -> None:
    """Saves a mapping preset to Supabase and saves a backup locally."""
    # Always save a local copy as backup
    file_path = os.path.join(MAPPINGS_DIR, f"{name}.json")
    try:
        with open(file_path, "w") as f:
            json.dump(config, f, indent=4)
    except Exception as e:
        print(f"Failed to write local backup preset: {e}")

    if supabase_client:
        try:
            # Upsert into presets table
            supabase_client.table("presets").upsert({
                "name": name,
                "config": config
            }).execute()
            print(f"Successfully saved preset '{name}' to Supabase.")
        except Exception as e:
            print(f"Failed to save preset to Supabase: {e}")

def db_list_presets() -> list:
    """Lists available presets from Supabase and merges them with local files."""
    presets = set()
    if supabase_client:
        try:
            response = supabase_client.table("presets").select("name").execute()
            if response.data:
                for row in response.data:
                    presets.add(row["name"])
        except Exception as e:
            print(f"Error listing presets from Supabase: {e}")
            
    # Load from local mappings folder
    for f in os.listdir(MAPPINGS_DIR):
        if f.endswith(".json"):
            presets.add(f[:-5])
            
    return sorted(list(presets))

def db_log_run(category: str, skus_count: int, filename: str, status: str) -> None:
    """Logs a generation run metadata to Supabase."""
    if supabase_client:
        try:
            supabase_client.table("runs_history").insert({
                "category": category or "Auto-Detect",
                "skus_count": skus_count,
                "filename": filename or "populated_listing.xlsm",
                "status": status
            }).execute()
            print("Logged run history to Supabase successfully.")
        except Exception as e:
            print(f"Failed to log run to Supabase: {e}")
    else:
        print(f"Local run log: {category} ({skus_count} SKUs) - {filename} - {status}")
