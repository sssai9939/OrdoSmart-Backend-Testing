import os
from pathlib import Path
from dotenv import load_dotenv
from supabase import create_client, Client

load_dotenv()
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
BUCKET_NAME = os.getenv("BUCKET_NAME")

# --- Initialize Supabase Client ---
if not all([SUPABASE_URL, SUPABASE_KEY, BUCKET_NAME]):
    print("❌ Critical Error: SUPABASE_URL, SUPABASE_KEY, and BUCKET_NAME must be set in .env file.")
    exit()

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def download_order(order_id: int, local_dir: Path) -> Path | None:
    """Downloads an order file from Supabase Storage."""
    file_name = f"wempy_order_{order_id}.docx"
    local_path = local_dir / file_name
    
    try:
        res = supabase.storage.from_(BUCKET_NAME).download(file_name)
        with open(local_path, "wb+") as f:
            f.write(res)
        print(f"✅ Found and downloaded {file_name}")
        return local_path
    except Exception as e:
        # This is expected when an order doesn't exist yet (404 Not Found).
        # We only log other, unexpected errors.
        if 'Object not found' not in str(e):
             print(f"\nAn unexpected error occurred while downloading order {order_id}: {e}")
        return None
