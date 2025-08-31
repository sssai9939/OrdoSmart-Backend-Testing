import os
import io
from supabase import create_client, Client
from dotenv import load_dotenv

# --- Supabase Client Initialization ---
load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
BUCKET_NAME = os.getenv("BUCKET_NAME")

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# --- Supabase Storage Functions ---

def upload_order_to_supabase(file_content: bytes, file_name: str) -> str:
    """Uploads a file's byte content to Supabase Storage."""
    try:
        # The `upload` method in supabase-py expects bytes, which is what we have.
        supabase.storage.from_(BUCKET_NAME).upload(file_name, file_content)
        
        # Construct the public URL manually
        public_url = f"{SUPABASE_URL}/storage/v1/object/public/{BUCKET_NAME}/{file_name}"
        return public_url
    except Exception as e:
        print(f"Failed to upload {file_name} to Supabase: {e}")
        raise

def download_order_from_supabase(file_name: str) -> bytes:
    """Downloads a file from Supabase Storage and returns its content as bytes."""
    try:
        response = supabase.storage.from_(BUCKET_NAME).download(file_name)
        return response
    except Exception as e:
        print(f"Failed to download {file_name} from Supabase: {e}")
        raise
