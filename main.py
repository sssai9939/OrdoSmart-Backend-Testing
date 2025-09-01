import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
load_dotenv()

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from helpers.order_models import OrderRequest
from controllers.order_generator import create_order_docx_bytes
from controllers.order_manager import get_next_order_id, build_order_docx_path
from controllers.supabase_uploader import upload_order_to_supabase, download_order_from_supabase
from models.enums import ResponseMessages
from supabase import create_client, Client

# Project root (directory that contains main.py)
ROOT = Path(__file__).resolve().parent

# Get custom orders path from environment variable, default to ROOT / "orders"
ORDERS_PATH = Path(os.getenv("ORDERS_PATH", str(ROOT / "orders")))

# --- Supabase Client Initialization ---
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("Supabase URL and Key must be set in environment variables.")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

app = FastAPI(title=ResponseMessages.RESTAURANT_TITLE.value)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- API Endpoints ---

@app.post("/submit_order")
def submit_order(order: OrderRequest):
    try:
        order_id = get_next_order_id(ORDERS_PATH)
        file_name = f"wempy_order_{order_id}.docx"

        # 1. Create DOCX file in memory
        docx_bytes = create_order_docx_bytes(order, order_id)

        # 2. Upload to Supabase Storage
        public_url = upload_order_to_supabase(file_content=docx_bytes, file_name=file_name)

        # 3. Insert into 'orders' table to trigger realtime
        try:
            supabase.table("orders").insert({"id": order_id, "status": "new"}).execute()
        except Exception as db_error:
            # Log the DB error but don't fail the whole request, as the order is already in storage
            print(f"⚠️ Warning: Failed to insert order {order_id} into DB for realtime update: {db_error}")

        return JSONResponse(
            status_code=200,
            content={"success": True, "order_id": order_id, "file_name": file_name, "url": public_url}
        )
    except Exception as e:
        error_message = f"{ResponseMessages.PROCESSING_ORDER_FAILED.value} : {e}"
        print(error_message)
        return JSONResponse(
            status_code=500,
            content={"success": False, "message": error_message}
        )
