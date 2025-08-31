
# store orders in supabase
from fastapi import FastAPI, Body
from docx import Document
import uuid
import io
from supabase import create_client, Client
import os
from dotenv import load_dotenv

load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

app = FastAPI()

@app.post("/create-docx/")
def create_docx(text: str = Body(..., embed=True)):

    file_name = f"{uuid.uuid4()}.docx"

    # إنشاء ملف Word في الذاكرة
    buffer = io.BytesIO()
    doc = Document()
    doc.add_paragraph(text)
    doc.save(buffer)
    buffer.seek(0)  # إعادة المؤشر لبداية الملف

    # رفع مباشرة إلى Supabase (بدون حفظ محلي)
    supabase.storage.from_("ordosmart_test").upload(file_name, buffer.getvalue())

    # رابط عام للملف
    public_url = f"{SUPABASE_URL}/storage/v1/public/ordosmart_test/{file_name}"
    return {"message": "File created & uploaded", "filename": file_name, "url": public_url}


# download orders from supabase
import os
import requests
from fastapi import FastAPI

app = FastAPI()

DOWNLOADS_DIR = "downloads"
os.makedirs(DOWNLOADS_DIR, exist_ok=True)

@app.get("/download-docx/")
def download_docx(url: str):
    """
    ينزل ملف docx من رابط Supabase ويحفظه محليًا
    """
    response = requests.get(url)
    if response.status_code != 200:
        return {"error": f"فشل التحميل: {response.status_code}"}

    # استخرج اسم الملف من الرابط
    file_name = url.split("/")[-1]
    file_path = os.path.join(DOWNLOADS_DIR, file_name)

    # احفظ الملف محليًا
    with open(file_path, "wb") as f:
        f.write(response.content)

    return {"message": "تم التحميل والحفظ", "local_path": file_path}
