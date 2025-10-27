from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import base64
import io
import os
import pdfplumber
import openpyxl
from openai import OpenAI
from dotenv import load_dotenv

# ==========================
# SETUP
# ==========================
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
app = FastAPI()

# ==========================
# MIDDLEWARE: handle chunked transfer (Power Automate)
# ==========================
@app.middleware("http")
async def allow_chunked_requests(request: Request, call_next):
    if request.headers.get("transfer-encoding", "").lower() == "chunked":
        body = await request.body()
        request._body = body
    response = await call_next(request)
    return response

# ==========================
# MODEL
# ==========================
class FilePayload(BaseModel):
    filename: str
    content: str  # base64 content

# ==========================
# HELPER
# ==========================
def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    text = ""
    filename = filename.lower()
    try:
        if filename.endswith(".pdf"):
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages[:3]:
                    text += page.extract_text() or ""
        elif filename.endswith(".xlsx"):
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    text += " ".join([str(cell) for cell in row if cell]) + "\n"
        else:
            text = file_bytes.decode(errors="ignore")
    except Exception as e:
        text = f"Error reading file: {e}"
    return text[:4000]

# ==========================
# ROUTE
# ==========================
@app.post("/analyze")
async def analyze(request: Request):
    try:
        print("🔵 Received POST /analyze")

        data = await request.json()
        filename = data.get("filename", "unknown")
        content_b64 = data.get("content", "")

        if not content_b64:
            print("⚠️ No content provided")
            return JSONResponse(content={"keyword": "OTHER", "error": "Empty content"})

        # Decode file
        file_bytes = base64.b64decode(content_b64)
        text = extract_text_from_file(file_bytes, filename)

        # Prepare prompt
        prompt = f"""
        You are a classifier. Identify which of these appears in the text:
        ANDALUSIA, PORTO PETRO, IKOS SPANISH HOTEL MANAGEMENT, or OTHER.
        Return only one of those words. No explanations.

        Text:
        {text}
        """

        # GPT classification
        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}]
            )
            result = response.choices[0].message.content.strip().upper()
        except Exception as gpt_error:
            print(f"⚠️ GPT error: {gpt_error}")
            result = ""

        # Simple rule-based fallback
        if "ANDALUSIA" in result:
            keyword = "ANDALUSIA"
        elif "PORTO" in result:
            keyword = "PORTO PETRO"
        elif "SPANISH" in result:
            keyword = "IKOS SPANISH HOTEL MANAGEMENT"
        elif "ANDALUSIA" in text.upper():
            keyword = "ANDALUSIA"
        elif "PORTO" in text.upper():
            keyword = "PORTO PETRO"
        elif "SPANISH" in text.upper():
            keyword = "IKOS SPANISH HOTEL MANAGEMENT"
        else:
            keyword = "OTHER"

        print(f"✅ Classification result: {keyword}")
        return JSONResponse(content={"keyword": keyword})

    except Exception as e:
        print(f"❌ Error: {e}")
        return JSONResponse(content={"keyword": "OTHER", "error": str(e)})

# ==========================
# PING (optional test endpoint)
# ==========================
@app.get("/ping")
async def ping():
    return JSONResponse(content={"status": "ok", "message": "Server reachable"})
