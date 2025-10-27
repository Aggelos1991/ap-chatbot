from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import base64, io, os, pdfplumber, openpyxl
from openai import OpenAI
from dotenv import load_dotenv

# ==========================
# SETUP
# ==========================
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
app = FastAPI()

# ==========================
# MODEL
# ==========================
class FilePayload(BaseModel):
    filename: str
    content: str  # base64

# ==========================
# MIDDLEWARE (Power Automate compatibility)
# ==========================
@app.middleware("http")
async def allow_chunked_requests(request: Request, call_next):
    if request.headers.get("transfer-encoding", "").lower() == "chunked":
        body = await request.body()
        request._body = body
    return await call_next(request)

# ==========================
# FILE PARSING HELPER
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
        # 1️⃣ Read JSON payload safely
        try:
            data = await request.json()
        except Exception:
            body = await request.body()
            return JSONResponse({"keyword": "OTHER", "error": f"Invalid JSON: {body[:100]}"})

        filename = data.get("filename", "unknown")
        content_b64 = data.get("content", "")

        if not content_b64:
            return JSONResponse({"keyword": "OTHER", "error": "Empty content"})

        # 2️⃣ Decode file
        file_bytes = base64.b64decode(content_b64)
        text = extract_text_from_file(file_bytes, filename)

        # 3️⃣ Classify text
        prompt = f"""
        You are a classifier. Identify which of these appears in the text:
        ANDALUSIA, PORTO PETRO, IKOS SPANISH HOTEL MANAGEMENT, or OTHER.
        Return only one of those words. No explanations.

        Text:
        {text}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )

        result = response.choices[0].message.content.strip().upper()
        if "ANDALUSIA" in result:
            keyword = "ANDALUSIA"
        elif "PORTO" in result:
            keyword = "PORTO PETRO"
        elif "SPANISH" in result:
            keyword = "IKOS SPANISH HOTEL MANAGEMENT"
        else:
            keyword = "OTHER"

        # 4️⃣ Always return JSON — never plain text
        return JSONResponse({"keyword": keyword})

    except Exception as e:
        # 5️⃣ Failsafe — make sure *any* exception still returns valid JSON
        return JSONResponse({"keyword": "OTHER", "error": str(e)})

# ✅ Test endpoint
@app.get("/ping")
async def ping():
    return JSONResponse({"status": "ok", "message": "Server reachable"})
