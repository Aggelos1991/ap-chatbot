
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import base64, io, os, pdfplumber, openpyxl, json
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
    content: str  # base64 string

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

        # 2️⃣ Decode file and extract text
        file_bytes = base64.b64decode(content_b64)
        text = extract_text_from_file(file_bytes, filename)
        upper_text = text.upper()

        # 3️⃣ Fast local detection before GPT call
        if any(k in upper_text for k in ["ANDALUSIA", "IKOS ANDALUSIA"]):
            return JSONResponse({"keyword": "ANDALUSIA", "error": ""})
        elif any(k in upper_text for k in ["PORTO PETRO", "IKOS PORTO"]):
            return JSONResponse({"keyword": "PORTO PETRO", "error": ""})
        elif "IKOS SPANISH" in upper_text or "SPANISH HOTEL MANAGEMENT" in upper_text:
            return JSONResponse({"keyword": "IKOS SPANISH HOTEL MANAGEMENT", "error": ""})

        # 4️⃣ Fallback — LLM classification if not found
        prompt = f"""
You are a strict JSON classifier for hotel identification.

Analyze the following text and detect which known entity it belongs to.
Return ONE valid JSON object with two fields: "keyword" and "error".
Never include explanations or text outside JSON.

Possible keyword values (case-insensitive, accept fuzzy mentions like 'Ikos Andalusia', 'Andalusía', etc.):
- "ANDALUSIA"
- "PORTO PETRO"
- "IKOS SPANISH HOTEL MANAGEMENT"
- "OTHER"  (use only if none of the above appear)

Return format example:
{{"keyword": "ANDALUSIA", "error": ""}}

Text to analyze:
{text[:3500]}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )

        raw = response.choices[0].message.content.strip()

        try:
            parsed = json.loads(raw)
            keyword = parsed.get("keyword", "OTHER").upper()
            error = parsed.get("error", "")
        except Exception:
            result = raw.upper()
            if "ANDALUSIA" in result:
                keyword = "ANDALUSIA"
            elif "PORTO" in result:
                keyword = "PORTO PETRO"
            elif "SPANISH" in result:
                keyword = "IKOS SPANISH HOTEL MANAGEMENT"
            else:
                keyword = "OTHER"
            error = f"Unstructured response: {raw}"

        return JSONResponse({"keyword": keyword, "error": error})

    except Exception as e:
        return JSONResponse({"keyword": "OTHER", "error": str(e)})

# ==========================
# HEALTH CHECK
# ==========================
@app.get("/ping")
async def ping():
    return JSONResponse({"status": "ok", "message": "Server reachable"})
