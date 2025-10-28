from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import base64, io, os, pdfplumber, openpyxl, json, fitz, pytesseract
from PIL import Image
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
# MIDDLEWARE (for Power Automate)
# ==========================
@app.middleware("http")
async def allow_chunked_requests(request: Request, call_next):
    if request.headers.get("transfer-encoding", "").lower() == "chunked":
        body = await request.body()
        request._body = body
    return await call_next(request)

# ==========================
# FILE PARSER
# ==========================
def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    text = ""
    filename = filename.lower()

    try:
        # --- PDF FILES ---
        if filename.endswith(".pdf"):
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages[:3]:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"

            # OCR fallback
            if len(text.strip()) < 30:
                doc = fitz.open(stream=file_bytes, filetype="pdf")
                for i in range(min(3, len(doc))):
                    pix = doc[i].get_pixmap(dpi=200)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    ocr_text = pytesseract.image_to_string(img, lang="eng")
                    text += ocr_text + "\n"

        # --- EXCEL FILES ---
        elif filename.endswith(".xlsx"):
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    text += " ".join([str(c) for c in row if c]) + "\n"

        # --- TXT / OTHER ---
        else:
            text = file_bytes.decode(errors="ignore")

    except Exception as e:
        text = f"Error reading file: {e}"

    return text[:6000]

# ==========================
# ROUTE
# ==========================
@app.post("/analyze")
async def analyze(request: Request):
    try:
        # --- Read payload ---
        try:
            data = await request.json()
        except Exception:
            body = await request.body()
            return JSONResponse({"keyword": "OTHER", "error": f"Invalid JSON: {body[:100]}"})

        filename = data.get("filename", "unknown")
        content_b64 = data.get("content", "")
        if not content_b64:
            return JSONResponse({"keyword": "OTHER", "error": "Empty content"})

        file_bytes = base64.b64decode(content_b64)
        text = extract_text_from_file(file_bytes, filename)
        upper_text = text.upper().replace("Í", "I").replace("Á", "A").replace("É", "E")

        # --- Fast Local Match ---
        if "ANDALUSIA" in upper_text or "IKOS ANDALUSIA" in upper_text or "ANDALUCIA" in upper_text:
            return JSONResponse({"keyword": "ANDALUSIA", "error": ""})
        elif any(k in upper_text for k in ["PORTO PETRO", "IKOS PORTO", "PORTOPETRO"]):
            return JSONResponse({"keyword": "PORTO PETRO", "error": ""})
        elif any(k in upper_text for k in ["IKOS SPANISH", "SPANISH HOTEL MANAGEMENT"]):
            return JSONResponse({"keyword": "IKOS SPANISH HOTEL MANAGEMENT", "error": ""})

        # --- GPT Classification ---
        prompt = f"""
You are a strict JSON classifier that identifies which IKOS hotel a text belongs to.

Analyze the text below and output **only** one valid JSON object in this exact format:
{{
  "keyword": "ANDALUSIA" | "PORTO PETRO" | "IKOS SPANISH HOTEL MANAGEMENT" | "OTHER",
  "error": ""
}}

Rules:
- Classify as "ANDALUSIA" if you see "Ikos Andalusia", "Andalucía", or similar.
- Classify as "PORTO PETRO" if you see "Ikos Porto Petro", "PortoPetro", or "Ikos Porto".
- Classify as "IKOS SPANISH HOTEL MANAGEMENT" if you see "Ikos Spanish Hotels Management".
- Otherwise, use "OTHER".

Text:
{text[:3500]}
        """

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
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
            error = f"Unstructured LLM output: {raw}"

        return JSONResponse({"keyword": keyword, "error": error})

    except Exception as e:
        return JSONResponse({"keyword": "OTHER", "error": str(e)})

# ==========================
# HEALTH CHECK
# ==========================
@app.get("/ping")
async def ping():
    return JSONResponse({"status": "ok", "message": "Server reachable"})
