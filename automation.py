from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import base64, io, os, json, re, pdfplumber, openpyxl, fitz, pytesseract
from PIL import Image
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
app = FastAPI()

class FilePayload(BaseModel):
    filename: str
    content: str

@app.middleware("http")
async def allow_chunked_requests(request: Request, call_next):
    if request.headers.get("transfer-encoding", "").lower() == "chunked":
        body = await request.body()
        request._body = body
    return await call_next(request)

def extract_text_from_pdf_all(file_bytes: bytes) -> str:
    """Combine text from pdfplumber, PyMuPDF, and OCR for maximum coverage."""
    text = ""

    # pdfplumber text
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages[:5]:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception:
        pass

    # PyMuPDF text
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            text += page.get_text("text") + "\n"
    except Exception:
        pass

    # OCR fallback
    if len(text.strip()) < 40:
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for i in range(min(3, len(doc))):
                pix = doc[i].get_pixmap(dpi=200)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text += pytesseract.image_to_string(img, lang="eng") + "\n"
        except Exception:
            pass

    return text

def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    if filename.lower().endswith(".pdf"):
        return extract_text_from_pdf_all(file_bytes)
    elif filename.lower().endswith(".xlsx"):
        text = ""
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(values_only=True):
                text += " ".join([str(c) for c in row if c]) + "\n"
        return text
    else:
        return file_bytes.decode(errors="ignore")

def normalize(txt: str) -> str:
    txt = txt.upper()
    txt = re.sub(r"[ÁÀÂÃÄ]", "A", txt)
    txt = re.sub(r"[ÉÈÊË]", "E", txt)
    txt = re.sub(r"[ÍÌÎÏ]", "I", txt)
    txt = re.sub(r"[ÓÒÔÕÖ]", "O", txt)
    txt = re.sub(r"[ÚÙÛÜ]", "U", txt)
    txt = txt.replace("\xa0", " ").replace("\u00a0", " ")
    txt = re.sub(r"[^A-Z0-9 ]+", " ", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()

@app.post("/analyze")
async def analyze(request: Request):
    try:
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
        raw_text = extract_text_from_file(file_bytes, filename)
        text = normalize(raw_text)

        print("=== DEBUG SAMPLE ===")
        print(text[:300])

        # === STRONG LOCAL DETECTION ===
        if re.search(r"IKOS\s*ANDALUSIA|ANDALUCIA", text):
            return JSONResponse({"keyword": "ANDALUSIA", "error": ""})
        if re.search(r"IKOS\s*PORTO\s*PETRO|PORTOPETRO", text):
            return JSONResponse({"keyword": "PORTO PETRO", "error": ""})
        if re.search(r"IKOS\s*SPANISH\s*HOTEL|SPANISH\s*HOTEL\s*MANAGEMENT", text):
            return JSONResponse({"keyword": "IKOS SPANISH HOTEL MANAGEMENT", "error": ""})

        # === FALLBACK: GPT classification ===
        prompt = f"""
You are a JSON classifier that detects which IKOS hotel this text refers to.

Return only this JSON:
{{
  "keyword": "ANDALUSIA" | "PORTO PETRO" | "IKOS SPANISH HOTEL MANAGEMENT" | "OTHER",
  "error": ""
}}

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
            if "ANDALUSIA" in raw.upper():
                keyword = "ANDALUSIA"
            elif "PORTO" in raw.upper():
                keyword = "PORTO PETRO"
            elif "SPANISH" in raw.upper():
                keyword = "IKOS SPANISH HOTEL MANAGEMENT"
            else:
                keyword = "OTHER"
            error = f"LLM raw output: {raw}"

        return JSONResponse({"keyword": keyword, "error": error})

    except Exception as e:
        return JSONResponse({"keyword": "OTHER", "error": str(e)})

@app.get("/ping")
async def ping():
    return JSONResponse({"status": "ok", "message": "Server reachable"})
