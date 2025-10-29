from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import base64, io, os, json, re, pdfplumber, openpyxl, fitz, pytesseract, hashlib, time
from PIL import Image
from openai import OpenAI
from dotenv import load_dotenv

# ================== SETUP =====================
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
app = FastAPI()

# cache for duplicate prevention
recent_hashes = {}  # {hash: timestamp}
CACHE_TTL = 300  # 5 minutes

class FilePayload(BaseModel):
    filename: str
    content: str


@app.middleware("http")
async def allow_chunked_requests(request: Request, call_next):
    if request.headers.get("transfer-encoding", "").lower() == "chunked":
        body = await request.body()
        request._body = body
    return await call_next(request)


# ================== HELPERS =====================
def extract_text_from_pdf_all(file_bytes: bytes) -> str:
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception as e:
        print(f"[PDFPlumber Error] {e}")

    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            text += page.get_text("text") + "\n"
        doc.close()
    except Exception as e:
        print(f"[PyMuPDF Error] {e}")

    if len(text.strip()) < 50:
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for i in range(min(10, len(doc))):
                pix = doc[i].get_pixmap(dpi=200)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                ocr = pytesseract.image_to_string(img, lang="eng+spa")
                text += ocr + "\n"
            doc.close()
        except Exception as e:
            print(f"[OCR Error] {e}")
    return text.strip()


def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    filename = filename.lower()
    if filename.endswith(".pdf"):
        return extract_text_from_pdf_all(file_bytes)
    elif filename.endswith((".xlsx", ".xls")):
        text = ""
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    row_text = " ".join([str(cell) for cell in row if cell not in (None, "")])
                    if row_text:
                        text += row_text + "\n"
        except Exception as e:
            print(f"[Excel Error] {e}")
        return text
    else:
        return file_bytes.decode("utf-8", errors="ignore")


def normalize(txt: str) -> str:
    txt = txt.upper()
    txt = re.sub(r"[ÁÀÂÃÄÅ]", "A", txt)
    txt = re.sub(r"[ÉÈÊË]", "E", txt)
    txt = re.sub(r"[ÍÌÎÏ]", "I", txt)
    txt = re.sub(r"[ÓÒÔÕÖØ]", "O", txt)
    txt = re.sub(r"[ÚÙÛÜ]", "U", txt)
    txt = re.sub(r"[Ñ]", "N", txt)
    txt = txt.replace("Ç", "C")
    txt = txt.replace("\xa0", " ").replace("\u00a0", " ")
    txt = re.sub(r"[^A-Z0-9\s]", " ", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


def detect_ikos_hotel(text: str) -> str:
    norm = normalize(text)
    if "IKOS" in norm and any(kw in norm for kw in ["ANDALUSIA", "ANDALUCIA"]):
        return "ANDALUSIA"
    if "IKOS" in norm and "PORTO PETRO" in norm:
        return "PORTO PETRO"
    if any(kw in norm for kw in ["IKOS SPANISH HOTEL MANAGEMENT", "ISHM", "IKOS HOTELS SPAIN"]):
        return "IKOS SPANISH HOTEL MANAGEMENT"
    return None


# ================== MAIN ROUTE =====================
@app.post("/analyze")
async def analyze(request: Request):
    try:
        # Parse input safely
        try:
            data = await request.json()
        except Exception:
            body = await request.body()
            return JSONResponse({"keyword": "OTHER", "error": "Invalid JSON"})

        filename = data.get("filename", "unknown")
        content_b64 = data.get("content", "")

        if not content_b64:
            return JSONResponse({"keyword": "OTHER", "error": "Empty content"})

        # ✅ Prevent infinite "Apply to each" — deduplicate based on file hash
        file_hash = hashlib.md5(content_b64.encode()).hexdigest()
        now = time.time()

        # remove expired entries
        for k, v in list(recent_hashes.items()):
            if now - v > CACHE_TTL:
                del recent_hashes[k]

        if file_hash in recent_hashes:
            return JSONResponse({"keyword": "OTHER", "error": "Duplicate request ignored"})

        recent_hashes[file_hash] = now

        # Decode and analyze
        file_bytes = base64.b64decode(content_b64)
        raw_text = extract_text_from_file(file_bytes, filename)
        text = normalize(raw_text)
        print(f"\n[DEBUG] File: {filename} | Hash: {file_hash}")
        print(f"[DEBUG] Text length: {len(text)}")

        # Filename hints
        filename_lower = filename.lower()
        if any(kw in filename_lower for kw in ["andalusia", "odisia", "estepona"]):
            return JSONResponse({"keyword": "ANDALUSIA", "error": ""})
        if any(kw in filename_lower for kw in ["porto petro", "portopetro", "mallorca"]):
            return JSONResponse({"keyword": "PORTO PETRO", "error": ""})

        # Text detection
        local_result = detect_ikos_hotel(text)
        if local_result:
            return JSONResponse({"keyword": local_result, "error": ""})

        # Fallback to GPT if needed
        if len(text) > 50:
            prompt = f"""
You are a strict JSON classifier. Return ONLY valid JSON.
Rules:
- "ANDALUSIA" → if mentions: ANDALUSIA or ANDALUCIA
- "PORTO PETRO" → if mentions: PORTO PETRO
- "IKOS SPANISH HOTEL MANAGEMENT" → if mentions: IKOS SPANISH HOTEL MANAGEMENT or ISHM
- Otherwise → "OTHER"
Text:
{text[:3000]}
Return ONLY:
{{"keyword": "ANDALUSIA" | "PORTO PETRO" | "IKOS SPANISH HOTEL MANAGEMENT" | "OTHER"}}
"""
            try:
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0,
                    response_format={"type": "json_object"}
                )
                raw = response.choices[0].message.content.strip()
                parsed = json.loads(raw)
                keyword = parsed.get("keyword", "OTHER").upper()
                valid = ["ANDALUSIA", "PORTO PETRO", "IKOS SPANISH HOTEL MANAGEMENT", "OTHER"]
                if keyword not in valid:
                    keyword = "OTHER"
                return JSONResponse({"keyword": keyword, "error": ""})
            except Exception as e:
                return JSONResponse({"keyword": "OTHER", "error": f"LLM failed: {str(e)[:80]}"})

        return JSONResponse({"keyword": "OTHER", "error": "Empty or unreadable file"})

    except Exception as e:
        return JSONResponse({"keyword": "OTHER", "error": f"Server error: {str(e)[:100]}"})

# ================== HEALTH CHECK =====================
@app.get("/ping")
async def ping():
    return JSONResponse({"status": "ok", "message": "Server reachable"})
