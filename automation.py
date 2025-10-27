from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse, PlainTextResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import base64
import io
import os
import pdfplumber
import openpyxl
from openai import OpenAI
from dotenv import load_dotenv

# ==========================
# ENV + OPENAI CLIENT
# ==========================
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# ==========================
# APP INITIALIZATION
# ==========================
app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================
# FIX 411: CHUNKED REQUESTS
# ==========================
@app.middleware("http")
async def fix_chunked(request: Request, call_next):
    # Power Automate uses Transfer-Encoding: chunked (no Content-Length)
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
# FILE TEXT EXTRACTOR
# ==========================
def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    filename = filename.lower()
    text = ""
    try:
        if filename.endswith(".pdf"):
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages[:5]:
                    text += page.extract_text() or ""
        elif filename.endswith(".xlsx"):
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    row_text = " ".join([str(cell) for cell in row if cell])
                    text += row_text + "\n"
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
        # ✅ Manually read body to guarantee chunked payload is parsed
        body = await request.body()

        # Decode JSON manually instead of relying on Pydantic
        import json
        data = json.loads(body.decode("utf-8"))

        filename = data.get("filename", "unknown.txt")
        content = data.get("content", "")
        file_bytes = base64.b64decode(content)

        text = extract_text_from_file(file_bytes, filename)

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
        elif "SPANISH" in result or "IKOS SPANISH HOTEL" in result:
            keyword = "IKOS SPANISH HOTEL MANAGEMENT"
        else:
            keyword = "OTHER"

        return JSONResponse({"keyword": keyword})

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

# ==========================
# ROOT ENDPOINT
# ==========================
@app.get("/")
def home():
    return PlainTextResponse("✅ FastAPI Analyzer running — ready for Power Automate uploads.")

# ==========================
# LOCAL ENTRYPOINT
# ==========================
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT", 8000)))
