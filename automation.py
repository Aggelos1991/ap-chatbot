from fastapi import FastAPI
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import base64
import openai
import pdfplumber
import openpyxl
import io

# ==========================
# CONFIGURATION
# ==========================
import os
from openai import OpenAI

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


app = FastAPI()

# ==========================
# HELPER: Extract text from file
# ==========================
def extract_text_from_file(file_bytes: bytes, filename: str) -> str:
    filename = filename.lower()
    text = ""
    try:
        if filename.endswith(".pdf"):
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages[:5]:  # limit to first 5 pages
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
    return text[:4000]  # keep it short for GPT

# ==========================
# MODEL
# ==========================
class FilePayload(BaseModel):
    filename: str
    content: str  # base64 content

# ==========================
# ROUTE
# ==========================
@app.post("/analyze")
async def analyze(payload: FilePayload):
    file_bytes = base64.b64decode(payload.content)
    text = extract_text_from_file(file_bytes, payload.filename)

    prompt = f"""
    You are a classifier. Identify which of these appears in the text:
    ANDALUSIA, PORTO PETRO, IKOS SPANISH HOTEL MANAGEMENT, or OTHER.
    Return only one of those words. No explanations.

    Text:
    {text}
    """

    try:
        response = openai.ChatCompletion.create(
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

        return JSONResponse({"keyword": keyword})
    except Exception as e:
        return JSONResponse({"error": str(e)})
