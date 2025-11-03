
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
import pytesseract
import pdfplumber
from PIL import Image
from io import BytesIO

app = FastAPI(title="DataFalcon OCR Worker")

@app.get("/")
def root():
    return {"status": "running", "message": "🦅 DataFalcon OCR Worker online"}

@app.post("/ocr")
async def ocr(file: UploadFile = File(...)):
    try:
        pdf_bytes = await file.read()
        text = ""

        # Try to extract text directly (no Poppler)
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"

                # If no text, OCR the image
                if not page_text.strip():
                    im = page.to_image(resolution=300).original
                    text += pytesseract.image_to_string(im, lang="spa+ell+eng") + "\n"

        if not text.strip():
            return JSONResponse({"error": "No text extracted from PDF"}, status_code=422)

        return JSONResponse({"text": text})

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
