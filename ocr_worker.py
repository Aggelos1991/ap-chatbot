from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from pdf2image import convert_from_bytes
import pytesseract
from PIL import Image
from io import BytesIO

app = FastAPI(title="DataFalcon OCR Worker — Cloud Edition (Portable Tesseract)")

@app.get("/")
def root():
    return {"status": "online", "engine": "Tesseract Portable", "languages": "spa+ell+eng"}

@app.post("/ocr")
async def ocr(file: UploadFile = File(...)):
    try:
        pdf_bytes = await file.read()

        # Convert each PDF page to image
        images = convert_from_bytes(pdf_bytes, dpi=200, fmt="png")
        text = ""

        for i, img in enumerate(images):
            extracted = pytesseract.image_to_string(img, lang="spa+ell+eng", config="--psm 6")
            text += extracted + "\n"

        if not text.strip():
            return JSONResponse({"error": "No text found"}, status_code=422)

        return JSONResponse({"text": text})

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
