from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from pdf2image import convert_from_bytes
from io import BytesIO
import easyocr
from PIL import Image

app = FastAPI(
    title="🦅 DataFalcon OCR Worker",
    description="Cloud OCR service for Spanish, Greek, and English PDFs",
    version="1.0"
)

@app.get("/")
def root():
    return {"status": "online", "engine": "EasyOCR", "languages": "spa+ell+eng"}

@app.post("/ocr")
async def ocr(file: UploadFile = File(...)):
    try:
        pdf_bytes = await file.read()

        # Initialize OCR only once (prevent Render from hanging)
        reader = easyocr.Reader(['es', 'el', 'en'], gpu=False)

        # Convert PDF to images
        images = convert_from_bytes(pdf_bytes, dpi=180)
        all_text = []

        for i, img in enumerate(images):
            results = reader.readtext(img, detail=0, paragraph=True)
            all_text.extend(results)

        if not all_text:
            return JSONResponse({"error": "No text detected"}, status_code=422)

        return {"text": "\n".join(all_text)}

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
