from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
import easyocr
from pdf2image import convert_from_bytes
from PIL import Image
from io import BytesIO

app = FastAPI(title="🦅 DataFalcon OCR Worker — Cloud OCR (Spanish + Greek + English)")

@app.get("/")
def root():
    return {"status": "online", "engine": "EasyOCR Cloud", "languages": "spa+ell+eng"}

@app.post("/ocr")
async def ocr(file: UploadFile = File(...)):
    try:
        pdf_bytes = await file.read()
        reader = easyocr.Reader(['es', 'el', 'en'], gpu=False)  # lightweight CPU mode

        # Convert PDF pages to images
        images = convert_from_bytes(pdf_bytes, dpi=200)
        text = ""
        for i, img in enumerate(images):
            result = reader.readtext(img, detail=0, paragraph=True)
            text += "\n".join(result) + "\n"

        if not text.strip():
            return JSONResponse({"error": "No text detected"}, status_code=422)

        return JSONResponse({"text": text})

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
