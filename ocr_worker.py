from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
import easyocr
import pdfplumber
from PIL import Image
from io import BytesIO

app = FastAPI(title="DataFalcon OCR Worker — EasyOCR Cloud Edition")

@app.get("/")
def root():
    return {"status": "online", "engine": "EasyOCR (spa+ell+eng)"}

@app.post("/ocr")
async def ocr(file: UploadFile = File(...)):
    try:
        pdf_bytes = await file.read()
        text = ""
        reader = easyocr.Reader(['es', 'el', 'en'], gpu=False)

        # Try pdfplumber first
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"

                # Fallback to OCR for image-based pages
                if not page_text.strip():
                    im = page.to_image(resolution=300).original
                    result = reader.readtext(im, detail=0, paragraph=True)
                    text += "\n".join(result) + "\n"

        if not text.strip():
            return JSONResponse({"error": "No text extracted"}, status_code=422)

        return JSONResponse({"text": text})

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
