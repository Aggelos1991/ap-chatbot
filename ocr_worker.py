from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2image import convert_from_bytes
import easyocr

# ==========================================================
# 🦅 DataFalcon OCR Worker
# Cloud OCR service for Spanish, Greek, and English PDFs
# ==========================================================

app = FastAPI(
    title="🦅 DataFalcon OCR Worker",
    description="Cloud OCR service for scanned PDFs (Spanish, Greek, English)",
    version="1.0"
)

# Allow access from Streamlit / other origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"]
)

# Initialize EasyOCR reader ONCE at startup
reader = easyocr.Reader(['es', 'el', 'en'], gpu=False)

@app.get("/")
def root():
    """Status endpoint"""
    return {
        "status": "online",
        "engine": "EasyOCR",
        "languages": "spa+ell+eng"
    }

@app.post("/ocr")
async def ocr(file: UploadFile = File(...)):
    """
    Perform OCR on a scanned PDF.
    Returns all text + page-separated results.
    """
    try:
        pdf_bytes = await file.read()
        images = convert_from_bytes(pdf_bytes, dpi=200)

        pages_output = []
        all_text = []

        for i, img in enumerate(images):
            results = reader.readtext(img, detail=0, paragraph=True)
            page_text = "\n".join(results)
            pages_output.append({"page": i + 1, "text": page_text})
            all_text.append(page_text)

        if not any(all_text):
            return JSONResponse({"error": "No text detected"}, status_code=422)

        return {
            "pages": pages_output,
            "text": "\n\n".join(all_text)
        }

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)
