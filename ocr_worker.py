from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from pdf2image import convert_from_bytes
import pytesseract

app = FastAPI()

@app.post("/ocr")
async def ocr(file: UploadFile = File(...)):
    pdf_bytes = await file.read()
    images = convert_from_bytes(pdf_bytes, dpi=200)
    text = ""
    for i, img in enumerate(images):
        text += pytesseract.image_to_string(img, lang="spa+ell+eng") + "\n"
    return JSONResponse({"text": text})
