import os, re
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Strict Referencia Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Strict Referencia Version (NO GPT)")


# ==========================================================
# PDF TEXT EXTRACTION (OCR FALLBACK)
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()

            if text:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean and "saldo" not in clean.lower():
                        all_lines.append(clean)
            else:
                try:
                    images = convert_from_bytes(pdf_bytes, dpi=260, first_page=idx, last_page=idx)
                    ocr_text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")
                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if clean and "saldo" not in clean.lower():
                            all_lines.append(clean)
                except:
                    pass

    return all_lines


# ==========================================================
# STRICT REFERENCIA EXTRACTOR
# ==========================================================
def extract_referencia(line):
    """Return ONLY official Referencia. If not found → empty cell."""
    matches = re.findall(r"\b\d{12,18}\b", line)
    return matches[0] if matches else ""


# ==========================================================
# NORMALIZE AMOUNTS
# ==========================================================
def normalize_amount(v):
    if not v:
        return ""
    v = v.replace(".", "").replace(",", ".")
    v = re.sub(r"[^\d\.\-]", "", v)
    try:
        return round(float(v), 2)
    except:
        return ""


# ==========================================================
# PARSE A LEDGER LINE
# ==========================================================
def parse_ledger_line(line):
    """
    Expected structure:
    Fecha | Asiento | Documento | Libro | Descripción (variable) | Referencia | F.Valor | Debe | Haber | Saldo (ignored)
    """

    parts = line.split()

    if len(parts) < 8:
        return None

    # Fecha = always first token
    fecha = parts[0]

    # Find Referencia token = ONLY token with
