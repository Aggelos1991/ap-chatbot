import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No OpenAI API key found. Add it to .env or Streamlit Secrets.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(value):
    """Normalize decimals like 1.234,56 → 1234.56 and handle negatives"""
    if not value:
        return ""
    s = str(value).strip().replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        num = float(s)
        return round(num, 2)
    except:
        return ""

def extract_raw_lines(uploaded_pdf):
    """Extract all text lines from every page of the PDF."""
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]?\d{0,2}", line):
                    all_lines.append(" ".join(line.split()))
    return all_lines

# ==========================================================
# GPT EXTRACTOR — Enhanced Document Number Detection + Negatives
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT to detect Debit (DEBE) and Credit (HABER) from vendor statements."""
    BATCH_SIZE = 150
    all_records = []
    
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        
        prompt = """You are an expert accountant fluent in SPANISH, GREEK, and accounting terminology.
You are reading extracted lines from a vendor statement (bank statement, AP statement, etc.).

## DOCUMENT NUMBER DETECTION - CRITICAL
Find document numbers in these formats/labels (prioritize in this order):
1. **Spanish**: Nº, Num, Número, Documento, Factura, Fra, Ref, Referencia, Fact, Fatura
2. **Greek**: Τιμολόγιο (Timologio), Αριθμός (Arithmos), Αρ., Νο., Παραστατικό, Τ/Λ, ΤΛ
3. **Common**: Invoice #, DOC, ID, RefNo
4. **Numbers alone**: 1-3 digits followed by dashes/dots or 6+ digits (e.g., 123, 123-45, 2024/001)

## TRANSACTION COLUMNS
- **DEBE**: Debit/Invoice amount (Fra. emitida, Cargo)
- **HABER**: Credit/Payment amount (Cobro, Pago, Abono)
- **SALDO**: Running balance (ignore for extraction)
- **CONCEPTO**: Description

## NEGATIVE NUMBER RULE - IMPORTANT
- **Negative in DEBE** → Move to CREDIT and classify as "Credit Note"
- **Negative in HABER** → Move to DEBIT and classify as "Invoice"

## CLASSIFICATION RULES
1. **Invoice**: DEBE > 0 OR contains "Fra", "Factura", "Τιμολόγιο", "emitida"
2. **Payment**: HABER > 0 OR contains "Cobro", "Pago", "Είσπραξη", "Επιταγή"  
3. **Credit Note**: Contains "NC", "Nota Credito", "Ακυρωτικό", "Πιστωτικό", "Abono" OR NEGATIVE DEBE

## OUTPUT FORMAT - EXACTLY
Return ONLY a valid JSON array. Each object:
```json
{
  "Alternative Document": "EXACT document number found (e.g. 'FRA-2024-001', '12345', 'ΤΛ 67890')",
  "Date": "dd/mm/yyyy OR dd/mm/yy OR empty string",
  "Reason": "Invoice|Payment|Credit Note",
  "Debit": "numeric value OR empty string",
  "Credit": "numeric value OR empty string",
  "Description": "short description of transaction"
}
