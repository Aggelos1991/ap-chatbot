import os, re, json
import pdfplumber
import streamlit as st
import pandas as pd
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="DataFalcon — FINAL GPT-SAFE", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL GPT-SAFE VERSION")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
client = OpenAI(api_key=api_key)

MODEL = "gpt-4o-mini"


# ==========================================================
# 1. RAW PDF EXTRACTION
# ==========================================================
def extract_raw(file):
    out = []
    pdf_bytes = file.read()
    file.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for p in pdf.pages:
            text = p.extract_text()
            if text:
                for ln in text.split("\n"):
                    ln = ln.strip()
                    if ln:
                        out.append(ln)
            else:
                img = convert_from_bytes(pdf_bytes, dpi=200, first_page=p.page_number, last_page=p.page_number)[0]
                t = pytesseract.image_to_string(img, lang="spa+eng+ell")
                for ln in t.split("\n"):
                    ln = ln.strip()
                    if ln:
                        out.append(ln)

    return out


# ==========================================================
# 2. MERGE MULTIPLE BROKEN LEDGER ROWS
# ==========================================================
def merge_rows(raw):
    rows = []
    buffer = ""

    date_re = r"^\d{2}/\d{2}/\d{4}"

    for ln in raw:
        if re.match(date_re, ln):
            if buffer:
                rows.append(buffer.strip())
            buffer = ln
        else:
            buffer += " " + ln

    if buffer:
        rows.append(buffer.strip())

    return rows


# ==========================================================
# 3. SPLIT MULTIPLE ENTRIES INSIDE A SINGLE ROW
# ==========================================================
def split_rows(lines):
    final = []
    date_re = r"(?=\d{2}/\d{2}/\d{4})"

    for ln in lines:
        parts = re.split(date_re, ln)
        for p in parts:
            p = p.strip()
            if re.match(r"^\d{2}/\d{2}/\d{4}", p):
                final.append(p)

    return final


# ==========================================================
# 4. PURE PYTHON PARSING (NO GPT)
# ==========================================================
def parse_line(ln):
    # DATE
    m = re.match(r"^(\d{2}/\d{2}/\d{4})", ln)
    date = m.group(1) if m else ""

    # ASIENTO (VEN / GRL / etc)
    asiento = ""
    m = re.search(r"\b(VEN|GRL|APL|DEV|ABN|V)\b", ln)
    if m:
        asiento = m.group(1)

    # REFERENCIA = 9–15 digit numbers
    ref = ""
    m = re.search(r"\b(\d{9,15})\b", ln)
    if m:
        ref = m.group(1)

    # DEBIT / CREDIT (last numbers)
    nums = re.findall(r"[-]?\d{1,3}(?:\.\d{3})*,\d{2}", ln)
    debit = credit = ""

    if len(nums) == 1:
        # only one number
        debit = nums[0]
    elif len(nums) >= 2:
        debit = nums[-2]
        credit = nums[-1]

    return {
        "Date": date,
        "Asiento": asiento,
        "Referencia": ref,
        "Debit": debit,
        "Credit": credit,
        "Concepto": ln
    }


# ==========================================================
# 5. GPT ONLY FOR CLASSIFICATION (FAST)
# ==========================================================
def classify(rows):
    out = []

    for r in rows:
        ref = r["Referencia"]
        asiento = r["Asiento"]
        debit = r["Debit"]
        credit = r["Credit"]

        if ref == "":
            reason = "Payment"
        elif asiento == "VEN" and credit not in ("", "0,00", "0.00"):
            reason = "Credit Note"
        else:
            reason = "Invoice"

        r["Reason"] = reason
        out.append(r)

    return out


# ==========================================================
# 6. EXCEL EXPORT
# ==========================================================
def to_excel(df):
    b = BytesIO()
    df.to_excel(b, index=False)
    b.seek(0)
    return b


# ==========================================================
# UI
# ==========================================================
file = st.file_uploader("Upload Ledger PDF", type=["pdf"])

if file:

    raw = extract_raw(file)

    merged = merge_rows(raw)
    split = split_rows(merged)

    st.success(f"Detected {len(split)} ledger rows.")
    st.text_area("Preview", "\n".join(split[:20]), height=200)

    parsed = [parse_line(ln) for ln in split]
    final = classify(parsed)

    df = pd.DataFrame(final)

    st.dataframe(df, hide_index=True, use_container_width=True)

    st.download_button(
        "Download Excel",
        data=to_excel(df),
        file_name="datafalcon_final.xlsx"
    )
else:
    st.info("Upload a PDF to start.")
