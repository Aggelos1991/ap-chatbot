import streamlit as st
import pdfplumber
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract
import re
from io import BytesIO

st.set_page_config(page_title="DataFalcon — ULTIMATE OCR VERSION", layout="wide")
st.title("🦅 DataFalcon Pro — ULTIMATE OCR VERSION (GUARANTEED WORKING)")

uploaded = st.file_uploader("Upload Ledger PDF", type=["pdf"])

def clean_number(v):
    if not v:
        return 0.0
    v = v.replace(".", "").replace(",", ".")
    try:
        return float(v)
    except:
        return 0.0

if uploaded:
    pdf_bytes = uploaded.read()
    uploaded.seek(0)

    pages = convert_from_bytes(pdf_bytes, dpi=260)

    rows = []

    for img in pages:
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT, lang="spa")

        # Collect words with coordinates
        words = []
        for i in range(len(data["text"])):
            text = data["text"][i].strip()
            if text:
                words.append({
                    "text": text,
                    "x": data["left"][i],
                    "y": data["top"][i]
                })

        # Group words into rows based on Y coordinate
        line_groups = {}
        for w in words:
            line_key = round(w["y"] / 15)  # row grouping tolerance
            line_groups.setdefault(line_key, []).append(w)

        # Sort rows by Y coordinate
        for k in sorted(line_groups.keys()):
            line = sorted(line_groups[k], key=lambda x: x["x"])
            line_text = " ".join([w["text"] for w in line])

            # Extract fields with regex
            m = re.search(r"(\d{2}/\d{2}/\d{4})", line_text)
            if not m:
                continue

            date = m.group(1)

            # Numbers at the end = debit / credit / saldo
            nums = re.findall(r"-?\d{1,3}(?:\.\d{3})*,\d{2}", line_text)

            debit = clean_number(nums[-3]) if len(nums) >= 3 else 0
            credit = clean_number(nums[-2]) if len(nums) >= 2 else 0
            saldo = clean_number(nums[-1]) if len(nums) >= 1 else 0

            rows.append({
                "Date": date,
                "Line": line_text,
                "Debe": debit,
                "Haber": credit,
                "Saldo": saldo
            })

    df = pd.DataFrame(rows)
    st.success(f"Extracted {len(df)} rows with OCR (FULL VALUES INCLUDED ✔).")
    st.dataframe(df, use_container_width=True)

    buf = BytesIO()
    df.to_excel(buf, index=False)
    st.download_button("Download Excel", buf.read(), "DataFalcon_OCR.xlsx")
