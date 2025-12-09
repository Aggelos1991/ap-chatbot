# ==========================================================
# 🦅 DataFalcon Pro v2 — ULTRA FAST EDITION (FINAL)
# ==========================================================

import streamlit as st
import pandas as pd
import pdfplumber
from pdf2image import convert_from_bytes
import pytesseract
from io import BytesIO
import re

# ==========================================================
# PAGE CONFIG
# ==========================================================

st.set_page_config(page_title="🦅 DataFalcon Pro v2 — Ultra Fast", layout="wide")
st.title("🦅 DataFalcon Pro v2 — Ultra Fast Edition (FINAL)")


# ==========================================================
# HELPERS
# ==========================================================

def clean_amount(v):
    """Normalize amounts: remove symbols, convert EU format."""
    if v is None or v == "":
        return None
    s = str(v).strip()

    # Remove everything except digits, dot, comma, minus
    s = re.sub(r"[^\d,.\-]", "", s)

    # Convert EU format 1.234,56 → 1234.56
    if "," in s and "." in s:
        if s.find(".") < s.find(","):
            s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return None


def extract_table_from_pdf(file_bytes):
    """
    Try structured extraction via pdfplumber.
    If no tables, fallback to OCR.
    """
    records = []

    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    header = table[0]
                    for row in table[1:]:
                        record = dict(zip(header, row))
                        records.append(record)

        if len(records) > 0:
            return pd.DataFrame(records)

    except:
        pass

    # =========== FALLBACK OCR =============
    images = convert_from_bytes(file_bytes)
    text = "\n".join([pytesseract.image_to_string(img) for img in images])

    rows = []
    for line in text.split("\n"):
        cols = re.split(r"\s{2,}", line.strip())
        if len(cols) >= 4:
            rows.append(cols[:4])

    if len(rows) == 0:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    df.columns = [f"col_{i}" for i in range(df.shape[1])]
    return df


# ==========================================================
# CLASSIFICATION LOGIC (FINAL)
# ==========================================================

def classify_entry(row):
    referencia = str(row.get("Referencia", "")).strip()
    debit = clean_amount(row.get("Debit"))
    credit = clean_amount(row.get("Credit"))

    if referencia in ["", None, "None"]:
        amount = debit if debit is not None else credit
        return {
            "Document": "",
            "Reason": "Payment",
            "Amount": amount
        }

    if debit is not None:
        return {
            "Document": referencia,
            "Reason": "Invoice",
            "Amount": debit
        }

    if credit is not None:
        return {
            "Document": referencia,
            "Reason": "Credit Note",
            "Amount": credit
        }

    return {
        "Document": referencia,
        "Reason": "Unknown",
        "Amount": None
    }


# ==========================================================
# UI
# ==========================================================

uploaded = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded:
    st.success("📄 PDF uploaded — extracting...")

    file_bytes = uploaded.read()

    df_raw = extract_table_from_pdf(file_bytes)

    if df_raw.empty:
        st.error("❌ No data extracted from PDF.")
        st.stop()

    st.write("### Extracted Raw Table")
    st.dataframe(df_raw, use_container_width=True)

    # ---- CLEAN HEADERS ----
    df = df_raw.rename(columns=lambda x: x.strip().replace(" ", "_"))

    # Ensure required columns exist
    required = ["Referencia", "Debit", "Credit"]
    for col in required:
        if col not in df.columns:
            df[col] = ""

    # ---- CLASSIFICATION ----
    results = []
    for _, row in df.iterrows():
        results.append(classify_entry(row))

    final_df = pd.DataFrame(results)

    st.write("### 🧠 Classified Results (FINAL)")
    st.dataframe(final_df, use_container_width=True)

    # ---- DOWNLOAD ----
    output = BytesIO()
    final_df.to_excel(output, index=False)
    st.download_button(
        label="⬇️ Download Excel",
        data=output.getvalue(),
        file_name="DataFalcon_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
