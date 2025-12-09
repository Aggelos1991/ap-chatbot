# ==========================================================
# 🦅 DataFalcon Pro v4 — GPT Hybrid (NO OCR, CORRECT FALLBACK)
# ==========================================================

import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from openai import OpenAI

# ==========================================================
# CONFIG
# ==========================================================

st.set_page_config(page_title="🦅 DataFalcon Pro v4 — GPT Hybrid", layout="wide")
st.title("🦅 DataFalcon Pro v4 — GPT Hybrid Edition (FINAL)")

api_key = st.secrets.get("OPENAI_API_KEY")
client = OpenAI(api_key=api_key)
GPT_MODEL = "gpt-4o-mini"


# ==========================================================
# CLEAN AMOUNTS
# ==========================================================

def clean_amount(v):
    if v is None or v == "":
        return None
    s = re.sub(r"[^\d,.\-]", "", str(v))

    if "," in s and "." in s and s.find(".") < s.find(","):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return None


# ==========================================================
# GPT EXTRACTOR
# ==========================================================

def gpt_extract_table(text):
    prompt = f"""
Extract a clean table in JSON array format.
Each object MUST contain exactly:
- "Referencia"
- "Debit"
- "Credit"

If a field is missing, return empty string.

Text to parse:
{text}
"""

    response = client.responses.create(
        model=GPT_MODEL,
        input=prompt,
        max_output_tokens=4000
    )

    try:
        return pd.read_json(BytesIO(response.output_text.encode()))
    except:
        return pd.DataFrame()


# ==========================================================
# HYBRID ENGINE
# ==========================================================

def extract_hybrid(file_bytes):

    # 1) Try pdfplumber
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        tables = []
        for page in pdf.pages:
            tbl = page.extract_table()
            if tbl:
                tables.append(tbl)

    # If no table: GPT
    if len(tables) == 0:
        raw_text = ""
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for p in pdf.pages:
                raw_text += p.extract_text() or ""
        return gpt_extract_table(raw_text), "gpt"

    # Convert first table
    table = tables[0]
    header = table[0]
    rows = table[1:]
    df = pd.DataFrame(rows, columns=header)

    # If table is garbage (1 column, 1 row, etc.) → GPT
    if df.shape[1] < 3 or df.shape[0] < 2:
        raw_text = ""
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for p in pdf.pages:
                raw_text += p.extract_text() or ""
        return gpt_extract_table(raw_text), "gpt"

    return df, "pdfplumber"


# ==========================================================
# CLASSIFICATION
# ==========================================================

def classify_entry(row):
    ref = str(row.get("Referencia", "")).strip()
    debit = clean_amount(row.get("Debit"))
    credit = clean_amount(row.get("Credit"))

    if ref == "" or ref.lower() == "none":
        return {"Document": "", "Reason": "Payment", "Amount": debit if debit else credit}

    if debit:
        return {"Document": ref, "Reason": "Invoice", "Amount": debit}

    if credit:
        return {"Document": ref, "Reason": "Credit Note", "Amount": credit}

    return {"Document": ref, "Reason": "Unknown", "Amount": None}


# ==========================================================
# UI
# ==========================================================

uploaded = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded:

    file_bytes = uploaded.read()

    df_raw, method = extract_hybrid(file_bytes)

    st.write(f"### Extraction Method: **{method.upper()}**")
    st.dataframe(df_raw, use_container_width=True)

    # Ensure required fields
    for col in ["Referencia", "Debit", "Credit"]:
        if col not in df_raw.columns:
            df_raw[col] = ""

    final = pd.DataFrame([classify_entry(r) for _, r in df_raw.iterrows()])

    st.write("### 🧠 Classified Results (FINAL)")
    st.dataframe(final, use_container_width=True)

    from io import BytesIO
    output = BytesIO()
    final.to_excel(output, index=False)

    st.download_button(
        "⬇️ Download Excel",
        data=output.getvalue(),
        file_name="DataFalcon_v4.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
