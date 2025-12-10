import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="DataFalcon Pro — FINAL VERSION", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL VERSION")

uploaded_pdf = st.file_uploader("📂 Upload Ledger PDF", type=["pdf"])

# ===========================================================
# REGEX PATTERN DESIGNED FOR SABA LEDGER FORMAT
# ===========================================================
pattern = re.compile(
    r"(?P<Date>\d{2}/\d{2}/\d{4})\s+"
    r"(?P<Asiento>VEN|GRL)\s*/\s*(?P<Code>\d+)\s+\d+\s*/?(?P<Doc>\S+)?\s*"
    r"(V|IR|FP)?\s*(?P<Reference>[A-Z0-9\-]+)?\s*"
    r"(?P<ValueDate>\d{2}/\d{2}/\d{4})?\s*"
    r"(?P<Debit>-?\d{1,3}(?:\.\d{3})*,\d{2})?\s*"
    r"(?P<Credit>-?\d{1,3}(?:\.\d{3})*,\d{2})?"
)

# ===========================================================
# NORMALIZE EUROPEAN NUMBERS
# ===========================================================
def to_float(x):
    if not x:
        return 0.0
    return float(x.replace(".", "").replace(",", "."))

# ===========================================================
# PARSE TEXT INTO CLEAN STRUCTURED ROWS
# ===========================================================
def parse_ledger(text):
    rows = []

    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue

        match = pattern.search(line)
        if match:
            d = match.groupdict()

            d["Debit"] = to_float(d["Debit"])
            d["Credit"] = to_float(d["Credit"])

            rows.append(d)

    return pd.DataFrame(rows)

# ===========================================================
# PROCESS PDF
# ===========================================================
if uploaded_pdf:
    text = ""

    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            block = page.extract_text()
            if block:
                text += block + "\n"

    df = parse_ledger(text)

    st.success(f"✅ Extracted {len(df)} clean ledger rows.")
    st.dataframe(df, use_container_width=True)

    # Excel download
    output = BytesIO()
    df.to_excel(output, index=False)

    st.download_button(
        "⬇️ Download Excel",
        data=output.getvalue(),
        file_name="DataFalcon_FINAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
