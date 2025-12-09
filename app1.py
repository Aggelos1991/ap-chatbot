# ==========================================================
# 🦅 DataFalcon Pro v3 — Ultra Fast (NO OCR, FINAL RULES)
# ==========================================================

import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
import re

# ==========================================================
# PAGE CONFIG
# ==========================================================

st.set_page_config(page_title="🦅 DataFalcon Pro v3 — Ultra Fast", layout="wide")
st.title("🦅 DataFalcon Pro v3 — Ultra Fast Edition (FINAL)")


# ==========================================================
# HELPERS
# ==========================================================

def clean_amount(v):
    """Normalize numbers — EU format support."""
    if v is None or v == "":
        return None
    s = str(v).strip()

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
    """Extract table using only pdfplumber (NO OCR)."""
    records = []

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                header = table[0]
                for row in table[1:]:
                    records.append(dict(zip(header, row)))

    if len(records) == 0:
        return pd.DataFrame()

    return pd.DataFrame(records)


# ==========================================================
# FINAL CLASSIFICATION RULES
# ==========================================================

def classify_entry(row):
    referencia = str(row.get("Referencia", "")).strip()

    debit = clean_amount(row.get("Debit"))
    credit = clean_amount(row.get("Credit"))

    # RULE 1 → Payment (NO referencia)
    if referencia == "" or referencia.lower() == "none":
        amount = debit if debit is not None else credit
        return {
            "Document": "",
            "Reason": "Payment",
            "Amount": amount
        }

    # RULE 2 → Invoice (Referencia + Debit)
    if debit is not None:
        return {
            "Document": referencia,
            "Reason": "Invoice",
            "Amount": debit
        }

    # RULE 3 → Credit Note (Referencia + Credit)
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
    st.success("📄 PDF uploaded — extracting table...")

    file_bytes = uploaded.read()

    df_raw = extract_table_from_pdf(file_bytes)

    if df_raw.empty:
        st.error("❌ No table found in PDF (and OCR is disabled).")
        st.stop()

    st.write("### Extracted Raw Table")
    st.dataframe(df_raw, use_container_width=True)

    # Clean headers
    df = df_raw.rename(columns=lambda x: x.strip().replace(" ", "_"))

    # Ensure required fields
    for col in ["Referencia", "Debit", "Credit"]:
        if col not in df.columns:
            df[col] = ""

    # --- Classification ---
    results = [classify_entry(row) for _, row in df.iterrows()]
    final_df = pd.DataFrame(results)

    st.write("### 🧠 Classified Results (FINAL)")
    st.dataframe(final_df, use_container_width=True)

    # --- Download ---
    output = BytesIO()
    final_df.to_excel(output, index=False)

    st.download_button(
        label="⬇️ Download Excel",
        data=output.getvalue(),
        file_name="DataFalcon_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
