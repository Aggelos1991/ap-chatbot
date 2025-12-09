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
st.set_page_config(page_title="🦅 DataFalcon Pro — FINAL STRICT VERSION", layout="wide")
st.title("🦅 DataFalcon Pro — FINAL STRICT REFERENCIA EXTRACTOR (NO GPT)")


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

            # If PDF text layer exists
            if text:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean and "saldo" not in clean.lower():
                        all_lines.append(clean)
            else:
                # OCR fallback
                try:
                    images = convert_from_bytes(pdf_bytes, dpi=250,
                                                first_page=idx, last_page=idx)
                    ocr_text = pytesseract.image_to_string(
                        images[0], lang="spa+eng+ell"
                    )
                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if clean and "saldo" not in clean.lower():
                            all_lines.append(clean)
                except:
                    pass

    return all_lines


# ==========================================================
# STRICT REFERENCIA EXTRACTION
# ==========================================================
def extract_referencia(line):
    """
    The ONLY acceptable reference is a 12–18 digit number.
    If it's missing → empty cell (payment row).
    """
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
# PARSE LEDGER ROW
# ==========================================================
def parse_ledger_line(line):
    parts = line.split()
    if len(parts) < 6:
        return None

    fecha = parts[0]

    # get referencia
    referencia = extract_referencia(line)

    # If no referencia → PAYMENT ROW
    if not referencia:
        # Credit amount is the second-to-last number
        amounts = re.findall(r"[-\d\.,]+", line)
        credit = normalize_amount(amounts[-2]) if len(amounts) >= 2 else ""
        concepto = " ".join(parts[4:])
        return {
            "Referencia": "",
            "Concepto": concepto,
            "Date": fecha,
            "Reason": "Payment",
            "Debit": "",
            "Credit": credit
        }

    # REFERENCIA EXISTS → INVOICE OR CREDIT NOTE
    # find index of referencia
    idx = parts.index(referencia)

    # concepto = between column 4 and referencia
    concepto = " ".join(parts[4:idx])

    # tail after referencia: [F.Valor, Debe, Haber, Saldo]
    tail = parts[idx + 1:]
    if len(tail) < 3:
        return None

    debe = normalize_amount(tail[-3])
    haber = normalize_amount(tail[-2])
    # saldo = tail[-1]  # NOT USED

    # classification
    if debe:
        reason = "Invoice"
    elif haber:
        reason = "Credit Note"
    else:
        return None

    return {
        "Referencia": referencia,
        "Concepto": concepto,
        "Date": fecha,
        "Reason": reason,
        "Debit": debe,
        "Credit": haber
    }


# ==========================================================
# PROCESS FULL STATEMENT
# ==========================================================
def extract_records(lines):
    records = []
    for line in lines:
        parsed = parse_ledger_line(line)
        if parsed:
            records.append(parsed)
    return records


# ==========================================================
# EXPORT TO EXCEL
# ==========================================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buff = BytesIO()
    df.to_excel(buff, index=False)
    buff.seek(0)
    return buff


# ==========================================================
# UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    lines = extract_raw_lines(uploaded_pdf)
    st.text_area("📄 Preview (first 30 lines)", "\n".join(lines[:30]), height=280)

    if st.button("🚀 Extract"):
        records = extract_records(lines)
        df = pd.DataFrame(records)

        st.success(f"Extracted {len(df)} valid rows.")
        st.dataframe(df, use_container_width=True, hide_index=True)

        if not df.empty:
            total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
            total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
            net = total_debit - total_credit

            c1, c2, c3 = st.columns(3)
            c1.metric("Total Debit", f"{total_debit:,.2f}")
            c2.metric("Total Credit", f"{total_credit:,.2f}")
            c3.metric("Net", f"{net:,.2f}")

            st.download_button(
                "⬇️ Download Excel",
                to_excel_bytes(records),
                "statement.xlsx"
            )

else:
    st.info("Upload a PDF to begin.")
