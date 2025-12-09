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
st.set_page_config(page_title="🦅 DataFalcon Pro — NO GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Ultra Fast Extractor (NO GPT)")

# ==========================================================
# PDF LINE EXTRACTION (TEXT + OCR FALLBACK)
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()

            # If text is readable
            if text:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean and "saldo" not in clean.lower():
                        all_lines.append(clean)
            else:
                # OCR fallback
                try:
                    images = convert_from_bytes(pdf_bytes, dpi=260, first_page=idx, last_page=idx)
                    ocr_text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")
                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if clean and "saldo" not in clean.lower():
                            all_lines.append(clean)
                except Exception as e:
                    st.warning(f"OCR skipped on page {idx}: {e}")

    return all_lines

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
# PARSER FOR LEDGER LINES
# ==========================================================
def parse_statement_line(line):
    """
    Expected structure (always stable):
    Fecha | Asiento | Documento | Libro | Descripción | Referencia | F. Valor | Debe | Haber
    """

    parts = line.split()

    # Not enough columns — skip
    if len(parts) < 9:
        return None

    # Extract fields by fixed positions
    fecha = parts[0]
    asiento = parts[1]
    documento = parts[2]
    libro = parts[3]

    # Descripción = variable length text from part 4 up to before Referencia
    # Referencia is always a long numeric code (12–18 digits)
    descripcion = []
    referencia = ""
    debe = ""
    haber = ""

    # Find referencia index
    for i, p in enumerate(parts):
        if re.fullmatch(r"\d{12,18}", p):
            referencia = p
            desc_end = i
            break

    if not referencia:
        return None  # invalid line

    descripcion = " ".join(parts[4:desc_end])

    # The last two numeric fields are Debe and Haber
    maybe_amounts = parts[desc_end + 2 :]

    # F. Valor -> skip 1 position
    # Then Debe / Haber
    if len(maybe_amounts) >= 2:
        debe = normalize_amount(maybe_amounts[-2])
        haber = normalize_amount(maybe_amounts[-1])

    # CLASSIFICATION
    if debe:
        reason = "Invoice"
    elif haber:
        # Payment or credit note?
        if re.search(r"pago|cobro|transfer", descripcion.lower()):
            reason = "Payment"
        else:
            reason = "Credit Note"
    else:
        return None

    return {
        "Referencia": referencia,
        "Concepto": descripcion,
        "Date": fecha,
        "Reason": reason,
        "Debit": debe,
        "Credit": haber
    }

# ==========================================================
# FULL EXTRACTOR
# ==========================================================
def extract_records(lines):
    records = []
    for line in lines:
        parsed = parse_statement_line(line)
        if parsed:
            records.append(parsed)
    return records

# ==========================================================
# EXPORT
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
    st.text_area("Preview (first 30 lines)", "\n".join(lines[:30]), height=280)

    if st.button("🚀 Extract"):
        records = extract_records(lines)
        df = pd.DataFrame(records)

        st.success(f"Extracted {len(df)} valid records.")
        st.dataframe(df, hide_index=True, use_container_width=True)

        # Totals
        if not df.empty:
            total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
            total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()

            c1, c2, c3 = st.columns(3)
            c1.metric("Total Debit", f"{total_debit:,.2f}")
            c2.metric("Total Credit", f"{total_credit:,.2f}")
            c3.metric("Net", f"{total_debit - total_credit:,.2f}")

            st.download_button(
                "⬇️ Download Excel",
                to_excel_bytes(records),
                "statement.xlsx"
            )

else:
    st.info("Upload a PDF to begin.")
