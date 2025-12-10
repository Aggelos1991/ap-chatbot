import os, re, json, time, concurrent.futures
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Final Version", layout="wide")
st.title("🦅 DataFalcon Pro — Final Version")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ Missing OpenAI API Key.")
    st.stop()

client = OpenAI(api_key=api_key)
PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"


# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(v):
    if not v:
        return ""
    s = str(v).replace(" ", "").strip()
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return round(float(s), 2)
    except:
        return ""


# ==========================================================
# PDF EXTRACTION WITH OCR FALLBACK
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)
    ocr_pages = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    clean = " ".join(line.split())
                    if clean and not re.search(r"saldo", clean, re.IGNORECASE):
                        all_lines.append(clean)
            else:
                ocr_pages.append(i)
                try:
                    img = convert_from_bytes(pdf_bytes, dpi=240, first_page=i, last_page=i)[0]
                    ocr_text = pytesseract.image_to_string(img, lang="spa+eng+ell")
                    for line in ocr_text.split("\n"):
                        clean = " ".join(line.split())
                        if clean and not re.search(r"saldo", clean, re.IGNORECASE):
                            all_lines.append(clean)
                except:
                    pass

    if ocr_pages:
        st.info(f"OCR applied on pages: {ocr_pages}")

    return all_lines


# ==========================================================
# JSON PARSER
# ==========================================================
def parse_gpt_response(content):
    m = re.search(r'\[.*\]', content, re.DOTALL)
    if not m:
        return []
    try:
        return json.loads(m.group(0))
    except:
        return []


# ==========================================================
# GPT TIMEOUT WRAPPER
# ==========================================================
def gpt_call_with_timeout(model, prompt, timeout=12):
    def task():
        return client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}]
        )

    with concurrent.futures.ThreadPoolExecutor() as executor:
        future = executor.submit(task)
        try:
            return future.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            return None


# ==========================================================
# MAIN GPT EXTRACTOR (SMALL BATCHES + TIMEOUT)
# ==========================================================
def extract_with_gpt(lines):
    all_records = []
    BATCH = 20  # small batches => fast, stable

    for i in range(0, len(lines), BATCH):
        block = "\n".join(lines[i:i+BATCH])

        prompt = f"""
Extract ledger data.

ABSOLUTE RULES:
- Document number = ONLY the field 'Referencia'.
- If Referencia empty → Payment.
- If Asiento = VEN AND Credit > 0 → Credit Note.
- Everything else → Invoice.
- NEVER extract numbers from Concepto.

FORMAT:
Fecha | Asiento | Documento | Libro | Descripción | Referencia | F. valor | Debe | Haber

Return ONLY JSON:
[
  {{
    "Fecha": "",
    "Asiento": "",
    "Referencia": "",
    "Concepto": "",
    "Debit": "",
    "Credit": ""
  }}
]

Text:
{block}
"""

        # Try primary model
        response = gpt_call_with_timeout(PRIMARY_MODEL, prompt, timeout=12)

        # Fallback
        if response is None:
            response = gpt_call_with_timeout(BACKUP_MODEL, prompt, timeout=12)

        # If still none → skip batch
        if response is None:
            st.warning(f"⚠️ GPT timeout on batch {i//BATCH+1}, skipping.")
            continue

        content = response.choices[0].message.content
        data = parse_gpt_response(content)
        if not data:
            continue

        # Final classification rules
        for r in data:
            ref = str(r.get("Referencia", "")).strip()
            asiento = str(r.get("Asiento", "")).strip().upper()

            debit = normalize_number(r.get("Debit", ""))
            credit = normalize_number(r.get("Credit", ""))

            if ref == "":
                reason = "Payment"
            elif asiento == "VEN" and credit not in ("", 0) and float(credit) > 0:
                reason = "Credit Note"
            else:
                reason = "Invoice"

            all_records.append({
                "Document": ref,
                "Date": r.get("Fecha", ""),
                "Asiento": asiento,
                "Concepto": r.get("Concepto", ""),
                "Reason": reason,
                "Debit": debit,
                "Credit": credit
            })

    return all_records


# ==========================================================
# EXCEL EXPORT
# ==========================================================
def to_excel(df):
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ==========================================================
# STREAMLIT UI
# ==========================================================
uploaded = st.file_uploader("📂 Upload Vendor Ledger PDF", type=["pdf"])

if uploaded:
    lines = extract_raw_lines(uploaded)
    st.success(f"Extracted {len(lines)} lines.")
    st.text_area("Preview", "\n".join(lines[:40]), height=250)

    if st.button("🚀 Run DataFalcon Extraction", type="primary"):
        with st.spinner("Processing with GPT…"):
            records = extract_with_gpt(lines)

        if records:
            df = pd.DataFrame(records)
            st.success(f"Extraction complete! {len(df)} rows.")
            st.dataframe(df, hide_index=True, use_container_width=True)

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel(df),
                file_name="datafalcon_output.xlsx"
            )
        else:
            st.error("No records extracted.")
else:
    st.info("Upload a PDF to begin.")
