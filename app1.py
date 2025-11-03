# ==========================================================
# 🦅 DataFalcon Pro — Cloud Version (OCR API + GPT Extractor)
# ==========================================================
import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
import requests

# ==========================================================
# STREAMLIT CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Cloud OCR", layout="wide")
st.title("🦅 DataFalcon Pro — Hybrid GPT + Cloud OCR Extractor")

# ==========================================================
# OPENAI CONFIG
# ==========================================================
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
PRIMARY_MODEL = "gpt-4o-mini"
BACKUP_MODEL = "gpt-4o"

# ==========================================================
# OCR.SPACE API CONFIG  (multi-language)
# ==========================================================
OCR_API_KEY = "K83109465388957"
OCR_LANGUAGES = "spa+ell+eng"  # Spanish + Greek + English

def cloud_ocr_extract(pdf_bytes):
    """Perform OCR on scanned PDFs using the OCR.space API."""
    st.info("☁️ Uploading scanned PDF to OCR.space for text recognition...")
    try:
        response = requests.post(
            "https://api.ocr.space/parse/image",
            files={"file": ("document.pdf", pdf_bytes)},
            data={
                "apikey": OCR_API_KEY,
                "language": OCR_LANGUAGES,
                "isOverlayRequired": False
            },
            timeout=180
        )
        result = response.json()
        text = "\n".join(
            [r.get("ParsedText", "") for r in result.get("ParsedResults", []) if r.get("ParsedText")]
        )
        if not text.strip():
            st.error("⚠️ OCR API returned no text.")
            return []
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        st.success(f"✅ OCR completed — {len(lines)} lines extracted.")
        return lines
    except Exception as e:
        st.error(f"❌ OCR API error: {e}")
        return []

# ==========================================================
# HELPER FUNCTIONS
# ==========================================================
def normalize_number(value):
    """Normalize decimals like 1.234,56 → 1234.56"""
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
        return round(float(s), 2)
    except:
        return ""

# ==========================================================
# TEXT + CLOUD OCR EXTRACTION
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    """Extract text from PDF; use OCR API if no text layer found."""
    all_lines = []
    pdf_bytes = uploaded_pdf.getvalue()

    try:
        with pdfplumber.open(uploaded_pdf) as pdf:
            sample_text = any(page.extract_text() for page in pdf.pages[:3] if page.extract_text())

        if sample_text:
            st.info("📄 Detected searchable PDF → fast text extraction")
            with pdfplumber.open(uploaded_pdf) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text:
                        continue
                    for line in text.split("\n"):
                        clean = " ".join(line.split())
                        if clean:
                            all_lines.append(clean)
        else:
            st.warning("📸 No text layer found → using Cloud OCR (Spanish + Greek + English)")
            all_lines = cloud_ocr_extract(pdf_bytes)

    except Exception as e:
        st.error(f"❌ Error extracting text: {e}")

    return all_lines

# ==========================================================
# GPT EXTRACTOR
# ==========================================================
def parse_gpt_response(content, batch_num):
    match = re.search(r"\[.*\]", content, re.DOTALL)
    if not match:
        st.warning(f"Batch {batch_num}: No JSON found.\n{content[:200]}")
        return []
    try:
        return json.loads(match.group(0))
    except json.JSONDecodeError as e:
        st.warning(f"Batch {batch_num}: JSON decode error → {e}")
        return []

def extract_with_gpt(lines):
    """Classify entries as Invoice / Payment / Credit Note."""
    BATCH_SIZE = 60
    all_records = []
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i+BATCH_SIZE]
        text_block = "\n".join(batch)
        prompt = f"""
You are a financial data extractor specialized in Spanish vendor statements.
Each line may include: Fecha, Documento, Descripción, DEBE, HABER, SALDO.
Extract structured data and classify each entry as Invoice, Payment, or Credit Note.
Output strict JSON array only.
FORMAT:
[
  {{
    "Alternative Document": "...",
    "Date": "dd/mm/yy",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "DEBE amount or empty",
    "Credit": "HABER amount or empty"
  }}
]
Text:
{text_block}
"""
        data = []
        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0
                )
                content = response.choices[0].message.content.strip()
                if i == 0:
                    st.text_area(f"🧠 GPT Response (Batch 1 – {model})", content, height=250)
                data = parse_gpt_response(content, i//BATCH_SIZE + 1)
                if data:
                    break
            except Exception as e:
                st.warning(f"GPT error ({model}): {e}")

        for row in data:
            alt = str(row.get("Alternative Document", "")).strip()
            if not alt or re.search(r"(asiento|saldo|total|iva)", alt, re.I):
                continue
            debit = normalize_number(row.get("Debit", ""))
            credit = normalize_number(row.get("Credit", ""))
            reason = row.get("Reason", "").strip()

            if debit and not credit:
                reason = "Invoice"
            elif credit and not debit:
                if re.search(r"abono|nota|crédit|descuento", str(row), re.I):
                    reason = "Credit Note"
                else:
                    reason = "Payment"
            elif not debit and not credit:
                continue

            all_records.append({
                "Alternative Document": alt,
                "Date": row.get("Date", "").strip(),
                "Reason": reason,
                "Debit": debit,
                "Credit": credit
            })
    return all_records

# ==========================================================
# EXPORT
# ==========================================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text / running OCR..."):
        lines = extract_raw_lines(uploaded_pdf)

    st.text_area("📋 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if lines and st.button("🤖 Run GPT Extraction", type="primary"):
        with st.spinner("Analyzing with GPT models..."):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)
            st.success(f"✅ Extraction complete — {len(df)} records found!")
            st.dataframe(df, use_container_width=True, hide_index=True)

            total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
            total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
            net = round(total_debit - total_credit, 2)

            c1, c2, c3 = st.columns(3)
            c1.metric("💰 Total Debit", f"{total_debit:,.2f}")
            c2.metric("💳 Total Credit", f"{total_credit:,.2f}")
            c3.metric("⚖️ Net", f"{net:,.2f}")

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(data),
                file_name=f"vendor_statement_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No structured data detected. Check GPT response above.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
