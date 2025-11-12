import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from pdf2image import convert_from_bytes
import pytesseract

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — OCR Accounting Edition (Final)")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No OpenAI API key found.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ==========================================================
# HELPERS
# ==========================================================
def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "":
        return ""
    s = str(v).strip().replace("€", "").replace(" ", "")
    # normalize 1.234,56 → 1234.56
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
# PDF + OCR EXTRACTION
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    all_lines = []
    pdf_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)
    ocr_pages = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text or len(text.strip()) < 10:
                ocr_pages.append(i)
                try:
                    images = convert_from_bytes(pdf_bytes, dpi=250, first_page=i, last_page=i)
                    text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")
                except Exception as e:
                    st.warning(f"OCR failed on page {i}: {e}")
                    continue

            for line in text.split("\n"):
                clean = " ".join(line.split())
                if not clean:
                    continue
                # Skip balance / summary lines
                if re.search(r"\b(saldo|total\s*saldo|anterior|final)\b", clean, re.IGNORECASE):
                    continue
                all_lines.append(clean)

    if ocr_pages:
        st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")
    return all_lines

# ==========================================================
# GPT STRUCTURAL PARSER (JSON only)
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT only to structure lines into columns, not interpret accounting meaning."""
    all_records = []
    BATCH_SIZE = 50

    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i+BATCH_SIZE]
        text_block = "\n".join(batch)

        prompt = f"""
You are an expert financial data parser.  
From the following ledger lines (Spanish/Greek), identify **only**:
- Alternative Document
- Date
- Description
- Debit (DEBE)
- Credit (HABER)

Rules:
- Output STRICT JSON array, no commentary.
- Exclude any running balances, totals, or saldo lines.
- Keep numeric amounts exactly as shown.

Example:
[
  {{
    "Alternative Document": "A250212",
    "Date": "25/02/25",
    "Description": "Cobro factura A250212 Rec",
    "Debit": "",
    "Credit": "1793.89"
  }}
]

Text:
{text_block}
"""
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            content = resp.choices[0].message.content.strip()
            match = re.search(r"\[.*\]", content, re.DOTALL)
            if not match:
                st.warning(f"Batch {i//BATCH_SIZE+1}: No JSON detected.")
                continue
            data = json.loads(match.group(0))
            all_records.extend(data)
        except Exception as e:
            st.warning(f"GPT batch {i//BATCH_SIZE+1} failed: {e}")

    return all_records

# ==========================================================
# LOGIC: classify by DEBE/HABER polarity
# ==========================================================
def classify_records(records):
    parsed = []
    for r in records:
        doc = str(r.get("Alternative Document", "")).strip()
        date = str(r.get("Date", "")).strip()
        desc = str(r.get("Description", "")).strip()

        debe = normalize_number(r.get("Debit", ""))
        haber = normalize_number(r.get("Credit", ""))
        reason, amount = "", 0.0

        # --- classify strictly by numeric direction ---
        if debe != "" and haber == "":
            reason = "Invoice" if debe > 0 else "Credit Note"
            amount = debe
        elif haber != "" and debe == "":
            reason = "Payment" if haber > 0 else "Reversal"
            amount = haber
        else:
            continue  # ignore malformed rows

        parsed.append({
            "Alternative Document": doc,
            "Date": date,
            "Description": desc,
            "Reason": reason,
            "Amount": amount
        })

    return pd.DataFrame(parsed)

# ==========================================================
# STREAMLIT APP
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("🔍 Extracting text (with OCR fallback)..."):
        lines = extract_raw_lines(uploaded_pdf)

    st.success(f"✅ Extracted {len(lines)} text lines.")
    st.text_area("📄 Preview of raw text (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Hybrid Extraction", type="primary"):
        with st.spinner("Analyzing with GPT..."):
            gpt_data = extract_with_gpt(lines)

        if not gpt_data:
            st.error("No structured records found.")
            st.stop()

        df = classify_records(gpt_data)
        st.success(f"✅ Extraction complete — {len(df)} valid records found!")

        st.dataframe(df, use_container_width=True, hide_index=True)

        totals = df.groupby("Reason")["Amount"].sum().round(2).reset_index()
        st.markdown("### 💰 Summary by Type")
        st.dataframe(totals, hide_index=True)

        # Excel export
        buf = BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)

        st.download_button(
            "⬇️ Download Excel",
            data=buf,
            file_name=f"DataFalcon_Extract_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload a vendor statement PDF to begin.")
