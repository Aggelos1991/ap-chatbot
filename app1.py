import os, re, json
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
st.set_page_config(page_title="🦅 DataFalcon Pro — OCR + Regex Hybrid", layout="wide")
st.title("🦅 DataFalcon Pro — Final OCR + Regex Accounting Edition")
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
    if pd.isna(v) or v is None or (isinstance(v, str) and str(v).strip() == ""):
        return ""
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    s = str(v).replace("€", "").replace(" ", "").replace("$", "").replace("£", "")
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
    except ValueError:
        return ""
# ==========================================================
# PDF + OCR
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
                    images = convert_from_bytes(pdf_bytes, dpi=300, first_page=i, last_page=i)
                    text = pytesseract.image_to_string(images[0], lang="spa+eng+ell")
                except Exception as e:
                    st.warning(f"OCR failed on page {i}: {e}")
                    continue
            for line in text.split("\n"):
                clean = " ".join(line.split())
                if not clean:
                    continue
                if re.search(r"\b(saldo|balance|total\s*saldo|anterior|final|total)\b", clean, re.IGNORECASE):
                    continue
                all_lines.append(clean)
    if ocr_pages:
        st.info(f"OCR applied on pages: {', '.join(map(str, ocr_pages))}")
    return all_lines
# ==========================================================
# GPT → Structure (only doc/date/desc/full_line)
# ==========================================================
def gpt_structure(lines):
    all_records = []
    BATCH = 40  # Smaller batch to avoid token limits
    for i in range(0, len(lines), BATCH):
        batch = lines[i:i+BATCH]
        text_block = "\n".join(batch)
        prompt = f"""
You are an expert accounting statement parser. From the text, identify only transaction lines (not headers, footers, summaries).
For each transaction line, extract:
- Full Line: the exact full line text
- Alternative Document: invoice or payment code
- Date: transaction date
- Description: concept or description
Ignore any balance or saldo columns. Do not extract or include balances, totals, or running saldo.
Output only a valid JSON array of objects, nothing else.
Example:
[
  {{
    "Full Line": "25/02/25 A250212 Cobro factura A250212 Rec 100,50",
    "Alternative Document": "A250212",
    "Date": "25/02/25",
    "Description": "Cobro factura A250212 Rec"
  }}
]
Text:
{text_block}
"""
        try:
            r = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1
            )
            content = r.choices[0].message.content.strip()
            # More robust JSON extraction
            if content.startswith('[') and content.endswith(']'):
                data = json.loads(content)
            else:
                match = re.search(r'\[[\s\S]*\]', content)
                if match:
                    data = json.loads(match.group(0))
                else:
                    raise ValueError("No JSON array found")
            all_records.extend(data)
        except Exception as e:
            st.warning(f"GPT batch {i//BATCH+1} failed: {e}")
    return all_records
# ==========================================================
# REGEX → Extract DEBE/HABER
# ==========================================================
def extract_numbers(line):
    # Find numeric values, ignoring dates and codes
    # Assume last two numbers are Debe and Haber, or last one if only one
    # Remove date-like and code-like before finding numbers
    line_no_date = re.sub(r'\b\d{1,2}/\d{1,2}/\d{2,4}\b', '', line)
    line_clean = re.sub(r'\b[A-Z]\d+\b', '', line_no_date)  # Remove codes like A250212
    nums = re.findall(r'[-]?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?', line_clean)
    nums = [n for n in nums if len(n) > 1]  # Avoid single digits
    if len(nums) == 0:
        return "", ""
    elif len(nums) == 1:
        return normalize_number(nums[0]), ""
    else:
        # Last two
        return normalize_number(nums[-2]), normalize_number(nums[-1])
# ==========================================================
# CLASSIFY BY POLARITY
# ==========================================================
def classify(records):
    parsed = []
    for r in records:
        full_line = str(r.get("Full Line", "")).strip()
        doc = str(r.get("Alternative Document", "")).strip()
        date = str(r.get("Date", "")).strip()
        desc = str(r.get("Description", "")).strip()
        debe, haber = extract_numbers(full_line)
        if debe == "" and haber == "":
            continue
        reason = ""
        amount = 0.0
        if debe != "" and haber == "":
            sign = 1 if debe > 0 else -1
            reason = "Invoice" if sign > 0 else "Credit Note"
            amount = abs(debe)
        elif haber != "" and debe == "":
            sign = 1 if haber > 0 else -1
            reason = "Payment" if sign > 0 else "Reversal"
            amount = abs(haber)
        elif debe != "" and haber != "":
            # Assume debe is charge, haber is credit, but classify based on which is non-zero or logic
            # But per original: prefer HABER as payment
            reason = "Payment"
            amount = abs(haber) if haber != 0 else abs(debe)
        parsed.append({
            "Alternative Document": doc,
            "Date": date,
            "Description": desc,
            "Debe": debe,
            "Haber": haber,
            "Reason": reason,
            "Amount": amount
        })
    return pd.DataFrame(parsed)
# ==========================================================
# STREAMLIT APP
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])
if uploaded_pdf:
    with st.spinner("🧩 Extracting text …"):
        lines = extract_raw_lines(uploaded_pdf)
    st.success(f"✅ {len(lines)} lines extracted.")
    st.text_area("Preview of text (30 lines):", "\n".join(lines[:30]), height=250)
    if st.button("🤖 Run Hybrid Extraction", type="primary"):
        with st.spinner("Parsing structure with GPT and regex …"):
            base = gpt_structure(lines)
            df = classify(base)
        if len(df) == 0:
            st.warning("⚠️ No records found.")
            st.stop()
        st.success(f"✅ Extraction complete — {len(df)} records found.")
        st.dataframe(df, use_container_width=True, hide_index=True)
        totals = df.groupby("Reason")["Amount"].sum().round(2).reset_index()
        st.markdown("### 💰 Summary by Type")
        st.dataframe(totals, hide_index=True)
        buf = BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)
        st.download_button(
            "⬇️ Download Excel",
            data=buf,
            file_name=f"DataFalcon_Hybrid_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload a vendor statement PDF to begin.")
