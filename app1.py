import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro")
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
MODEL = "gpt-4o-mini"
# ==========================================================
# HELPERS
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
def extract_raw_lines(uploaded_pdf):
 """Extract ALL text lines from every page of the PDF."""
 all_lines = []
 with pdfplumber.open(uploaded_pdf) as pdf:
  for page in pdf.pages:
   text = page.extract_text()
   if not text:
    continue
   for line in text.split("\n"):
    clean_line = " ".join(line.split())
    if clean_line.strip():
     all_lines.append(clean_line)
 return all_lines
# ==========================================================
# GPT EXTRACTOR — FIXED CREDIT + NEGATIVE HANDLING
# ==========================================================
def extract_with_gpt(lines):
 """Use GPT to detect Debit (DEBE) and Credit (HABER) from vendor statements."""
 BATCH_SIZE = 100
 all_records = []
 
 for i in range(0, len(lines), BATCH_SIZE):
  batch = lines[i:i + BATCH_SIZE]
  text_block = "\n".join(batch)
 
  prompt = f"""Extract accounting transactions from this text.

**COLUMNS:**
- N° DOC → Document number (1729, 1775, etc.)
- DEBE → Invoice amounts (Debit)
- HABER/CREDIT → Payment amounts (Credit) 
- SALDO → Running balance (IGNORE for extraction)

**For each transaction:**
{{"Alternative Document": "N° DOC number", 
 "Date": "dd/mm/yy", 
 "Reason": "Invoice|Payment|Credit Note",
 "Debit": "DEBE amount", 
 "Credit": "HABER amount"}}

**RULES:**
1. DEBE > 0 = "Invoice" 
2. HABER/CREDIT > 0 AND contains payment keywords = "Payment"
3. DEBE < 0 OR reason indicates credit note = "Credit Note" (put ABSOLUTE value in Credit)
4. NEVER use SALDO values
5. Return ONLY JSON array: []

**PAYMENT KEYWORDS (for Reason="Payment"):** πληρωμή,payment,bank transfer,transferencia,transfer,trf,remesa,pago,deposit,έμβασμα,εξοφληση,pagado,paid

Text:
{text_block}"""
 
  try:
   response = client.chat.completions.create(
    model=MODEL,
    messages=[{"role": "user", "content": prompt}],
    temperature=0.0
   )
   content = response.choices[0].message.content.strip()
 
   # Debug
   if i == 0: # Only show first batch
    st.text_area("GPT Response (Batch 1):", content, height=200, key="debug_1")
 
   json_match = re.search(r'\[.*\]', content, re.DOTALL)
   if not json_match:
    json_match = re.search(r'(\[.*?\])', content, re.DOTALL)
 
   if json_match:
    json_str = json_match.group(0)
    data = json.loads(json_str)
 
    for row in data:
     alt_doc = str(row.get("Alternative Document", "")).strip()
 
     # Skip invalid documents
     if not alt_doc or re.search(r"concil|saldo|total|iva", alt_doc, re.IGNORECASE):
      continue
 
     debit_raw = row.get("Debit", "")
     credit_raw = row.get("Credit", "")
 
     debit_val = normalize_number(debit_raw)
     credit_val = normalize_number(credit_raw)
 
     reason = row.get("Reason", "Invoice").strip()
 
     # 🆕 FIXED: Handle negative DEBE as Credit Note
     if debit_val != "" and float(debit_val) < 0:
      credit_val = abs(float(debit_val))
      debit_val = ""
      reason = "Credit Note"
 
     # FIXED: ONLY classify as Payment if GPT already marked it as Payment (has payment keywords)
     # Don't override Credit Notes or Invoices
     if reason == "Payment" and credit_val != "" and float(credit_val) > 0:
      pass  # Keep as Payment
     elif reason == "Credit Note" or (debit_val != "" and float(debit_val) < 0):
      reason = "Credit Note"
      if credit_val == "":
       credit_val = abs(float(debit_val)) if debit_val != "" else ""
       debit_val = ""
     elif debit_val != "" and float(debit_val) > 0:
      reason = "Invoice"
 
     all_records.append({
      "Alternative Document": alt_doc,
      "Date": str(row.get("Date", "")).strip(),
      "Reason": reason,
      "Debit": debit_val,
      "Credit": credit_val
     })
   else:
    st.warning(f"No JSON found in batch {i//BATCH_SIZE + 1}")
 
  except Exception as e:
   st.warning(f"GPT error batch {i//BATCH_SIZE + 1}: {e}")
   continue
 
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
 with st.spinner("📄 Extracting text from all pages..."):
  lines = extract_raw_lines(uploaded_pdf)
 
 st.success(f"✅ Found {len(lines)} lines of text!")
 st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)
 
 if st.button("🤖 Run Hybrid Extraction", type="primary"):
  with st.spinner("Analyzing with GPT-4o-mini..."):
   data = extract_with_gpt(lines)
 
  if data:
   df = pd.DataFrame(data)
   st.success(f"✅ Extraction complete — {len(df)} valid records found!")
   st.dataframe(df, use_container_width=True, hide_index=True)
 
   # Totals
   try:
    total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
    total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
    net = round(total_debit - total_credit, 2)
 
    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
    col2.metric("
