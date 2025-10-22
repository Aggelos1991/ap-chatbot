import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Clean Numbers", layout="wide")
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
# 🔥 GIBBERISH REMOVER - KEEP ONLY FINAL NUMBER
# ==========================================================
def clean_invoice_number(alt_doc):
    """Remove ALL gibberish - keep ONLY final number"""
    if not alt_doc:
        return ""
    
    s = str(alt_doc).strip()
    
    # Pattern 1: ANYTHING + dashes + FINAL NUMBER (7+ digits)
    match = re.search(r'[-–—/]\s*(\d{7,})$', s)
    if match:
        return match.group(1)
    
    # Pattern 2: Gibberish + ANY 6+ digits at END
    match = re.search(r'.*?(\d{6,})$', s)
    if match:
        return match.group(1)
    
    # Pattern 3: Pure number extraction
    match = re.search(r'(\d{6,})', s)
    if match:
        return match.group(1)
    
    return ""

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
    """Extract all text lines from every page of the PDF."""
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split("\n"):
                if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", line):
                    all_lines.append(" ".join(line.split()))
    return all_lines

# ==========================================================
# GPT EXTRACTOR — WITH GIBBERISH CLEANER
# ==========================================================
def extract_with_gpt(lines):
    """Use GPT + CLEAN numbers from gibberish."""
    BATCH_SIZE = 150
    all_records = []
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        prompt = f"""
You are an expert accountant fluent in Spanish and Greek.
You are reading extracted lines from a vendor statement.
Each line may include columns labeled as:
- DEBE → Debit (Invoice)
- HABER → Credit (Payment)
- SALDO → Running Balance
- CONCEPTO → Description such as "Fra. emitida", "Cobro Efecto", etc.

Your task:
For each valid transaction line, output:
- "Alternative Document": document number (under Nº, Num, Documento, Factura, ΕνδIοNκVο, etc.)
- "Date": date if visible (dd/mm/yy or dd/mm/yyyy)
- "Reason": classify as "Invoice", "Payment", or "Credit Note"
- "Debit": numeric value under DEBE column (if exists)
- "Credit": numeric value under HABER column (if exists)

Rules:
1. If DEBE > 0 → Reason = "Invoice"
2. If HABER > 0 → Reason = "Payment"
3. If the line includes "Abono", "Nota de Credito", "NC", "πιστω", "Ακυρωτικό" → Reason = "Credit Note"
4. Ignore summary lines: "Saldo", "Apertura", "Total General", "IVA", "Base", "Impuestos".
5. Exclude any line where document contains "concil"
6. Ensure output is valid JSON array.

Lines:
\"\"\"{text_block}\"\"\"
"""
        try:
            response = client.chat.completions.create(  # ✅ FIXED API
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=4000
            )
            content = response.choices[0].message.content.strip()
            json_match = re.search(r"\[.*\]", content, re.DOTALL)
            if not json_match:
                continue
            data = json.loads(json_match.group(0))
        except Exception as e:
            st.warning(f"⚠️ GPT failed on batch {i//BATCH_SIZE + 1}: {e}")
            continue
        
        for row in data:
            alt_doc_raw = str(row.get("Alternative Document", "")).strip()
            
            # 🔥 GIBBERISH REMOVER - MAGIC HAPPENS HERE
            alt_doc_clean = clean_invoice_number(alt_doc_raw)
            
            # Skip if no clean number
            if not alt_doc_clean:
                continue
            
            # Exclude concil
            if re.search(r"concil", alt_doc_raw, re.IGNORECASE):
                continue
            
            debit_val = normalize_number(row.get("Debit"))
            credit_val = normalize_number(row.get("Credit"))
            
            # Move Cobro/Efecto to Credit
            concept = alt_doc_raw.lower()
            if "cobro" in concept or "efecto" in concept:
                credit_val = credit_val or debit_val
                debit_val = ""
            
            all_records.append({
                "Alternative Document": alt_doc_clean,  # ✅ CLEAN NUMBER ONLY
                "Raw Document": alt_doc_raw[:30] + "..." if len(alt_doc_raw) > 30 else alt_doc_raw,  # Debug
                "Date": str(row.get("Date", "")).strip(),
                "Reason": row.get("Reason", "").strip(),
                "Debit": debit_val,
                "Credit": credit_val
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
    with st.spinner("📄 Extracting text from all pages..."):
        lines = extract_raw_lines(uploaded_pdf)
    if not lines:
        st.warning("⚠️ No readable text lines found. Check if the PDF is scanned.")
    else:
        st.text_area("📄 Preview (first 25 lines):", "\n".join(lines[:25]), height=250)
        if st.button("🤖 Run Hybrid Extraction"):
            with st.spinner("🔍 Cleaning gibberish → Pure numbers..."):
                data = extract_with_gpt(lines)
            if not data:
                st.warning("⚠️ No structured data detected.")
            else:
                df = pd.DataFrame(data)
                st.success(f"✅ Extraction complete — {len(df)} valid records found.")
                
                # Show Raw vs Clean
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("✅ CLEAN Numbers")
                    clean_df = df[['Alternative Document', 'Date', 'Reason', 'Debit', 'Credit']]
                    st.dataframe(clean_df, use_container_width=True)
                with col2:
                    st.subheader("🔍 Raw vs Clean")
                    debug_df = df[['Raw Document', 'Alternative Document']]
                    st.dataframe(debug_df, use_container_width=True)
                
                # Totals
                try:
                    total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                    total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                    net = round(total_debit - total_credit, 2)
                    st.markdown(f"**💰 Total Debit:** {total_debit:,.2f} | **Total Credit:** {total_credit:,.2f} | **Net:** {net:,.2f}")
                except:
                    pass
                st.download_button(
                    "⬇️ Download CLEAN Excel",
                    data=to_excel_bytes(data),
                    file_name="vendor_statement_clean.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("👆 Upload PDF → Get CLEAN numbers from ΕνδIοNκVο-ιEνοUτ-ι0κ0ό00000001!")
