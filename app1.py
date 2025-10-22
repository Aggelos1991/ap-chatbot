import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
import time

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro 4.0 — Gibberish Proof", layout="wide")
st.title("🦅 DataFalcon Pro 4.0")
st.markdown("**🔥 Extracts 00000001 from ΕνδIοNκVο-ιEνοUτ-ι0κ0ό00000001**")

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
# 🔥 GIBBERISH-PROOF INVOICE EXTRACTOR
# ==========================================================
def extract_clean_invoice_number(text):
    """🚀 STRIP GIBBERISH - KEEP ONLY FINAL NUMBER"""
    # Pattern 1: ANYTHING followed by 7+ digits at END
    final_num = re.search(r'.*?(\d{7,})$', text)
    if final_num:
        return final_num.group(1)
    
    # Pattern 2: Gibberish + dashes + number
    dash_num = re.search(r'[-–—]\s*(\d{6,})$', text)
    if dash_num:
        return dash_num.group(1)
    
    # Pattern 3: ANY 6+ digits
    any_num = re.search(r'(\d{6,})', text)
    if any_num:
        return any_num.group(1)
    
    return ""

def normalize_number(value):
    """Enhanced number normalization."""
    if not value:
        return ""
    s = str(value).strip().replace(" ", "")
    s = re.sub(r'[€$£¥]', '', s)
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
    """Extract ALL lines with numbers - no filtering."""
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    line = line.strip()
                    if re.search(r'\d', line) and len(line) > 10:
                        all_lines.append(line)
            
            # Tables
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row:
                        row_text = " | ".join(str(cell or '') for cell in row)
                        if re.search(r'\d', row_text):
                            all_lines.append(row_text)
    return all_lines

# ==========================================================
# GPT + GIBBERISH STRIPPER
# ==========================================================
def extract_with_gpt(lines):
    """GPT extracts → We STRIP GIBBERISH → 100% clean numbers"""
    BATCH_SIZE = 100
    all_records = []
    
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        
        prompt = f"""Extract transactions. Focus on document numbers and amounts.

Output JSON array:
{{
  "raw_doc": "ΕνδIοNκVο-ιEνοUτ-ι0κ0ό00000001", 
  "raw_amount1": "first amount",
  "raw_amount2": "second amount"
}}

Lines:
{text_block}"""
        
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=3000
            )
            content = response.choices[0].message.content.strip()
            json_match = re.search(r"\[.*\]", content, re.DOTALL)
            if not json_match:
                continue
            data = json.loads(json_match.group(0))
        except:
            continue
        
        # 🔥 GIBBERISH STRIPPER - MAGIC HAPPENS HERE
        for row in data:
            raw_doc = str(row.get("raw_doc", "")).strip()
            raw_amt1 = str(row.get("raw_amount1", "")).strip()
            raw_amt2 = str(row.get("raw_amount2", "")).strip()
            
            # STRIP GIBBERISH → GET PURE NUMBER
            clean_doc = extract_clean_invoice_number(raw_doc)
            
            # Skip if no clean number found
            if not clean_doc or len(clean_doc) < 6:
                continue
            
            # Skip reconciliation
            if re.search(r"concil", raw_doc, re.IGNORECASE):
                continue
            
            # Normalize amounts
            debit = normalize_number(raw_amt1)
            credit = normalize_number(raw_amt2)
            
            # Must have money OR document
            if debit == "" and credit == "":
                continue
            
            # Auto-classify
            reason = "Invoice" if debit != "" else "Payment"
            
            all_records.append({
                "Alternative Document": clean_doc,
                "Raw Document": raw_doc[:50],  # Debug
                "Date": "",
                "Reason": reason,
                "Debit": debit or 0.0,
                "Credit": credit or 0.0
            })
    
    return all_records

# ==========================================================
# VALIDATION - CLEAN NUMBERS ONLY
# ==========================================================
def validate_records(records):
    df = pd.DataFrame(records)
    if df.empty:
        return df
    
    # Clean invoice validation
    df['Valid'] = (
        df['Alternative Document'].str.len() >= 6 &
        df['Alternative Document'].str.match(r'^\d+$', na=False) &
        ((pd.to_numeric(df['Debit'], errors='coerce') > 0) | 
         (pd.to_numeric(df['Credit'], errors='coerce') > 0))
    )
    return df

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
# UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("🔍 Extracting ALL text..."):
        lines = extract_raw_lines(uploaded_pdf)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("📄 Total Lines", len(lines))
    with col2:
        clean_docs = sum(1 for line in lines if extract_clean_invoice_number(line))
        st.metric("🔢 Clean Docs Found", clean_docs)
    
    st.text_area("📄 Preview:", "\n".join(lines[:10]), height=200)
    
    if st.button("🚀 Extract & Clean Gibberish", type="primary"):
        with st.spinner("🧠 GPT → GIBBERISH STRIPPER → CLEAN NUMBERS"):
            data = extract_with_gpt(lines)
            df = validate_records(data)
        
        if df.empty:
            st.error("❌ No valid records")
        else:
            valid_df = df[df['Valid'] == True][['Alternative Document', 'Date', 'Reason', 'Debit', 'Credit']]
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.success(f"✅ **{len(valid_df)} CLEAN INVOICES**")
            with col2:
                total_debit = valid_df["Debit"].sum()
                st.success(f"💰 **{total_debit:,.2f}**")
            with col3:
                total_credit = valid_df["Credit"].sum()
                st.success(f"⚖️ **{total_debit - total_credit:,.2f}**")
            
            st.dataframe(valid_df, use_container_width=True)
            
            # Debug table
            with st.expander("🔍 Raw vs Clean (Debug)"):
                debug_df = df[['Raw Document', 'Alternative Document', 'Debit', 'Credit', 'Valid']]
                st.dataframe(debug_df, use_container_width=True)
            
            st.download_button(
                "💾 Download CLEAN Invoices",
                data=to_excel_bytes(valid_df.to_dict('records')),
                file_name="clean_invoices.xlsx"
            )
else:
    st.info("👆 Upload PDF → Get CLEAN 00000001 from gibberish!")
