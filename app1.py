import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io, BytesIO
from openai import OpenAI
import time

# ==========================================================
# CONFIGURATION
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro DEBUG", layout="wide")
st.title("🦅 DataFalcon Pro - DEBUG MODE")
st.markdown("**Finding why no invoices detected...**")

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
# DEBUG HELPERS - LESS STRICT
# ==========================================================
def extract_invoice_number(line):
    """🚀 Extract ANY number - super lenient"""
    # Pattern 1: ANY 4+ digits
    num = re.search(r'\b(\d{4,})\b', line)
    if num:
        return num.group(1)
    
    # Pattern 2: With prefixes
    patterns = [
        r'(?:inv|fact|fra|n\d?|num|doc|ref|tim|par|αρ)\s*[:\-/]*\s*(\d{3,})',
        r'(\d{3,})\s*(?:inv|fact|fra|tim|par)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, line, re.IGNORECASE)
        if match:
            return match.group(1)
    
    return ""

def normalize_number(value):
    if not value:
        return 0.0
    s = re.sub(r'[^\d,\.€$£ ]', '', str(value).strip())
    s = s.replace(' ', '').replace('€', '')
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return round(float(s), 2)
    except:
        return 0.0

def extract_raw_lines(uploaded_pdf):
    """Extract ALL lines with ANY number"""
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    line = line.strip()
                    if re.search(r'\d', line) and len(line) > 5:
                        all_lines.append(line)
            
            # Tables
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row:
                        row_text = " | ".join(str(cell) for cell in row if cell)
                        if re.search(r'\d', row_text):
                            all_lines.append(row_text)
    return all_lines[:500]  # Limit for speed

# ==========================================================
# GPT - SIMPLIFIED FOR DEBUG
# ==========================================================
def extract_with_gpt(lines):
    BATCH_SIZE = 80
    all_records = []
    
    st.info(f"🔍 Processing {len(lines)} lines in {len(lines)//BATCH_SIZE + 1} batches...")
    
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        
        prompt = f"""Extract ALL lines with money amounts. Output JSON array.

For EVERY line with numbers:
{{
  "raw_line": "exact line text",
  "doc_num": "largest number found",
  "amount1": "first money amount",
  "amount2": "second money amount" 
}}

TEXT:
{text_block}"""
        
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=2000
            )
            content = response.choices[0].message.content.strip()
            json_match = re.search(r"\[.*\]", content, re.DOTALL)
            if json_match:
                data = json.loads(json_match.group(0))
                all_records.extend(data)
        except Exception as e:
            st.warning(f"Batch {i//BATCH_SIZE + 1} error: {e}")
        
        if i % 200 < BATCH_SIZE:
            st.write(f"Processed batch {i//BATCH_SIZE + 1}")
    
    return all_records

# ==========================================================
# DEBUG VALIDATION - SHOW EVERYTHING
# ==========================================================
def analyze_records(records):
    """Show exactly what's happening"""
    df = pd.DataFrame(records)
    if df.empty:
        return pd.DataFrame()
    
    # Extract numbers from raw lines
    df['doc_num'] = df['raw_line'].apply(extract_invoice_number)
    df['money_found'] = df['raw_line'].str.contains(r'\d+[.,]\d{2}', na=False)
    df['has_doc'] = df['doc_num'].str.len() > 3
    df['amount1_num'] = df['amount1'].apply(normalize_number)
    df['amount2_num'] = df['amount2'].apply(normalize_number)
    df['has_money'] = (df['amount1_num'] > 0) | (df['amount2_num'] > 0)
    df['valid'] = df['has_doc'] & df['has_money']
    
    return df

# ==========================================================
# UI - FULL DEBUG
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload PDF", type=["pdf"])

if uploaded_pdf:
    st.header("🔍 **DEBUG ANALYSIS**")
    
    with st.spinner("Extracting ALL lines..."):
        lines = extract_raw_lines(uploaded_pdf)
    
    st.subheader("📄 Raw Lines Found")
    st.text_area("Sample:", "\n".join(lines[:20]), height=200)
    st.metric("Total Lines", len(lines))
    
    # Show invoice numbers found
    invoice_lines = [l for l in lines if extract_invoice_number(l)]
    st.metric("💼 Lines with Invoice #", len(invoice_lines))
    
    if st.button("🔬 FULL GPT DEBUG EXTRACTION", type="primary"):
        with st.spinner("Running GPT on every line..."):
            raw_data = extract_with_gpt(lines)
            df = analyze_records(raw_data)
        
        st.header("📊 **EXTRACTION RESULTS**")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📄 Processed", len(df))
        with col2:
            st.metric("💰 Money Found", len(df[df['money_found']]))
        with col3:
            st.metric("📱 Invoice # Found", len(df[df['has_doc']]))
        with col4:
            st.metric("✅ VALID", len(df[df['valid']]))
        
        st.subheader("🔍 DEBUG TABLE - See Everything")
        st.dataframe(df[['raw_line', 'doc_num', 'amount1', 'amount2', 'has_doc', 'has_money', 'valid']], 
                    use_container_width=True, height=400)
        
        # VALID RECORDS
        valid_records = df[df['valid'] == True].copy()
        if not valid_records.empty:
            st.success(f"🎉 **{len(valid_records)} VALID INVOICES FOUND!**")
            final_df = valid_records[['doc_num', 'raw_line']].rename(columns={'doc_num': 'Invoice', 'raw_line': 'Description'})
            st.dataframe(final_df, use_container_width=True)
            
            # Download
            buf = BytesIO()
            final_df.to_excel(buf, index=False)
            buf.seek(0)
            st.download_button("💾 Download", buf.getvalue(), "debug_invoices.xlsx")
        else:
            st.error("❌ NO VALID INVOICES - Check DEBUG table above")
            st.info("👆 Look at 'has_doc' and 'has_money' columns to see what's failing")

else:
    st.info("Upload PDF to see FULL DEBUG analysis")
