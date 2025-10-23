import os, re, json
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI

st.set_page_config(page_title="🦅 DataFalcon Pro", layout="wide")
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

def normalize_number(value):
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
        num = float(s)
        return round(num, 2)
    except:
        return ""

def extract_raw_lines(uploaded_pdf):
    all_lines = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split("\n")
            for line in lines:
                clean_line = " ".join(line.split())
                if clean_line.strip():
                    all_lines.append(clean_line)
    return all_lines

def extract_with_gpt(lines):
    BATCH_SIZE = 50
    all_records = []
    
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        
        prompt = (
            "Extract ONLY lines that contain DOCUMENT NUMBERS. "
            "NEVER use DEBE/HABER amounts as document numbers.\n\n"
            "DOCUMENT = Factura number, Invoice number, Ref number\n\n"
            "LOOK FOR THESE PATTERNS:\n"
            "• Nº 12345, Num 678, Factura 001, Fra 2024/001\n"
            "• Τιμολόγιο 123, Αρ. 456, ΤΛ 789, Παραστατικό 001\n"
            "• 12345, 2024-001, 24/123, INV001\n\n"
            "DOCUMENT RULES:\n"
            "1. Must be 3-12 DIGITS or with prefix (Nº, Fra, ΤΛ)\n"
            "2. NEVER extract DEBE or HABER amounts (1.234,56 = WRONG)\n"
            "3. Skip if no clear document identifier\n\n"
            "For EACH valid document line extract:\n"
            '{"Alternative Document": "12345", "Date": "01/10/24", "Debit": "1234.56", "Credit": "", "Reason": "Invoice", "Description": "text"}'
            "\n\nONLY return valid JSON array for lines with DOCUMENTS:\n" + text_block
        )
        
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.0
            )
            content = response.choices[0].message.content.strip()
            
            json_start = content.find('[')
            json_end = content.rfind(']') + 1
            if json_start == -1 or json_end <= json_start:
                continue
                
            json_str = content[json_start:json_end]
            data = json.loads(json_str)
            
        except:
            continue
        
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            
            # STRICT document validation
            if not alt_doc:
                continue
                
            # Must have digits AND be reasonable document length
            if not re.search(r"\d{3,}", alt_doc):
                continue
                
            # NEVER allow decimal amounts as documents
            if re.search(r"[.,]\d{2}$", alt_doc):
                continue
                
            # Block common exclusion words
            exclude_words = ['concil', 'total', 'saldo', 'iva', 'apertur', 'cierre']
            if any(word in alt_doc.lower() for word in exclude_words):
                continue
            
            debit_val = normalize_number(row.get("Debit"))
            credit_val = normalize_number(row.get("Credit"))
            reason = row.get("Reason", "Invoice").strip()
            
            # Handle negatives
            if debit_val and isinstance(debit_val, (int, float)) and debit_val < 0:
                credit_val = abs(debit_val)
                debit_val = ""
                reason = "Credit Note"
            elif credit_val and isinstance(credit_val, (int, float)) and credit_val < 0:
                debit_val = abs(credit_val)
                credit_val = ""
                reason = "Invoice"
            
            all_records.append({
                "Alternative Document": alt_doc,
                "Date": str(row.get("Date", "")).strip(),
                "Reason": reason,
                "Debit": debit_val,
                "Credit": credit_val,
                "Description": str(row.get("Description", "")).strip()
            })
    return all_records

def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text..."):
        lines = extract_raw_lines(uploaded_pdf)
    
    if not lines:
        st.warning("⚠️ No readable text found.")
    else:
        st.text_area("📄 Preview:", "\n".join(lines[:30]), height=300)
        
        col1, col2 = st.columns([3,1])
        with col1:
            if st.button("🤖 Extract Documents", type="primary"):
                with st.spinner("🔍 Finding documents..."):
                    data = extract_with_gpt(lines)
                
                if data:
                    df = pd.DataFrame(data)
                    st.success(f"✅ {len(df)} documents extracted!")
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    
                    df_num = df.copy()
                    df_num["Debit"] = pd.to_numeric(df_num["Debit"], errors="coerce")
                    df_num["Credit"] = pd.to_numeric(df_num["Credit"], errors="coerce")
                    
                    total_debit = df_num["Debit"].sum()
                    total_credit = df_num["Credit"].sum()
                    net = round(total_debit - total_credit, 2)
                    
                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("💰 Debit", f"{total_debit:,.2f}")
                    col_b.metric("💳 Credit", f"{total_credit:,.2f}")
                    col_c.metric("⚖️ Net", f"{net:,.2f}")
                    
                    st.download_button(
                        "⬇️ Download Excel",
                        data=to_excel_bytes(data),
                        file_name=f"documents_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ No valid documents found. Check preview for document numbers.")
        
        with col2:
            st.metric("Lines", len(lines))

else:
    st.info("👆 Upload PDF")
