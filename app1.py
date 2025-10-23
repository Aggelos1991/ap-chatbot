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
            for line in text.split("\n"):
                if re.search(r"\d{1,3}(?:[.,]\d{3})*[.,]?\d{0,2}", line):
                    all_lines.append(" ".join(line.split()))
    return all_lines

def extract_with_gpt(lines):
    BATCH_SIZE = 100
    all_records = []
    
    for i in range(0, len(lines), BATCH_SIZE):
        batch = lines[i:i + BATCH_SIZE]
        text_block = "\n".join(batch)
        
        prompt = (
            "You are extracting accounting transactions from vendor statements. "
            "CRITICAL: Find EVERY document number accurately.\n\n"
            "DOCUMENT NUMBERS appear as:\n"
            "Spanish: Nº 12345, Num. 678, Factura 001234, Fra 2024/001, Ref 456, Documento 789\n"
            "Greek: Τιμολόγιο 12345, Αρ. 67890, ΤΛ 001, Παραστατικό 456\n"
            "Numbers: 123, 12345, 2024-001, 24/001, 001234\n\n"
            "For EACH transaction line with a document number, extract:\n"
            "1. Alternative Document: EXACT document number ONLY (12345, FRA001, ΤΛ678)\n"
            "2. Date: dd/mm/yy or dd/mm/yyyy\n"
            "3. Debit: number from DEBE column (keep negative if present)\n"
            "4. Credit: number from HABER column (keep negative if present)\n"
            "5. Reason: Invoice, Payment, or Credit Note\n"
            "6. Description: short text\n\n"
            "Rules:\n"
            "- DEBE > 0 = Invoice\n"
            "- HABER > 0 = Payment\n"
            "- DEBE < 0 or NC = Credit Note\n"
            "- NEVER extract lines with 'concil', 'total', 'saldo', 'iva'\n\n"
            "Return ONLY JSON array:\n"
            '[{"Alternative Document":"12345","Date":"01/10/24","Debit":"1234.56","Credit":"","Reason":"Invoice","Description":"Factura emitida"}]'
            "\n\nText:\n" + text_block
        )
        
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.0
            )
            content = response.choices[0].message.content.strip()
            
            # Extract JSON more reliably
            json_start = content.find('[')
            json_end = content.rfind(']') + 1
            if json_start != -1 and json_end > json_start:
                json_str = content[json_start:json_end]
                data = json.loads(json_str)
            else:
                continue
                
        except:
            continue
        
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            
            if not alt_doc or not re.search(r"\d", alt_doc):
                continue
            
            if re.search(r"concil|total|saldo|iva|apertur|cierre", alt_doc, re.IGNORECASE):
                continue
            
            debit_val = normalize_number(row.get("Debit"))
            credit_val = normalize_number(row.get("Credit"))
            reason = row.get("Reason", "Invoice").strip()
            
            # Handle negatives
            if debit_val and float(debit_val) < 0:
                credit_val = abs(float(debit_val))
                debit_val = ""
                reason = "Credit Note"
            elif credit_val and float(credit_val) < 0:
                debit_val = abs(float(credit_val))
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
        st.text_area("📄 Preview:", "\n".join(lines[:25]), height=250)
        
        col1, col2 = st.columns([3,1])
        with col1:
            if st.button("🤖 Extract Documents", type="primary"):
                with st.spinner("🔍 Analyzing..."):
                    data = extract_with_gpt(lines)
                
                if data:
                    df = pd.DataFrame(data)
                    st.success(f"✅ {len(df)} records extracted!")
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
                        file_name=f"extraction_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ No documents found.")
        
        with col2:
            st.metric("Lines", len(lines))

else:
    st.info("👆 Upload PDF to extract documents")
