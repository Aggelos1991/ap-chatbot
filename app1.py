import os, re, json, platform, shutil
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
import fitz  # PyMuPDF for OCR fallback
import pytesseract
from PIL import Image

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
# TESSERACT CHECK
# ==========================================================
def set_windows_tesseract_path_if_exists():
    """On Windows, set pytesseract cmd to default install path if present."""
    if platform.system().lower().startswith("win"):
        possible = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(possible):
            pytesseract.pytesseract.tesseract_cmd = possible  # ✅ ensure always set

def has_tesseract():
    """Return True if tesseract is available on PATH or at known Windows path."""
    set_windows_tesseract_path_if_exists()
    if shutil.which(getattr(pytesseract.pytesseract, "tesseract_cmd", "tesseract")) is None:
        return False
    try:
        _ = pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False

TESS_AVAILABLE = has_tesseract()

# Show message depending on OCR availability
if TESS_AVAILABLE:
    st.info("✅ OCR engine (Tesseract) detected and active — scanned PDFs supported.")
else:
    st.warning(
        "🔎 OCR is disabled because **Tesseract** is not installed or not on PATH.\n\n"
        "Install it and refresh the app to enable OCR:\n"
        "- **macOS (Homebrew):** `brew install tesseract`\n"
        "- **Ubuntu/Debian:** `sudo apt update && sudo apt install -y tesseract-ocr tesseract-ocr-spa tesseract-ocr-ell`\n"
        "- **Windows:** Install from the UB Mannheim build, then restart the app.\n"
        "Common path: `C:\\Program Files\\Tesseract-OCR\\tesseract.exe` (auto-detected).",
        icon="⚠️"
    )

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
    """Extract ALL text lines from every page of the PDF, with OCR fallback (if available)."""
    all_lines = []
    raw_bytes = uploaded_pdf.read()
    uploaded_pdf.seek(0)

    with pdfplumber.open(BytesIO(raw_bytes)) as pdf:
        doc = fitz.open(stream=raw_bytes, filetype="pdf") if TESS_AVAILABLE else None

        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    clean_line = " ".join(line.split())
                    if clean_line.strip():
                        all_lines.append(clean_line)
                continue

            # OCR fallback if page has no text
            if TESS_AVAILABLE and doc is not None:
                try:
                    pix = doc.load_page(i).get_pixmap()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    ocr_text = pytesseract.image_to_string(img, lang="eng+spa+ell")
                    for line in ocr_text.split("\n"):
                        clean_line = " ".join(line.split())
                        if clean_line.strip():
                            all_lines.append(clean_line)
                except Exception as e:
                    st.warning(f"OCR failed on page {i+1}: {e}")

        if doc is not None:
            doc.close()

    return all_lines

# ==========================================================
# GPT EXTRACTOR — FIXED CREDIT NOTE + FILTER HANDLING
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
- N° DOC → Document number (1729, 1775, etc.) , COMENTARIO -> You can find here sometimes the invoice number
- DEBE → Invoice amounts (Debit)
- HABER/CREDIT → Payment amounts (Credit) 
- SALDO → Running balance (IGNORE for extraction)
- Don't count Asiento for Document number

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

            if i == 0:
                st.text_area("GPT Response (Batch 1):", content, height=200, key="debug_1")

            json_match = re.search(r'\[.*\]', content, re.DOTALL)
            if not json_match:
                json_match = re.search(r'(\[.*?\])', content, re.DOTALL)

            if json_match:
                json_str = json_match.group(0)
                data = json.loads(json_str)

                for row in data:
                    alt_doc = str(row.get("Alternative Document", "")).strip()

                    # FILTER: skip Asiento, Saldo, Comentario, Total, IVA
                    if not alt_doc or re.search(r"(asiento|saldo|comentario|total|iva)", alt_doc, re.IGNORECASE):
                        continue

                    debit_raw = row.get("Debit", "")
                    credit_raw = row.get("Credit", "")
                    debit_val = normalize_number(debit_raw)
                    credit_val = normalize_number(credit_raw)
                    reason = row.get("Reason", "Invoice").strip()

                    # Negative DEBE → Credit Note
                    if debit_val != "" and float(debit_val) < 0:
                        credit_val = abs(float(debit_val))
                        debit_val = ""
                        reason = "Credit Note"

                    if reason == "Payment" and credit_val != "" and float(credit_val) > 0:
                        pass
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
    with st.spinner("📄 Extracting text (OCR used if available)..."):
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

            try:
                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)

                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("⚖️ Net", f"{net:,.2f}")
            except Exception as e:
                st.error(f"Totals error: {e}")

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(data),
                file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No structured data detected. Check GPT response above.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
