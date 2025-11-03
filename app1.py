# ==========================================================
# 🦅 DataFalcon Pro — Hybrid GPT + OCR Vendor Statement Extractor (FINAL macOS FIX)
# ==========================================================
import os, re, json, shutil, time
import pdfplumber
import pandas as pd
import streamlit as st
from io import BytesIO
from openai import OpenAI
from PIL import Image

# --- OCR / Image Handling ---
import pytesseract
from pdf2image import convert_from_bytes

# ==========================================================
# POPPLER & TESSERACT CHECKS
# ==========================================================
POPPLER_PATH = "/opt/homebrew/bin"  # Update if Poppler installed elsewhere
os.environ["PATH"] = POPPLER_PATH + os.pathsep + os.environ["PATH"]

def check_dependencies():
    """Ensure Poppler and Tesseract are installed and accessible."""
    st.info("🧩 Checking OCR dependencies...")
    pdftoppm_path = shutil.which("pdftoppm")
    tesseract_path = shutil.which("tesseract")

    if not pdftoppm_path:
        st.error("❌ Poppler not found. Install via: brew install poppler")
        st.stop()

    if not tesseract_path:
        st.error("❌ Tesseract not found. Install via: brew install tesseract")
        st.stop()

    st.success(f"✅ Poppler: {pdftoppm_path}")
    st.success(f"✅ Tesseract: {tesseract_path}")

check_dependencies()

# ==========================================================
# STREAMLIT CONFIG
# ==========================================================
st.set_page_config(page_title="🦅 DataFalcon Pro — Hybrid GPT Extractor", layout="wide")
st.title("🦅 DataFalcon Pro — Hybrid GPT + OCR Extractor")

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
# SAFE OCR EXTRACTION (Timeout Protected)
# ==========================================================
def safe_convert_from_bytes(pdf_bytes, timeout=60):
    """Convert PDF → images with timeout to avoid infinite hang."""
    from multiprocessing import Process, Queue
    q = Queue()

    def _worker(q):
        try:
            imgs = convert_from_bytes(
                pdf_bytes, dpi=300, fmt="png", poppler_path=POPPLER_PATH
            )
            q.put(imgs)
        except Exception as e:
            q.put(e)

    p = Process(target=_worker, args=(q,))
    p.start()
    p.join(timeout)
    if p.is_alive():
        p.terminate()
        raise TimeoutError("Poppler conversion timed out — check Poppler install.")

    result = q.get()
    if isinstance(result, Exception):
        raise result
    return result


def ocr_extract(pdf_bytes):
    """Perform OCR extraction safely"""
    lines = []
    try:
        st.info("📸 Starting OCR extraction (Poppler + Tesseract)...")
        images = safe_convert_from_bytes(pdf_bytes)
        for i, img in enumerate(images):
            with st.status(f"OCR Page {i+1}/{len(images)}") as status:
                status.write(f"Reading page {i+1}…")
                text = pytesseract.image_to_string(img, lang="spa+eng", config="--psm 6")
                for line in text.split("\n"):
                    if line.strip():
                        lines.append(line.strip())
                status.update(label=f"Page {i+1} done ✅", state="complete")
        st.success(f"OCR finished — {len(lines)} lines extracted.")
    except Exception as e:
        st.error(f"❌ OCR failed: {e}")
        st.info(f"🧩 Poppler tried from: {POPPLER_PATH}")
    return lines

# ==========================================================
# HYBRID TEXT + OCR EXTRACTION
# ==========================================================
def extract_raw_lines(uploaded_pdf):
    """Extract text from PDF; fallback to OCR if needed."""
    all_lines = []
    pdf_bytes = uploaded_pdf.getvalue()

    with pdfplumber.open(uploaded_pdf) as pdf:
        sample_text = any(page.extract_text() for page in pdf.pages[:3] if page.extract_text())

    if sample_text:
        st.info("📄 Detected searchable PDF → using fast text extraction")
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
        st.warning("📸 No text layer found → switching to OCR (slower, 1–3 min)")
        all_lines = ocr_extract(pdf_bytes)

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
    with st.spinner("📄 Extracting text (OCR-enabled)..."):
        lines = extract_raw_lines(uploaded_pdf)

    st.success(f"✅ Found {len(lines)} lines of text!")
    st.text_area("📋 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Hybrid Extraction", type="primary"):
        with st.spinner("Analyzing with GPT models..."):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)
            st.success(f"✅ Extraction complete — {len(df)} valid records found!")
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
