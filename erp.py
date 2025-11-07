import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft ERP Translation Audit — Hybrid ERP Intelligence + Speed")

# ==========================================================
# OPENAI
# ==========================================================
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()
client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ==========================================================
# OPTIONAL GLOSSARY
# ==========================================================
def load_glossary(df):
    df.columns = [c.strip().lower() for c in df.columns]
    g = next((c for c in df.columns if "greek" in c or "ελλην" in c), None)
    e = next((c for c in df.columns if "approved" in c or "english" in c), None)
    if g and e:
        return "\n".join([f"{row[g]} → {row[e]}" for _, row in df.iterrows()])
    return ""

glossary_text = ""
upl = st.file_uploader("📘 (Optional) Upload ERP glossary CSV", type=["csv"])
if upl:
    glossary_df = pd.read_csv(upl)
    glossary_text = load_glossary(glossary_df)
elif os.path.exists("erp_glossary.csv"):
    glossary_df = pd.read_csv("erp_glossary.csv")
    glossary_text = load_glossary(glossary_df)
else:
    glossary_text = "(no glossary provided)"

# ==========================================================
# SOURCE EXCEL
# ==========================================================
upl_file = st.file_uploader("📂 Upload Excel (Report_Name | Report_Description | Field_Name | Greek | English)", type=["xlsx"])
if not upl_file:
    st.info("Please upload your exported Excel file from SQL.")
    st.stop()

df = pd.read_excel(upl_file)
st.write(f"✅ File loaded successfully — {len(df)} rows detected.")

if st.checkbox("Run only first 30 rows (test mode)", value=True):
    df = df.head(30)
    st.warning("⚠️ Audit limited to 30 rows for testing.")

req_cols = {"Report_Name", "Report_Description", "Field_Name", "Greek", "English"}
if not req_cols.issubset(df.columns):
    st.error(f"❌ Excel must contain columns: {req_cols}")
    st.stop()

# ==========================================================
# GPT HELPERS
# ==========================================================
def get_status_via_gpt(client, model, greek: str, english_original: str, glossary_text: str) -> str:
    """Judge if the existing English is a correct conceptual translation of the Greek term."""
    g = (greek or "").strip()
    e = (english_original or "").strip()

    if not g or g.lower() == "nan":
        return "Field_Not_Found_On_Report_View"
    if not e or e.lower() == "nan" or e.strip() == "":
        return "Field_Not_Translated"

    prompt = f"""
You are an Entersoft ERP localization auditor.
Judge if the EXISTING English term correctly translates the Greek term in ERP/accounting context.

Return ONLY one label:
Translated_Correct
Translated_Not_Accurate
Field_Not_Translated
Field_Not_Found_On_Report_View

Greek: {g}
Existing English: {e}

Glossary reference (if any):
{glossary_text}
"""
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        label = resp.choices[0].message.content.strip()
        allowed = {
            "Translated_Correct",
            "Translated_Not_Accurate",
            "Field_Not_Translated",
            "Field_Not_Found_On_Report_View",
        }
        return label if label in allowed else "Translated_Not_Accurate"
    except Exception:
        return "Translated_Not_Accurate"


def get_quality_label(client, model, greek: str, corrected: str) -> str:
    """Assess conceptual translation quality between Greek and Corrected English."""
    g = (greek or "").strip()
    c = (corrected or "").strip()
    if not g or not c:
        return "🟡 Review"

    prompt = f"""
Judge conceptual translation quality for the ERP/accounting context.

Greek: {g}
English: {c}

Return EXACTLY one of:
🟢 Excellent
🟡 Review
🔴 Poor
"""
    try:
        r = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        out = r.choices[0].message.content.strip()
        return out if out in {"🟢 Excellent","🟡 Review","🔴 Poor"} else "🟡 Review"
    except Exception:
        return "🟡 Review"

# ==========================================================
# BATCH SIZE
# ==========================================================
batch_size = st.slider("⚙️ Select batch size (rows per GPT call):", 10, 200, 50, step=10)
st.caption("Larger batches = faster. Recommended: 50–100 rows per call.")

# ==========================================================
# MAIN AUDIT
# ==========================================================
if st.button("🚀 Run Full Auto Audit"):
    results = []
    total = len(df)
    progress = st.progress(0)
    info = st.empty()

    for start in range(0, total, batch_size):
        end = min(start + batch_size, total)
        batch = df.iloc[start:end]

        for i, r in batch.iterrows():
            rn = str(r["Report_Name"]).strip()
            rd = str(r["Report_Description"]).strip()
            fn = str(r["Field_Name"]).strip()
            gr = str(r["Greek"]).strip()
            en_orig = str(r["English"]).strip()

            # Step 1: Determine STATUS based on Greek ↔ original English
            status = get_status_via_gpt(client, MODEL, gr, en_orig, glossary_text)

            # Step 2: Generate improved ERP-corrected English (Hybrid Expert Mode)
            seed_en = en_orig if en_orig and en_orig.lower() != "nan" else ""
            try:
                fix = client.chat.completions.create(
                    model=MODEL,
                    messages=[{"role": "user",
                               "content": f"""
You are a senior ERP and accounting localization expert specialized in Entersoft ERP.
Improve or translate the following field name into precise, concise ERP English used in real systems.
Prefer standard ERP terminology such as:
Net Value, Posting Date, Credit Note, Cost Center, Ledger Account, VAT Amount, Warehouse, Supplier, Customer, Invoice Number, Payment Method, Transaction Date, etc.

Maintain capitalization in Title Case.
Return only the corrected ERP English label.

Greek: {gr}
Existing English: {seed_en}
"""}],
                    temperature=0
                )
                corrected = fix.choices[0].message.content.strip()
            except Exception:
                corrected = seed_en

            # Step 3: Quality between Greek ↔ Corrected English
            quality = get_quality_label(client, MODEL, gr, corrected)

            results.append(dict(
                Report_Name=rn,
                Report_Description=rd,
                Field_Name=fn,
                Greek=gr,
                English=en_orig,
                Corrected_English=corrected,
                Status=status,
                Quality=quality
            ))

            progress.progress(end / total)
            info.write(f"Processed {end}/{total} rows...")

    out = pd.DataFrame(results)
    st.session_state["audit_results"] = out
    st.success("✅ Full audit complete (Hybrid ERP Intelligence + Speed).")
    st.dataframe(out.head(30))

# ==========================================================
# EXPORT
# ==========================================================
if "audit_results" in st.session_state:
    out = st.session_state["audit_results"]
    wb = Workbook()
    ws = wb.active
    ws.title = "ERP Translation Audit"
    ws.append(list(out.columns))
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
    for _, r in out.iterrows():
        ws.append([r[col] for col in out.columns])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(
            max(len(str(c.value or "")) for c in col) + 2, 60
        )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.download_button(
        "📥 Download Final Excel (Status + Quality)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
