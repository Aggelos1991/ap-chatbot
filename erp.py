import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from difflib import SequenceMatcher

# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft ERP Translation Audit — Status From Greek vs English Edition")

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
# HELPERS
# ==========================================================
def extract_num(s):
    try:
        n = "".join(ch for ch in str(s) if ch.isdigit() or ch == ".")
        return float(n) if n else 0
    except:
        return 0.0

def quality_icon(score):
    s = extract_num(score)
    if s >= 90: return "🟢 Excellent"
    if s >= 70: return "🟡 Review"
    return "🔴 Poor"

def similarity(a, b):
    return SequenceMatcher(None, str(a).lower().strip(), str(b).lower().strip()).ratio()

def get_status(greek, english):
    """Compare Greek vs English to classify translation status."""
    if not greek or str(greek).lower() == "nan":
        return "Field_Not_Found_On_Report_View"
    if not english or str(english).lower() == "nan":
        return "Field_Not_Translated"
    sim = similarity(greek, english)
    if sim > 0.75:
        return "Translated_Correct"
    elif sim > 0.35:
        return "Translated_Not_Accurate"
    else:
        return "Field_Not_Translated"

# ==========================================================
# BATCH SIZE SELECTOR
# ==========================================================
batch_size = st.slider("⚙️ Select batch size (rows per GPT call):", 10, 200, 50, step=10)
st.caption("Smaller batches are slower but safer. Recommended: 50–100 rows per call.")

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
            rn, rd, fn = str(r["Report_Name"]).strip(), str(r["Report_Description"]).strip(), str(r["Field_Name"]).strip()
            gr, en = str(r["Greek"]).strip(), str(r["English"]).strip()
            if not en or en.lower() == "nan":
                en = ""

            # ---------- Step 1: translate blank immediately ----------
            if not en:
                try:
                    tr = client.chat.completions.create(
                        model=MODEL,
                        messages=[{"role": "user",
                                   "content": f"Translate the following Greek ERP field into proper English ERP/accounting terminology:\n\n{gr}"}],
                        temperature=0
                    )
                    en = tr.choices[0].message.content.strip()
                except Exception as e:
                    st.warning(f"Translation failed at row {i}: {e}")
                    en = "(translation missing)"

            # ---------- Step 2: determine STATUS only from Greek ↔ English ----------
            status = get_status(gr, en)

            # ---------- Step 3: audit quality for corrected term ----------
            prompt = f"""
You are an Entersoft ERP translation auditor.
Greek: {gr}
English: {en}
Determine the correct ERP/accounting English term and return only the improved translation.
"""
            try:
                r2 = client.chat.completions.create(
                    model=MODEL,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0
                )
                corrected = r2.choices[0].message.content.strip()
            except Exception:
                corrected = en

            score = 100 if status == "Translated_Correct" else 85 if status == "Translated_Not_Accurate" else 0
            results.append(dict(
                Report_Name=rn, Report_Description=rd, Field_Name=fn,
                Greek=gr, English=r["English"], Corrected_English=corrected,
                Status=status, Score=score
            ))
            progress.progress(end / total)
            info.write(f"Processed {end}/{total} rows...")

    out = pd.DataFrame(results)
    out["Score"] = out["Score"].apply(extract_num)
    out["Quality"] = out["Score"].apply(quality_icon)
    st.session_state["audit_results"] = out
    st.success("✅ Full audit + status from Greek-English comparison complete.")
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
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button(
        "📥 Download Final Excel (All Corrected)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
