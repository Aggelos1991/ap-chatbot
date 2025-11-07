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
st.title("🧠 Entersoft ERP Translation Audit — Status + Quality Edition")

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
def similarity(a, b):
    return SequenceMatcher(None, str(a).lower().strip(), str(b).lower().strip()).ratio()

def get_status(greek, english):
    """Compare Greek vs English to classify translation status."""
    if not greek or str(greek).lower() == "nan":
        return "Field_Not_Found_On_Report_View"
    if not english or str(english).lower() == "nan" or english.strip() == "":
        return "Field_Not_Translated"
    sim = similarity(greek, english)
    if sim > 0.75:
        return "Translated_Correct"
    elif sim > 0.35:
        return "Translated_Not_Accurate"
    else:
        return "Field_Not_Translated"

def get_quality_label(greek, corrected):
    """Ask GPT to rate conceptual translation quality."""
    try:
        prompt = f"""
Judge the conceptual translation quality between these two ERP fields:
Greek: {greek}
English: {corrected}

Choose only one option and return exactly it:
🟢 Excellent
🟡 Review
🔴 Poor
"""
        r = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        return r.choices[0].message.content.strip()
    except:
        return "🟡 Review"

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

            # ---------- Step 1: auto-translate blank English ----------
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

            # ---------- Step 2: STATUS (Greek ↔ English) ----------
            status = get_status(gr, en)

            # ---------- Step 3: Corrected English ----------
            try:
                fix = client.chat.completions.create(
                    model=MODEL,
                    messages=[{"role": "user",
                               "content": f"Improve this ERP translation if needed to a correct professional English ERP/accounting term:\nGreek: {gr}\nEnglish: {en}\nReturn only the corrected English term."}],
                    temperature=0
                )
                corrected = fix.choices[0].message.content.strip()
            except Exception:
                corrected = en

            results.append(dict(
                Report_Name=rn, Report_Description=rd, Field_Name=fn,
                Greek=gr, English=r["English"], Corrected_English=corrected,
                Status=status
            ))

            progress.progress(end / total)
            info.write(f"Processed {end}/{total} rows...")

    out = pd.DataFrame(results)

    # ---------- Step 4: QUALITY (Greek ↔ Corrected English) ----------
    st.info("🔍 Assessing conceptual quality (Greek ↔ Corrected English)...")
    out["Quality"] = [get_quality_label(row["Greek"], row["Corrected_English"]) for _, row in out.iterrows()]

    st.session_state["audit_results"] = out
    st.success("✅ Full audit complete (Status + Quality).")
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
        "📥 Download Final Excel (Status + Quality)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
