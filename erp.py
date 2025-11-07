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
st.title("🧠 Entersoft ERP Translation Audit — Ultra Fast Edition")

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
    if s >= 90:
        return "🟢 Excellent"
    elif s >= 70:
        return "🟡 Review"
    else:
        return "🔴 Poor"

status_map = {
    "1": "Translated_Correct",
    "2": "Translated_Not_Accurate",
    "3": "Field_Not_Translated",
    "4": "Field_Not_Found_On_Report_View"
}

# ==========================================================
# BATCH SIZE SELECTOR
# ==========================================================
batch_size = st.slider("⚙️ Select batch size (rows per GPT call):", 10, 200, 80, step=10)
st.caption("Larger batches = faster, smaller = safer. Recommended: 80–100 rows per call.")

# ==========================================================
# MAIN AUDIT (ULTRA FAST)
# ==========================================================
if st.button("🚀 Run Full Auto Audit"):
    results = []
    total = len(df)
    progress = st.progress(0)
    info = st.empty()

    for start in range(0, total, batch_size):
        end = min(start + batch_size, total)
        batch = df.iloc[start:end]

        # combine lines for GPT
        lines = []
        for _, r in batch.iterrows():
            rn, rd, fn = str(r["Report_Name"]).strip(), str(r["Report_Description"]).strip(), str(r["Field_Name"]).strip()
            gr, en = str(r["Greek"]).strip(), str(r["English"]).strip()
            if not en or en.lower() == "nan":
                en = ""
            lines.append(f"{rn} | {rd} | {fn} | {gr} | {en}")
        joined = "\n".join(lines)

        prompt = f"""
You are an Entersoft ERP localization auditor and translator.
Compare each Greek field to its English equivalent and classify it conceptually.

If English is missing or wrong, translate Greek into proper ERP/accounting English (e.g. Net Value, VAT Amount, Posting Date).

Statuses:
1=Translated_Correct
2=Translated_Not_Accurate
3=Field_Not_Translated
4=Field_Not_Found_On_Report_View

Scoring:
90–100 Excellent | 70–89 Good | 50–69 Fair | <50 Poor

Reference ERP glossary:
{glossary_text}

Output one line exactly as:
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Score

Analyze now:
{joined}
"""
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            text = resp.choices[0].message.content.strip()
            for ln in text.splitlines():
                p = [x.strip() for x in ln.split("|")]
                if len(p) >= 8:
                    results.append(dict(
                        Report_Name=p[0],
                        Report_Description=p[1],
                        Field_Name=p[2],
                        Greek=p[3],
                        English=p[4],
                        Corrected_English=p[5],
                        Status=status_map.get(p[6], p[6]),
                        Score=p[7]
                    ))
        except Exception as e:
            st.warning(f"Batch {start}-{end} failed: {e}")

        progress.progress(end / total)
        info.write(f"Processed {end}/{total} rows...")

    # build dataframe
    out = pd.DataFrame(results)
    out["Score"] = out["Score"].apply(extract_num)
    out["Quality"] = out["Score"].apply(quality_icon)

    # drop unneeded columns from display
    display_cols = ["Report_Name", "Report_Description", "Field_Name", "Greek", "English", "Corrected_English", "Status", "Quality"]
    st.session_state["audit_results"] = out
    st.success("✅ Ultra-fast audit complete.")
    st.dataframe(out[display_cols].head(30))

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
        "📥 Download Final Excel (All Corrected)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
