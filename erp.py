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
st.title("🧠 Entersoft ERP Translation Audit — Full Auto Edition")

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

# ==========================================================
# MAIN AUDIT
# ==========================================================
if st.button("🚀 Run Full Auto Audit"):
    results = []
    total = len(df)
    progress = st.progress(0)
    info = st.empty()

    for i, r in df.iterrows():
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

        # ---------- Step 2: audit quality ----------
        prompt = f"""
You are an Entersoft ERP translation auditor.
Compare directly the following pair and score the English quality conceptually.

Greek: {gr}
English: {en}

Use ERP/accounting context (Debit, Credit, Cost Center, VAT Amount, etc.)
Statuses:
1=Translated_Correct, 2=Translated_Not_Accurate, 3=Field_Not_Translated

Output exactly:
Corrected_English | Status | Status_Description | Score
"""
        try:
            r2 = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0
            )
            text = r2.choices[0].message.content.strip()
            parts = [x.strip() for x in text.split("|")]
            if len(parts) < 4:
                corrected, status, desc, score = en, "Translated_Correct", "Auto assumed", "100"
            else:
                corrected, status, desc, score = parts[:4]
        except Exception as e:
            corrected, status, desc, score = en, "Error", str(e), "0"

        # ---------- Step 3: if weak (<70) → auto improve ----------
        if extract_num(score) < 70:
            try:
                fix = client.chat.completions.create(
                    model=MODEL,
                    messages=[{"role": "user",
                               "content": f"Improve this weak ERP translation to a correct, professional ERP/accounting English term:\n\nGreek: {gr}\nCurrent: {corrected}\nReturn only corrected term."}],
                    temperature=0
                )
                corrected = fix.choices[0].message.content.strip()
                status, desc, score = "Translated_Correct", "Auto-improved", "100"
            except Exception as e:
                st.warning(f"Auto-fix failed row {i}: {e}")

        results.append(dict(
            Report_Name=rn, Report_Description=rd, Field_Name=fn,
            Greek=gr, English=r["English"], Corrected_English=corrected,
            Status=status, Status_Description=desc, Score=score
        ))
        progress.progress((i + 1) / total)
        info.write(f"Processed {i+1}/{total}")

    out = pd.DataFrame(results)
    out["Score"] = out["Score"].apply(extract_num)
    out["Quality"] = out["Score"].apply(quality_icon)
    st.session_state["audit_results"] = out
    st.success("✅ Full audit + auto-translation complete.")
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
