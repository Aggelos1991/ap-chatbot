import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os, time
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# ========== CONFIG ==========
st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft AI Translation Audit — ERP Expert Edition")

# ========== OPENAI ==========
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key: st.stop()
client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ========== OPTIONAL ERP GLOSSARY ==========
glossary_text = ""
def load_glossary(df):
    df.columns = [c.strip().lower() for c in df.columns]
    g = next((c for c in df.columns if "greek" in c or "ελλην" in c), None)
    e = next((c for c in df.columns if "approved" in c or "english" in c), None)
    if g and e:
        return "\n".join([f"{row[g]} → {row[e]}" for _, row in df.iterrows()])
    return ""
upl = st.file_uploader("📘 (Optional) Upload ERP glossary CSV", type=["csv"])
if upl:
    glossary_df = pd.read_csv(upl)
    glossary_text = load_glossary(glossary_df)
elif os.path.exists("erp_glossary.csv"):
    glossary_df = pd.read_csv("erp_glossary.csv")
    glossary_text = load_glossary(glossary_df)
else:
    glossary_text = "(no glossary provided)"

# ========== SOURCE EXCEL ==========
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

# ========== HELPERS ==========
def parse_ai_output(text):
    rows = []
    for ln in text.strip().splitlines():
        p = [x.strip() for x in ln.split("|")]
        if len(p) >= 9:
            rows.append(dict(
                Report_Name=p[0], Report_Description=p[1], Field_Name=p[2],
                Greek=p[3], English=p[4], Corrected_English=p[5],
                Status=p[6], Status_Description=p[7], Score=p[8]
            ))
    return rows

def extract_num(s):
    try:
        n = "".join(ch for ch in str(s) if ch.isdigit() or ch == ".")
        return float(n) if n else 0
    except: return 0.0

def quality_icon(score):
    s = extract_num(score)
    if s >= 90: return "🟢 Excellent"
    if s >= 70: return "🟡 Review"
    return "🔴 Poor"

# ========== MAIN AUDIT ==========
if st.button("🚀 Run ERP AI Audit"):
    results = []
    total = len(df)
    progress = st.progress(0)
    info = st.empty()

    # ✅ Added: Batch size input field
    BATCH_SIZE = st.number_input("Batch size (rows per GPT call)", min_value=10, max_value=200, value=50, step=10)

    for start in range(0, total, BATCH_SIZE):
        end = min(start + BATCH_SIZE, total)
        batch = df.iloc[start:end]
        lines = []
        for _, r in batch.iterrows():
            rn, rd, fn = str(r["Report_Name"]).strip(), str(r["Report_Description"]).strip(), str(r["Field_Name"]).strip()
            gr, en = str(r["Greek"]).strip(), str(r["English"]).strip()
            if not en or en.lower() == "nan": en = ""
            lines.append(f"{rn} | {rd} | {fn} | {gr} | {en}")
        joined = "\n".join(lines)

        prompt = f"""
You are a senior ERP localization consultant specialized in Entersoft ERP and accounting terminology.
Judge conceptually (not literally). Prefer ERP/accounting English: Net Value, Posting Date, Credit Note, Cost Center, Ledger Account, VAT Amount, Warehouse, etc.

Reference ERP glossary:
{glossary_text}

Statuses:
1 = Translated_Correct
2 = Translated_Not_Accurate
3 = Field_Not_Translated
4 = Field_Not_Found_On_Report_View

Scoring (0–100):
90–100 Excellent | 70–89 Good | 50–69 Fair | <50 Poor

Rules:
• If English is blank, translate Greek immediately → Corrected_English.
• Always assess translation quality based on Corrected_English (ignore old English).
• Output one per line exactly as:
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description | Score

Now analyze:
{joined}
"""
        try:
            r = client.chat.completions.create(
                model=MODEL,
                messages=[{"role":"system","content":"You are an ERP translation auditor."},
                          {"role":"user","content":prompt}],
                temperature=0)
            results.extend(parse_ai_output(r.choices[0].message.content))
        except Exception as e:
            st.warning(f"Batch {start}-{end} failed: {e}")
        progress.progress(end/total)
        info.write(f"Processed {end}/{total} rows...")

    out = pd.DataFrame(results)
    out["Score"] = out["Score"].apply(extract_num)
    out["Quality"] = out["Score"].apply(quality_icon)
    weak = out[out["Score"] < 70]

    st.session_state["audit_results"] = out
    st.success("✅ Audit completed.")
    st.dataframe(out.head(30))

    if len(weak) > 0:
        st.warning(f"⚠️ {len(weak)} weak translations found (<70). Click below to improve them.")
        if st.button("🔁 Improve Weak Translations"):
            for idx, r in weak.iterrows():
                gr = r["Greek"]
                ce = r["Corrected_English"]
                fix_prompt = f"""
You are an Entersoft ERP expert. Improve this weak translation to a precise ERP/accounting term.

Greek: {gr}
Current English: {ce}

Return only the improved ERP term.
"""
                try:
                    fx = client.chat.completions.create(
                        model=MODEL,
                        messages=[{"role":"user","content":fix_prompt}],
                        temperature=0)
                    new_term = fx.choices[0].message.content.strip()
                    out.at[idx,"Corrected_English"] = new_term
                    out.at[idx,"Score"] = 90
                    out.at[idx,"Status_Description"] += " | Auto-Improved"
                except Exception as e:
                    st.warning(f"Could not improve row {idx}: {e}")
            out["Quality"] = out["Score"].apply(quality_icon)
            st.session_state["audit_results"] = out
            st.success("✅ Weak translations improved.")
            st.dataframe(out.head(30))

# ========== EXPORT ==========
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
        ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or "")) for c in col)+2, 60)

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button("📥 Download Final Excel (Simplified)",
        data=buf, file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
