import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os, time
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# =========================
# STREAMLIT CONFIG
# =========================
st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft AI Translation Audit — ERP Expert Edition")

# =========================
# OPENAI
# =========================
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()
client = OpenAI(api_key=api_key)

# =========================
# OPTIONAL ERP GLOSSARY (UPLOAD OR AUTOLOAD)
# Auto-detects column names, avoids KeyError
# =========================
glossary_text = ""

def load_glossary(df):
    df.columns = [c.strip().lower() for c in df.columns]
    greek_col = next((c for c in df.columns if "greek" in c or "ελλην" in c), None)
    eng_col   = next((c for c in df.columns if "approved" in c or "english" in c), None)
    if greek_col and eng_col:
        return "\n".join([f"{row[greek_col]} → {row[eng_col]}" for _, row in df.iterrows()])
    return ""

uploaded_glossary = st.file_uploader("📘 Upload optional ERP glossary (CSV)", type=["csv"], key="gloss_upl")

if uploaded_glossary:
    glossary_df = pd.read_csv(uploaded_glossary)
    glossary_text = load_glossary(glossary_df)
    st.success(f"✅ Loaded uploaded glossary with {len(glossary_df)} ERP terms.")
elif os.path.exists("erp_glossary.csv"):
    glossary_df = pd.read_csv("erp_glossary.csv")
    glossary_text = load_glossary(glossary_df)
    st.success(f"✅ Loaded local glossary with {len(glossary_df)} ERP terms.")
else:
    glossary_df = pd.DataFrame()
    st.info("No glossary provided — running with AI-only terminology intelligence.")

with st.expander("👀 Preview glossary (first 25 rows)"):
    if not glossary_df.empty:
        st.dataframe(glossary_df.head(25))
    else:
        st.write("—")

# =========================
# SOURCE EXCEL (FROM YOUR SQL EXPORT)
# Required columns: Report_Name | Report_Description | Field_Name | Greek | English
# =========================
uploaded_file = st.file_uploader("📂 Upload Excel (Report_Name | Report_Description | Field_Name | Greek | English)", type=["xlsx"], key="data_upl")
if not uploaded_file:
    st.info("Please upload your exported Excel file from SQL.")
    st.stop()

df = pd.read_excel(uploaded_file)
st.write(f"✅ File loaded successfully — {len(df)} rows detected.")

#remove this for live
if st.checkbox("Run only first 30 rows (test mode)", value=True):
    df = df.head(30)
    st.warning("⚠️ Audit limited to first 30 rows for testing.")
    #remove this for live
required_cols = {"Report_Name", "Report_Description", "Field_Name", "Greek", "English"}
if not required_cols.issubset(df.columns):
    st.error(f"❌ Excel must contain these columns: {required_cols}")
    st.stop()

# =========================
# UI PARAMS
# =========================
col_a, col_b, col_c = st.columns([1,1,1])
with col_a:
    BATCH_SIZE = st.number_input("Batch size", value=50, min_value=10, max_value=200, step=10)
with col_b:
    SHOW_PREVIEW = st.checkbox("Show interim preview (per batch)", value=False)
with col_c:
    MODEL = st.selectbox("Model", ["gpt-4o-mini", "gpt-4o"], index=0)

# =========================
# HELPERS
# =========================
def parse_ai_output(text: str):
    rows = []
    for raw in text.strip().splitlines():
        parts = [p.strip() for p in raw.split("|")]
        if len(parts) >= 9:
            rows.append({
                "Report_Name": parts[0],
                "Report_Description": parts[1],
                "Field_Name": parts[2],
                "Greek": parts[3],
                "English": parts[4],
                "Corrected_English": parts[5],
                "Status": parts[6],
                "Status_Description": parts[7],
                "Score": parts[8],
                "Retranslated": ""
            })
    return rows

def extract_score_number(s: str) -> float:
    if s is None: return 0.0
    s = str(s)
    num = ""
    dot = False
    for ch in s:
        if ch.isdigit():
            num += ch
        elif ch == "." and not dot:
            num += ch
            dot = True
        elif num:
            break
    try:
        return float(num) if num else 0.0
    except:
        return 0.0

def quality_icon(score_value):
    try: s = float(score_value)
    except: return "⚪ Unknown"
    if s >= 90: return "🟢 Excellent"
    if s >= 70: return "🟡 Review"
    return "🔴 Poor"

# =========================
# INITIAL AUDIT
# =========================
if st.button("🚀 Run ERP AI Audit"):
    results = []
    total = len(df)
    progress = st.progress(0)
    status_text = st.empty()

    for start in range(0, total, BATCH_SIZE):
        end = min(start + BATCH_SIZE, total)
        batch = df.iloc[start:end]
        prompt_rows = []

        for _, row in batch.iterrows():
            rn = str(row["Report_Name"]).strip()
            rd = str(row["Report_Description"]).strip()
            fn = str(row["Field_Name"]).strip()
            gr = str(row["Greek"]).strip()
            en = str(row["English"]).strip()
            if not en or en.lower() == "nan": en = ""
            prompt_rows.append(f"{rn} | {rd} | {fn} | {gr} | {en}")

        joined = "\n".join(prompt_rows)
        main_prompt = f"""
You are a senior ERP localization consultant specialized in Entersoft ERP and accounting terminology.
Judge conceptually (not literally). Prefer proper ERP/accounting English: Net Value, Posting Date, Credit Note, Cost Center, Ledger Account, VAT Amount, Warehouse, etc.

Reference ERP glossary (authoritative pairs):
{glossary_text or '(no glossary provided)'}

Statuses:
1 = Translated_Correct
2 = Translated_Not_Accurate
3 = Field_Not_Translated
4 = Field_Not_Found_On_Report_View

Scoring (0–100):
90–100 Excellent | 70–89 Good | 50–69 Fair | <50 Poor

Rules:
• If English is blank, translate Greek → put translation ONLY in Corrected_English.
• Output exactly as:
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description | Score

Now analyze:
{joined}
""".strip()

        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": "You are an ERP translation auditor."},
                    {"role": "user", "content": main_prompt}
                ],
                temperature=0
            )
            text = resp.choices[0].message.content
            batch_rows = parse_ai_output(text)
            results.extend(batch_rows)

            if SHOW_PREVIEW:
                st.dataframe(pd.DataFrame(batch_rows).head(5))
            progress.progress(end / total)
            status_text.write(f"Processed {end}/{total} rows...")
            time.sleep(0.2)
        except Exception as e:
            st.warning(f"⚠️ Batch {start}-{end} failed: {e}")

    out = pd.DataFrame(results)
    out["Score"] = out["Score"].apply(lambda x: f"{extract_score_number(x):.0f}")
    out["Quality"] = out["Score"].apply(quality_icon)
    st.session_state["audit_results"] = out

    st.success("✅ Audit completed.")
    st.dataframe(out.head(30))

# =========================
# RE-EVALUATION
# =========================
if "audit_results" in st.session_state and st.button("🔁 Re-Evaluate Low-Score Rows (<70)"):
    out = st.session_state["audit_results"].copy()
    low_idx = [i for i, r in out.iterrows() if extract_score_number(r["Score"]) < 70]

    st.info(f"Re-evaluating {len(low_idx)} rows...")
    for i in low_idx:
        greek = str(out.at[i, "Greek"])
        corr = str(out.at[i, "Corrected_English"])
        re_prompt = f"""
Re-evaluate ERP translation quality.
Greek: {greek}
Corrected English: {corr}
Return ONLY a number from 0 to 100.
""".strip()
        try:
            fix = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": re_prompt}],
                temperature=0
            )
            new_score = extract_score_number(fix.choices[0].message.content)
            out.at[i, "Score"] = f"{new_score:.0f}"
            out.at[i, "Retranslated"] = "🔁 Re-evaluated"
            desc = str(out.at[i, "Status_Description"])
            if "Re-evaluated" not in desc:
                out.at[i, "Status_Description"] = (desc + " | Re-evaluated").strip(" |")
        except Exception as e:
            st.warning(f"Could not re-evaluate row {i}: {e}")

    out["Quality"] = out["Score"].apply(quality_icon)
    st.session_state["audit_results"] = out
    st.success("✅ Re-evaluation complete.")
    st.dataframe(out.head(30))

# =========================
# EXPORT
# =========================
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
        ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or "")) for c in col) + 2, 60)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.download_button(
        "📥 Download Final Excel (with Quality Icons)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    try:
        num = pd.to_numeric(out["Score"], errors="coerce")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Avg Score", f"{num.mean():.1f}")
        c2.metric("🟢 Excellent", (num >= 90).sum())
        c3.metric("🟡 Review", ((num >= 70) & (num < 90)).sum())
        c4.metric("🔴 Poor", (num < 70).sum())
    except:
        pass
