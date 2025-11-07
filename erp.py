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
# Expected CSV columns: Greek, Approved_English [, Category]
# =========================
glossary_text = ""
uploaded_glossary = st.file_uploader("📘 Upload optional ERP glossary (CSV)", type=["csv"], key="gloss_upl")

if uploaded_glossary:
    glossary_df = pd.read_csv(uploaded_glossary)
    st.success(f"✅ Loaded uploaded glossary with {len(glossary_df)} ERP terms.")
    glossary_text = "\n".join([f"{row['Greek']} → {row['Approved_English']}" for _, row in glossary_df.iterrows()])
elif os.path.exists("erp_glossary.csv"):
    glossary_df = pd.read_csv("erp_glossary.csv")
    st.success(f"✅ Loaded local glossary with {len(glossary_df)} ERP terms.")
    glossary_text = "\n".join([f"{row['Greek']} → {row['Approved_English']}" for _, row in glossary_df.iterrows()])
else:
    st.info("No glossary provided — running with AI-only terminology intelligence.")

with st.expander("👀 Preview glossary (first 25 rows)"):
    try:
        st.dataframe(glossary_df.head(25))
    except Exception:
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
    """
    Expects lines:
    Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description | Score
    """
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

def quality_icon(score_value):
    try:
        s = float(score_value)
    except Exception:
        return "⚪ Unknown"
    if s >= 90: return "🟢 Excellent"
    if s >= 70: return "🟡 Review"
    return "🔴 Poor"

def extract_score_number(s: str) -> float:
    """Robust numeric extraction (handles '95', '95/100', 'Score: 95')."""
    if s is None:
        return 0.0
    s = str(s)
    num = ""
    dot_seen = False
    for ch in s:
        if ch.isdigit():
            num += ch
        elif ch == "." and not dot_seen:
            num += ch
            dot_seen = True
        elif num:  # stop once number started and a non-numeric appears
            break
    try:
        return float(num) if num else 0.0
    except Exception:
        return 0.0

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
            report_name = str(row["Report_Name"]).strip()
            report_desc = str(row["Report_Description"]).strip()
            field_name  = str(row["Field_Name"]).strip()
            greek       = str(row["Greek"]).strip()
            english     = str(row["English"]).strip()
            if not english or english.lower() == "nan":
                english = ""
            prompt_rows.append(f"{report_name} | {report_desc} | {field_name} | {greek} | {english}")

        joined = "\n".join(prompt_rows)

        main_prompt = f"""
You are a senior ERP localization consultant specialized in Entersoft ERP and accounting terminology.
Judge conceptually (not literally). Prefer proper ERP/accounting English: Net Value, Posting Date, Credit Note, Cost Center, Ledger Account, VAT Amount, Warehouse, etc.

Reference ERP glossary (authoritative pairs):
{glossary_text or '(no glossary provided)'}

Statuses:
1 = Translated_Correct (conceptually accurate)
2 = Translated_Not_Accurate (literal/partial/wrong ERP term)
3 = Field_Not_Translated (English missing → translate Greek to ERP English)
4 = Field_Not_Found_On_Report_View (irrelevant to captions)

Scoring (0–100):
90–100 = Excellent ERP term
70–89  = Good (minor nuance)
50–69  = Fair (literal/partial)
< 50   = Poor (misleading or wrong)

Rules:
• If English is blank, translate Greek — put translation ONLY in Corrected_English (do NOT change English column).
• Evaluate accuracy primarily against Corrected_English (the new version).
• Output one line per input EXACTLY as:
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

            # Warn if counts mismatch (helps catch model formatting issues)
            if len(batch_rows) != len(prompt_rows):
                st.warning(f"Batch {start}-{end}: expected {len(prompt_rows)} lines, got {len(batch_rows)}. Check model output formatting.")

            results.extend(batch_rows)

            if SHOW_PREVIEW:
                st.dataframe(pd.DataFrame(batch_rows).head(5))
            progress.progress(end / total)
            status_text.write(f"Processed {end}/{total} rows...")
            time.sleep(0.2)
        except Exception as e:
            st.warning(f"⚠️ Batch {start}-{end} failed: {e}")

    out = pd.DataFrame(results)
    # Normalize Score to numeric strings (0–100)
    out["Score"] = out["Score"].apply(lambda x: f"{extract_score_number(x):.0f}")
    out["Quality"] = out["Score"].apply(quality_icon)

    st.session_state["audit_results"] = out
    st.success("✅ Audit completed. You can now run manual re-evaluation if needed.")
    st.dataframe(out.head(30))

# =========================
# MANUAL RE-EVALUATION (BASED ON Corrected_English vs Greek)
# =========================
if "audit_results" in st.session_state and st.button("🔁 Re-Evaluate Low-Score Rows (<70)"):
    out = st.session_state["audit_results"].copy()
    low_idx = []

    for idx, row in out.iterrows():
        score = extract_score_number(row["Score"])
        if score < 70:
            low_idx.append(idx)

    st.info(f"Re-evaluating {len(low_idx)} rows based on Greek ↔ Corrected_English…")
    for idx in low_idx:
        greek = str(out.at[idx, "Greek"])
        corr  = str(out.at[idx, "Corrected_English"])

        re_prompt = f"""
You are an Entersoft ERP expert. Re-evaluate translation accuracy between Greek and Corrected English
— prioritize ERP/accounting correctness (e.g., Net Value, Posting Date, Credit Note, Cost Center, Ledger Account, VAT Amount).

Greek: {greek}
Corrected English: {corr}

Return ONLY a number from 0 to 100 representing accuracy/terminology quality.
""".strip()

        try:
            fix = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": re_prompt}],
                temperature=0
            )
            new_score = extract_score_number(fix.choices[0].message.content.strip())
            out.at[idx, "Score"] = f"{new_score:.0f}"
            prev_desc = str(out.at[idx, "Status_Description"])
            if "Re-evaluated" not in prev_desc:
                out.at[idx, "Status_Description"] = (prev_desc + " | Re-evaluated").strip(" |")
            out.at[idx, "Retranslated"] = "🔁 Re-evaluated"
        except Exception as e:
            st.warning(f"Could not re-evaluate row {idx}: {e}")

    out["Quality"] = out["Score"].apply(quality_icon)
    st.session_state["audit_results"] = out
    st.success("✅ Manual re-evaluation completed.")
    st.dataframe(out.head(30))

# =========================
# EXPORT
# =========================
if "audit_results" in st.session_state:
    out = st.session_state["audit_results"]
    wb = Workbook()
    ws = wb.active
    ws.title = "ERP Translation Audit"

    # Header
    ws.append(list(out.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Rows
    for _, row in out.iterrows():
        ws.append([row[col] for col in out.columns])

    # Auto widths
    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.download_button(
        "📥 Download Final Excel (with Quality Icons)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Quick KPIs
    try:
        numeric_scores = pd.to_numeric(out["Score"], errors="coerce")
        avg_score = numeric_scores.mean()
        excellent = (numeric_scores >= 90).sum()
        review    = ((numeric_scores >= 70) & (numeric_scores < 90)).sum()
        poor      = (numeric_scores < 70).sum()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Avg Score", f"{avg_score:.1f}")
        c2.metric("🟢 Excellent", excellent)
        c3.metric("🟡 Review", review)
        c4.metric("🔴 Poor", poor)
    except Exception:
        pass
