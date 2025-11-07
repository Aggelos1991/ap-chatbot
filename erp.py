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
st.title("🧠 Entersoft ERP Translation Audit — Optimized Batch Edition")

# ==========================================================
# OPENAI
# ==========================================================
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()
client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

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
# MAIN AUDIT (BATCHED)
# ==========================================================
if st.button("🚀 Run Full Auto Audit (Optimized)"):
    results = []
    total = len(df)

    # ✅ Added: dynamic batch size input
    BATCH_SIZE = st.number_input("Batch size (rows per GPT call)", min_value=10, max_value=200, value=50, step=10)

    progress = st.progress(0)
    info = st.empty()

    for start in range(0, total, BATCH_SIZE):
        end = min(start + BATCH_SIZE, total)
        batch = df.iloc[start:end]

        # --- build batch text ---
        batch_lines = []
        for _, r in batch.iterrows():
            rn = str(r["Report_Name"]).strip()
            rd = str(r["Report_Description"]).strip()
            fn = str(r["Field_Name"]).strip()
            gr = str(r["Greek"]).strip()
            en = str(r["English"]).strip()
            if not en or en.lower() == "nan":
                en = ""
            batch_lines.append(f"{rn} | {rd} | {fn} | {gr} | {en}")

        joined = "\n".join(batch_lines)

        # --- one GPT call for the whole batch ---
        prompt = f"""
You are an Entersoft ERP translation auditor.
For each line below (Report_Name | Report_Description | Field_Name | Greek | English):

1️⃣ If English is blank, translate Greek to ERP English.
2️⃣ Assess if translation is correct, not accurate, or missing.
3️⃣ If weak (<70), improve it automatically.

Output one line per input EXACTLY as:
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description | Score
---
{joined}
"""

        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": "You are an ERP translation auditor."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0
            )
            text = resp.choices[0].message.content
            for ln in text.strip().splitlines():
                p = [x.strip() for x in ln.split("|")]
                if len(p) >= 9:
                    results.append({
                        "Report_Name": p[0],
                        "Report_Description": p[1],
                        "Field_Name": p[2],
                        "Greek": p[3],
                        "English": p[4],
                        "Corrected_English": p[5],
                        "Status": p[6],
                        "Status_Description": p[7],
                        "Score": p[8]
                    })
        except Exception as e:
            st.warning(f"Batch {start}-{end} failed: {e}")

        progress.progress(end / total)
        info.write(f"Processed {end}/{total} rows...")

    # === Final processing ===
    out = pd.DataFrame(results)

    def extract_num(s):
        try:
            return float("".join(ch for ch in str(s) if ch.isdigit() or ch == "."))
        except:
            return 0.0

    out["Score"] = out["Score"].apply(extract_num)

    def quality_icon(score):
        if score >= 90:
            return "🟢 Excellent"
        if score >= 70:
            return "🟡 Review"
        return "🔴 Poor"

    out["Quality"] = out["Score"].apply(quality_icon)

    st.session_state["audit_results"] = out
    st.success("✅ Optimized audit complete.")
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
        "📥 Download Final Excel (Optimized)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    num = pd.to_numeric(out["Score"], errors="coerce")
    c1, c2, c3 = st.columns(3)
    c1.metric("🟢 Excellent", (num >= 90).sum())
    c2.metric("🟡 Review", ((num >= 70) & (num < 90)).sum())
    c3.metric("🔴 Poor", (num < 70).sum())
