import pandas as pd
import streamlit as st
from openai import OpenAI
import time
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# === STREAMLIT CONFIG ===
st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft AI Translation Audit — Manual Re-Evaluation Edition")

# === OPENAI SETUP ===
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()
client = OpenAI(api_key=api_key)

# === OPTIONAL GLOSSARY ===
glossary_text = ""
if os.path.exists("erp_glossary.csv"):
    glossary_df = pd.read_csv("erp_glossary.csv")
    glossary_text = "\n".join([f"{row['Greek']} → {row['Approved_English']}" for _, row in glossary_df.iterrows()])
    st.success(f"📘 Loaded ERP glossary with {len(glossary_df)} terms.")
else:
    st.info("No 'erp_glossary.csv' found — continuing without glossary.")

# === FILE UPLOAD ===
uploaded_file = st.file_uploader("📂 Upload Excel (Report_Name | Report_Description | Field_Name | Greek | English)", type=["xlsx"])
if not uploaded_file:
    st.info("Please upload your exported Excel file from SQL.")
    st.stop()

df = pd.read_excel(uploaded_file)
st.write(f"✅ File loaded successfully — {len(df)} rows detected.")

required_cols = {"Report_Name", "Report_Description", "Field_Name", "Greek", "English"}
if not required_cols.issubset(df.columns):
    st.error(f"❌ Excel must contain these columns: {required_cols}")
    st.stop()

BATCH_SIZE = st.number_input("Batch size (recommended 50–100)", value=50, step=10)
results = []

# === PARSER ===
def parse_ai_output(text):
    out = []
    for line in text.strip().splitlines():
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 9:
            out.append({
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
    return out

# === INITIAL AUDIT ===
if st.button("🚀 Run ERP AI Audit"):
    total_batches = len(df) // BATCH_SIZE + (1 if len(df) % BATCH_SIZE else 0)
    progress = st.progress(0)
    st.write("Processing translations...")

    for i in range(0, len(df), BATCH_SIZE):
        batch = df.iloc[i:i+BATCH_SIZE]
        rows = []
        for _, row in batch.iterrows():
            report_name = str(row["Report_Name"]).strip()
            report_desc = str(row["Report_Description"]).strip()
            field_name = str(row["Field_Name"]).strip()
            greek = str(row["Greek"]).strip()
            english = str(row["English"]).strip()
            if not english or english.lower() == "nan":
                english = ""
            rows.append(f"{report_name} | {report_desc} | {field_name} | {greek} | {english}")
        joined = "\n".join(rows)

        prompt = f"""
You are a senior ERP localization consultant specialized in Entersoft ERP.
Judge conceptually — not literally.
Prefer accounting English like (Net Value, Posting Date, Credit Note, Cost Center, Ledger Account, etc.).

Reference glossary:
{glossary_text or '(no glossary provided)'}

Statuses:
1 = Translated_Correct
2 = Translated_Not_Accurate
3 = Field_Not_Translated (English missing → translate Greek)
4 = Field_Not_Found_On_Report_View

Score (0–100):
90–100 = Excellent
70–89 = Good
50–69 = Fair
<50 = Poor

If English is blank, translate Greek → Corrected_English.
Do NOT modify English column.

Return:
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description | Score

{joined}
"""
        try:
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are an ERP translation auditor."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0
            )
            text = resp.choices[0].message.content
            batch_results = parse_ai_output(text)
            results.extend(batch_results)
            progress.progress(min(1.0, (i + BATCH_SIZE) / len(df)))
            time.sleep(0.3)
        except Exception as e:
            st.warning(f"⚠️ Batch {i} failed: {e}")

    out = pd.DataFrame(results)
    st.session_state["audit_results"] = out
    st.success("✅ Audit completed successfully. You can now manually re-evaluate low-score rows.")
    st.dataframe(out.head(30))

# === MANUAL RE-EVALUATION BUTTON ===
if "audit_results" in st.session_state and st.button("🔁 Re-Evaluate Low-Score Rows (<70)"):
    out = st.session_state["audit_results"]
    st.info("Re-evaluating all low-score rows based on Corrected English version...")

    for idx, row in out.iterrows():
        try:
            score = float(row["Score"])
        except:
            score = 0
        if score < 70:
            re_prompt = f"""
You are an Entersoft ERP expert.
Re-evaluate the translation accuracy between Greek and Corrected English.

Greek: {row['Greek']}
Corrected English: {row['Corrected_English']}

Assign a score 0–100 based on accuracy and terminology quality.
Return ONLY a number.
"""
            try:
                fix = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": re_prompt}],
                    temperature=0
                )
                new_score = fix.choices[0].message.content.strip()
                out.at[idx, "Score"] = new_score
                out.at[idx, "Retranslated"] = "🔁 Re-evaluated"
                out.at[idx, "Status_Description"] += " | Re-evaluated based on corrected version"
            except Exception as e:
                st.warning(f"Could not re-evaluate row {idx}: {e}")

    # === QUALITY ICONS BASED ON CORRECTED ENGLISH ===
    def quality_icon(score):
        try:
            s = float(score)
        except:
            return "⚪ Unknown"
        if s >= 90:
            return "🟢 Excellent"
        elif s >= 70:
            return "🟡 Review"
        else:
            return "🔴 Poor"

    out["Quality"] = out["Score"].apply(quality_icon)
    st.session_state["audit_results"] = out

    st.success("✅ Manual re-evaluation completed.")
    st.dataframe(out.head(30))

# === EXPORT BUTTON ===
if "audit_results" in st.session_state:
    out = st.session_state["audit_results"]
    wb = Workbook()
    ws = wb.active
    ws.title = "ERP Translation Audit"
    ws.append(list(out.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for _, row in out.iterrows():
        ws.append(list(row))

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.download_button(
        "📥 Download Final Excel (After Re-Evaluation)",
        data=buffer,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
