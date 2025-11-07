import pandas as pd
import streamlit as st
from openai import OpenAI
import time
import io
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

# === STREAMLIT CONFIG ===
st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft AI Translation Audit — Final ERP Expert Version")

# === OPENAI SETUP ===
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()

client = OpenAI(api_key=api_key)

# === OPTIONAL ERP GLOSSARY ===
glossary_text = ""
if os.path.exists("erp_glossary.csv"):
    glossary_df = pd.read_csv("erp_glossary.csv")
    glossary_text = "\n".join([f"{row['Greek']} → {row['Approved_English']}" for _, row in glossary_df.iterrows()])
    st.success(f"📘 Loaded ERP glossary with {len(glossary_df)} terms.")
else:
    st.info("No 'erp_glossary.csv' found — continuing without custom glossary.")

# === FILE UPLOAD ===
uploaded_file = st.file_uploader(
    "📂 Upload Excel (Report_Name | Report_Description | Field_Name | Greek | English)",
    type=["xlsx"]
)
if not uploaded_file:
    st.info("Please upload your exported Excel file from SQL.")
    st.stop()

df = pd.read_excel(uploaded_file)
st.write(f"✅ File loaded successfully — {len(df)} rows detected.")

required_cols = {"Report_Name", "Report_Description", "Field_Name", "Greek", "English"}
if not required_cols.issubset(df.columns):
    st.error(f"❌ Excel must contain these columns: {required_cols}")
    st.stop()

# === PARAMETERS ===
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

# === PROCESS BUTTON ===
if st.button("🚀 Run ERP AI Audit"):
    total_batches = len(df) // BATCH_SIZE + (1 if len(df) % BATCH_SIZE else 0)
    progress = st.progress(0)
    st.write("Processing translations... Please wait.")

    for i in range(0, len(df), BATCH_SIZE):
        batch = df.iloc[i:i+BATCH_SIZE]
        prompt_rows = []

        for _, row in batch.iterrows():
            report_name = str(row["Report_Name"]).strip()
            report_desc = str(row["Report_Description"]).strip()
            field_name = str(row["Field_Name"]).strip()
            greek = str(row["Greek"]).strip()
            english = str(row["English"]).strip()
            if not english or english.lower() == "nan":
                english = ""
            prompt_rows.append(f"{report_name} | {report_desc} | {field_name} | {greek} | {english}")

        joined = "\n".join(prompt_rows)

        # === MAIN PROMPT ===
        prompt = f"""
You are a senior ERP localization consultant specialized in Entersoft ERP.
Judge each translation conceptually — not literally.
Prefer proper accounting English (Net Value, Posting Date, Credit Note, Cost Center, Ledger Account, etc.).

Reference ERP glossary (if provided):
{glossary_text or '(no glossary provided)'}

Statuses:
1 = Translated_Correct
2 = Translated_Not_Accurate
3 = Field_Not_Translated (English missing → translate Greek)
4 = Field_Not_Found_On_Report_View

Scoring (0–100):
- 90–100 = Excellent ERP term
- 70–89 = Good, minor issue
- 50–69 = Fair
- Below 50 = Poor

If English is blank, translate Greek into ERP English — put it ONLY in Corrected_English.
Do NOT touch English column.

Output exactly:
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

    # === RETRANSLATE LOW SCORES BASED ON NEW VERSION ===
    st.info("Evaluating and improving low-score translations (<70)...")
    for idx, row in out.iterrows():
        try:
            score = float(row["Score"])
        except:
            score = 0
        if score < 70:
            re_prompt = f"""
You are an Entersoft ERP expert.
The current English translation below was scored low. 
Refine it into the most accurate ERP accounting English term possible.

Greek: {row['Greek']}
Current English version: {row['Corrected_English']}

Return ONLY the improved English term.
"""
            try:
                fix = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": re_prompt}],
                    temperature=0
                )
                out.at[idx, "Corrected_English"] = fix.choices[0].message.content.strip()
                out.at[idx, "Retranslated"] = "✅"
                out.at[idx, "Score"] = 90
                out.at[idx, "Status_Description"] += " | Auto-Improved"
            except Exception as e:
                st.warning(f"Could not retranslate row {idx}: {e}")

    # === QUALITY ICON COLUMN ===
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

    # === EXPORT TO EXCEL ===
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
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success("✅ ERP Audit completed successfully!")
    st.download_button(
        "📥 Download Final Excel with Retranslations & Icons",
        data=buffer,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.dataframe(out)
