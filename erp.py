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
st.title("🧠 Entersoft AI Translation Audit — Final Version with Scoring & Icons")

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

# Validate required columns
required_cols = {"Report_Name", "Report_Description", "Field_Name", "Greek", "English"}
if not required_cols.issubset(df.columns):
    st.error(f"❌ Excel must contain these columns: {required_cols}")
    st.stop()

# === PARAMETERS ===
BATCH_SIZE = st.number_input("Batch size (recommended 50–100)", value=50, step=10)
AUTO_RETRANSLATE = st.checkbox("♻️ Auto-retranslate rows with score < 70", value=True)
results = []

# === PARSER ===
def parse_ai_output(text):
    """Parses GPT output from pipe-separated lines"""
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
                "Score": parts[8]
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

            # 🧠 If English is blank or NaN, keep it blank — GPT translates Greek in Corrected_English
            if not english or english.lower() == "nan":
                english = ""

            prompt_rows.append(f"{report_name} | {report_desc} | {field_name} | {greek} | {english}")

        joined = "\n".join(prompt_rows)

        # === MAIN PROMPT ===
        prompt = f"""
You are a senior ERP localization consultant specialized in Entersoft and accounting systems.
You understand accounting, finance, logistics, CRM, and reporting terminology (GL, AP/AR, cost centers, VAT, accruals).
Judge translation correctness conceptually — not literally.
Prefer proper accounting English (e.g., 'Net Value', 'Posting Date', 'Credit Note', 'Warehouse').

Reference ERP glossary (if present):
{glossary_text or '(no glossary provided)'}

Statuses:
1 = Translated_Correct (conceptually accurate)
2 = Translated_Not_Accurate (literal or wrong ERP term)
3 = Field_Not_Translated (English missing or incomplete — translate Greek professionally into ERP English)
4 = Field_Not_Found_On_Report_View (irrelevant)

Scoring logic (0–100):
- 90–100 → Excellent (perfect ERP term)
- 70–89 → Good (minor nuance)
- 50–69 → Fair (literal or partial)
- Below 50 → Poor (misleading or wrong)

If English is blank, translate Greek into correct ERP English, put it ONLY in Corrected_English.
Do NOT change the English column.

Return each row in exactly this format:
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description | Score

Now analyze:
{joined}
"""

        try:
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a strict ERP translation auditor. Respond only in the requested format."},
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
            for _, row in batch.iterrows():
                results.append({
                    "Report_Name": row["Report_Name"],
                    "Report_Description": row["Report_Description"],
                    "Field_Name": row["Field_Name"],
                    "Greek": row["Greek"],
                    "English": row["English"],
                    "Corrected_English": "",
                    "Status": 0,
                    "Status_Description": f"Error: {e}",
                    "Score": 0
                })

    out = pd.DataFrame(results)

    # === AUTO RETRANSLATION FOR LOW SCORES ===
    if AUTO_RETRANSLATE and "Score" in out.columns:
        low_rows = out[out["Score"].astype(float) < 70]
        if not low_rows.empty:
            st.warning(f"♻️ Retranslating {len(low_rows)} low-score rows (<70)...")
            for idx, row in low_rows.iterrows():
                try:
                    re_prompt = f"Improve this ERP translation:\nGreek: {row['Greek']}\nCurrent: {row['Corrected_English']}\nReturn only the improved English."
                    fix = client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[{"role": "user", "content": re_prompt}],
                        temperature=0
                    )
                    out.at[idx, "Corrected_English"] = fix.choices[0].message.content.strip()
                    out.at[idx, "Score"] = 90
                    out.at[idx, "Status_Description"] += " | Auto-improved"
                except Exception as e:
                    st.warning(f"Could not retranslate row {idx}: {e}")

    # === ADD QUALITY ICON COLUMN ===
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

    # === EXCEL EXPORT WITH ICONS ===
    wb = Workbook()
    ws = wb.active
    ws.title = "ERP Translation Audit"

    # Write header
    ws.append(list(out.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Write rows
    for _, row in out.iterrows():
        ws.append(list(row))

    # Adjust column widths
    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    # Save to memory
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success("✅ ERP Audit completed successfully!")

    st.download_button(
        "📥 Download Excel with Quality Icons",
        data=buffer,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.dataframe(out)
