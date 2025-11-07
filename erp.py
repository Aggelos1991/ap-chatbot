import pandas as pd
import streamlit as st
from openai import OpenAI
import time
import io
import os

# === STREAMLIT CONFIG ===
st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft AI Translation Audit — Auto-Translate Edition")

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
results = []

def parse_ai_output(text):
    """Parses GPT output from pipe-separated lines"""
    out = []
    for line in text.strip().splitlines():
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 8:
            out.append({
                "Report_Name": parts[0],
                "Report_Description": parts[1],
                "Field_Name": parts[2],
                "Greek": parts[3],
                "English": parts[4],
                "Corrected_English": parts[5],
                "Status": parts[6],
                "Status_Description": "|".join(parts[7:])
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

            # 🧠 Auto-translate when English is missing or NaN
            if not english or english.lower() == "nan":
                english = f"[TRANSLATE] {greek}"

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

If an English field contains “[TRANSLATE] …”, translate the Greek text into correct ERP/Accounting English terminology.
Even when translating, keep the same status logic.

Return each row in exactly this format:
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description

Now analyze:
{joined}
"""

        try:
            resp = client.chat.completions.create(
                model="gpt-4o-mini",  # ✅ Cheap, fast, multilingual
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
                    "Status_Description": f"Error: {e}"
                })

    out = pd.DataFrame(results)

    # === EXPORT TO EXCEL ===
    buffer = io.BytesIO()
    out.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    st.success("✅ ERP Audit completed successfully!")
    st.download_button(
        "📥 Download Results Excel",
        data=buffer,
        file_name="erp_translation_audit_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.dataframe(out)  # show all rows
