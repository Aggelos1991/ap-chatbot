import pandas as pd
import streamlit as st
from openai import OpenAI
import time
import io

# === STREAMLIT CONFIG ===
st.set_page_config(page_title="Entersoft Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft AI Translation Audit (Greek ↔ English, with Report Context)")

# === OPENAI SETUP ===
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()

client = OpenAI(api_key=api_key)

# === FILE UPLOAD ===
uploaded_file = st.file_uploader("📂 Upload Excel (Report_Name | Report_Description | Field_Name | Greek | English)", type=["xlsx"])
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
if st.button("🚀 Run AI Audit"):
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
            prompt_rows.append(f"{report_name} | {report_desc} | {field_name} | {greek} | {english}")

        joined = "\n".join(prompt_rows)

        prompt = f"""
You are an ERP translation auditor. For each line below (Report_Name | Report_Description | Field_Name | Greek | English):

1️⃣ Check if the English translation correctly reflects the Greek meaning.
2️⃣ If the translation is perfect, return:
Status = 1 | Status Description = Translated_Correct | Corrected_English = same as English
3️⃣ If inaccurate or partially wrong, return:
Status = 2 | Status Description = Translated_Not_Accurate | Corrected_English = your best English fix
4️⃣ If missing translation (English or Greek empty), return:
Status = 3 | Status Description = Field_Not_Translated
5️⃣ If field doesn’t correspond to a report caption (rare), return:
Status = 4 | Status Description = Field_Not_Found_on_Report_View

Return each result in **exactly this format (pipe-separated)**:
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description

Now analyze:
{joined}
"""

        try:
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a strict bilingual Greek–English ERP translation auditor. Output only in the exact pipe-separated format."},
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

    st.success("✅ Audit completed successfully!")
    st.download_button("📥 Download Results Excel", data=buffer, file_name="translation_audit_results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.dataframe(out.head(20))
