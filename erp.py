import pandas as pd
import streamlit as st
from openai import OpenAI
import time

# === STREAMLIT CONFIG ===
st.set_page_config(page_title="Entersoft Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft AI Translation Audit")

# === OPENAI SETUP ===
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()

client = OpenAI(api_key=api_key)

# === FILE UPLOAD ===
uploaded_file = st.file_uploader("📂 Upload Entersoft dictionary Excel (OBJECTID | Greek | English)", type=["xlsx"])
if not uploaded_file:
    st.info("Please upload your Excel file to start.")
    st.stop()

df = pd.read_excel(uploaded_file)
st.write(f"✅ File loaded successfully — {len(df)} rows detected.")

# === PARAMETERS ===
BATCH_SIZE = st.number_input("Batch size (recommended 50–100)", value=50, step=10)
results = []

def parse_ai_output(text):
    out = []
    for line in text.strip().splitlines():
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 6:
            out.append({
                "OBJECTID": parts[0],
                "Greek": parts[1],
                "English": parts[2],
                "Status": parts[3],
                "Status Description": parts[4],
                "Reason": "|".join(parts[5:])
            })
    return out

# === PROCESS BUTTON ===
if st.button("🚀 Run AI Audit"):
    total_batches = len(df) // BATCH_SIZE + (1 if len(df) % BATCH_SIZE else 0)
    progress = st.progress(0)
    st.write("Processing... Please wait.")

    for i in range(0, len(df), BATCH_SIZE):
        batch = df.iloc[i:i+BATCH_SIZE]

        prompt_rows = []
        for _, row in batch.iterrows():
            greek = str(row["Greek"]).strip()
            english = str(row["English"]).strip()
            objid = str(row["OBJECTID"]).strip()
            if greek and english:
                prompt_rows.append(f"{objid} | {greek} | {english}")

        joined = "\n".join(prompt_rows)

        prompt = f"""
You are an ERP translation auditor. For each line below (OBJECTID | Greek | English),
decide which status applies and explain briefly.

Statuses:
1 = Translated_Correct
2 = Translated_Not Accurate
3 = Field Not Translated
4 = Field Not Found on the Report View
0 = Pending Review (unclear)

Return one line per input, formatted exactly as:
OBJECTID | Greek | English | Status | Status Description | Reason

Now analyze:
{joined}
"""

        try:
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role": "system", "content": "You are a strict ERP translation checker. Output only in pipe-separated table lines."},
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
                    "OBJECTID": row["OBJECTID"],
                    "Greek": row["Greek"],
                    "English": row["English"],
                    "Status": 0,
                    "Status Description": "Pending Review",
                    "Reason": f"Error: {e}"
                })

    out = pd.DataFrame(results)
    st.success("✅ Audit completed successfully!")
    st.download_button("📥 Download Results Excel", out.to_excel(index=False, engine="openpyxl"), file_name="translation_audit_results.xlsx")

    st.dataframe(out.head(20))
