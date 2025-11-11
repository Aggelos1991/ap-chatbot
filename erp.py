# ============================================================
# 🧠 Entersoft ERP Translation Dual Audit — Final Aligned Edition
# ============================================================

import streamlit as st
import pandas as pd
from openai import OpenAI
from io import BytesIO
import time

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
st.set_page_config(page_title="Entersoft ERP Dual Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft ERP Translation Audit — Dual Alignment Edition")

# ------------------------------------------------------------
# OPENAI API
# ------------------------------------------------------------
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ------------------------------------------------------------
# FILE UPLOAD
# ------------------------------------------------------------
uploaded_file = st.file_uploader("📤 Upload your ERP translation Excel file", type=["xlsx", "xls"])
if not uploaded_file:
    st.stop()

df = pd.read_excel(uploaded_file)
df.columns = [c.strip() for c in df.columns]

# Detect main columns automatically
greek_col = next((c for c in df.columns if "greek" in c.lower()), None)
english_col = next((c for c in df.columns if "english" in c.lower()), None)
title_col = next((c for c in df.columns if "title" in c.lower() and "english" not in c.lower()), None)
english_title_col = next((c for c in df.columns if "english title" in c.lower()), None)

if not all([greek_col, english_col, title_col, english_title_col]):
    st.error("❌ Missing required columns (Greek, English, Title, English Title).")
    st.stop()

# ------------------------------------------------------------
# AI AUDIT FUNCTIONS
# ------------------------------------------------------------
def audit_translation(greek, english):
    prompt = f"""Compare Greek phrase and English translation. 
    Return a JSON with keys: status (Translated_Correct, Translated_Not_Accurate, Field_Not_Translated)
    and quality (Excellent, Review, Poor). 
    Greek: {greek}
    English: {english}"""
    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        text = response.choices[0].message.content
        return text
    except Exception as e:
        return f"error: {e}"

def audit_corrected(english_text):
    prompt = f"Polish and correct this English ERP field name to professional standard: {english_text}"
    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        return response.choices[0].message.content.strip()
    except:
        return english_text

# ------------------------------------------------------------
# PROCESSING
# ------------------------------------------------------------
progress = st.progress(0)
records = len(df)

greek_to_english = []
for i, row in df.iterrows():
    greek_text = row[greek_col]
    english_text = row[english_col]

    result = audit_translation(greek_text, english_text)
    corrected = audit_corrected(english_text)

    greek_to_english.append({
        "Title": row[title_col],
        "English_Title": row[english_title_col],
        "Corrected_English": corrected,
        "Status": "Translated_Correct" if "Correct" in result else "Translated_Not_Accurate",
        "Quality": "Excellent" if "Excellent" in result else "Review"
    })
    progress.progress((i + 1) / records)
    time.sleep(0.05)

greek_to_english_df = pd.DataFrame(greek_to_english)

# ------------------------------------------------------------
# SECOND AUDIT (Title ↔ English Title)
# ------------------------------------------------------------
title_to_english = []
for i, row in df.iterrows():
    title = row[title_col]
    english_title = row[english_title_col]

    result = audit_translation(title, english_title)
    corrected = audit_corrected(english_title)

    title_to_english.append({
        "Title": row[title_col],
        "English_Title": row[english_title_col],
        "Corrected_English_Title": corrected,
        "Status_Title": "Translated_Correct" if "Correct" in result else "Translated_Not_Accurate",
        "Quality_Title": "Excellent" if "Excellent" in result else "Review"
    })
    progress.progress((i + 1) / records)
    time.sleep(0.05)

title_to_english_title_df = pd.DataFrame(title_to_english)

# ------------------------------------------------------------
# FINAL MERGE (Horizontal Alignment)
# ------------------------------------------------------------
final_df = pd.merge(
    greek_to_english_df,
    title_to_english_title_df,
    on=["Title", "English_Title"],
    how="outer",
    suffixes=("", "_Title")
)

cols_order = [
    "Title",
    "English_Title",
    "Corrected_English",
    "Status",
    "Quality",
    "Corrected_English_Title",
    "Status_Title",
    "Quality_Title"
]
final_df = final_df.reindex(columns=cols_order)

# ------------------------------------------------------------
# DISPLAY RESULTS
# ------------------------------------------------------------
st.success("✅ Full dual audit complete (Greek ↔ English + Title ↔ English Title).")

st.dataframe(
    final_df.style.set_properties(
        **{
            "text-align": "center",
            "white-space": "nowrap"
        }
    ),
    use_container_width=True
)

# ------------------------------------------------------------
# EXPORT FINAL EXCEL
# ------------------------------------------------------------
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    final_df.to_excel(writer, index=False, sheet_name="Dual Audit")

st.download_button(
    label="📂 Download Final Excel (Dual Audit)",
    data=output.getvalue(),
    file_name="Dual_Audit.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
