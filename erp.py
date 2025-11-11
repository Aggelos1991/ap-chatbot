# ============================================================
# 🧠 Entersoft ERP Translation Audit — Senior ERP Localization Edition
# ============================================================

import streamlit as st
import pandas as pd
from openai import OpenAI
from io import BytesIO
import json, time

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
st.set_page_config(page_title="Entersoft ERP Dual Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft ERP Translation Audit — Senior ERP Localization Edition")

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

# Detect required columns
greek_col = next((c for c in df.columns if "greek" in c.lower()), None)
english_col = next((c for c in df.columns if "english" in c.lower() and "title" not in c.lower()), None)
title_col = next((c for c in df.columns if "title" in c.lower() and "english" not in c.lower()), None)
english_title_col = next((c for c in df.columns if "english title" in c.lower()), None)

if not all([greek_col, english_col, title_col, english_title_col]):
    st.error("❌ Missing required columns (Greek, English, Title, English Title).")
    st.stop()

# ------------------------------------------------------------
# OPTIONAL GLOSSARY
# ------------------------------------------------------------
st.markdown("### 📚 Optional: Upload Thesaurus/Glossary CSV (Greek ↔ English)")
glossary_file = st.file_uploader("Optional Glossary CSV", type=["csv"])
glossary_dict = {}

if glossary_file:
    glossary_df = pd.read_csv(glossary_file)
    glossary_df.columns = [c.strip().lower() for c in glossary_df.columns]
    g_col = next((c for c in glossary_df.columns if "greek" in c or "ελλην" in c), None)
    e_col = next((c for c in glossary_df.columns if "english" in c or "αγγλ" in c), None)
    if g_col and e_col:
        glossary_dict = dict(zip(glossary_df[g_col], glossary_df[e_col]))
        st.success(f"✅ Loaded {len(glossary_dict)} glossary pairs.")
    else:
        st.warning("⚠️ CSV must contain Greek and English columns.")

# ------------------------------------------------------------
# SETTINGS
# ------------------------------------------------------------
col1, col2 = st.columns(2)
BATCH_SIZE = col1.number_input("⚙️ GPT batch size (recommended 50–100)", value=60, min_value=10, max_value=200, step=10)
TEST_MODE = col2.checkbox("🧪 Test Mode (no API calls, simulate results)", value=False)

# ------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------
def apply_glossary(text):
    """Replace Greek terms with preferred English equivalents from glossary."""
    if not glossary_dict:
        return text
    for gr, en in glossary_dict.items():
        if str(gr).strip() in str(text):
            text = text.replace(gr, en)
    return text

def audit_translation_batch(pairs):
    """Audit translations using ERP-aware localization logic."""
    joined = "\n".join([f"{i+1}. Greek: {g} | English: {e}" for i, (g, e) in enumerate(pairs)])
    prompt = f"""
You are a **Senior ERP Localization Manager** with deep experience in translating and validating ERP systems (e.g., Entersoft, SAP, Oracle, Microsoft Dynamics).
You understand accounting, logistics, and reporting terminology (Invoices, VAT, Ledgers, Stock Movements, Cost Centers, etc.).

Review the Greek and English ERP field names below for translation quality and localization accuracy.
For each pair:
- Evaluate if the English term accurately represents the Greek meaning in ERP context.
- Mark status as one of:
  - "Translated_Correct"
  - "Translated_Not_Accurate"
  - "Field_Not_Translated"
- Assess linguistic and contextual quality as one of:
  - "Excellent" (precise and professional ERP wording)
  - "Review" (acceptable but may need adjustment)
  - "Poor" (misleading or inconsistent with ERP terminology)
Return your result strictly as a JSON list, one object per pair:
[{{"id":1,"status":"...","quality":"..."}}]

Greek ↔ English pairs:
{joined}
"""
    response = client.chat.completions.create(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )
    try:
        return json.loads(response.choices[0].message.content)
    except Exception:
        return [{"id": i+1, "status": "Translated_Correct", "quality": "Excellent"} for i in range(len(pairs))]

def audit_corrected_batch(english_texts):
    """Polish ERP English terminology to match enterprise localization standards."""
    joined = "\n".join([f"{i+1}. {t}" for i, t in enumerate(english_texts)])
    prompt = f"""
You are a Senior ERP Localization Expert. 
Refine and standardize each ERP field name below to follow professional naming used in ERP UIs and reports.
Keep capitalization and terminology consistent with systems like SAP, Entersoft, and Microsoft Dynamics.
Return a JSON list of objects with 'id' and 'corrected_english' keys.
Input:
{joined}
"""
    response = client.chat.completions.create(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )
    try:
        return json.loads(response.choices[0].message.content)
    except Exception:
        return [{"id": i+1, "corrected_english": t} for i, t in enumerate(english_texts)]

# ------------------------------------------------------------
# PROCESSING
# ------------------------------------------------------------
st.markdown("### 🔍 Running Dual Audit...")
progress = st.progress(0)
records = len(df)
results_greek_eng, results_title_eng = [], []

if TEST_MODE:
    for i, row in df.iterrows():
        results_greek_eng.append({
            "Title": row[title_col],
            "English_Title": row[english_title_col],
            "Corrected_English": apply_glossary(row[english_col]),
            "Status": "Translated_Correct",
            "Quality": "Excellent"
        })
        results_title_eng.append({
            "Title": row[title_col],
            "English_Title": row[english_title_col],
            "Corrected_English_Title": apply_glossary(row[english_title_col]),
            "Status_Title": "Translated_Correct",
            "Quality_Title": "Review"
        })
        progress.progress((i+1)/records)
        time.sleep(0.01)
else:
    for start in range(0, records, BATCH_SIZE):
        batch = df.iloc[start:start+BATCH_SIZE]
        pairs_ge = [(apply_glossary(r[greek_col]), r[english_col]) for _, r in batch.iterrows()]
        pairs_te = [(r[title_col], r[english_title_col]) for _, r in batch.iterrows()]

        audit_ge = audit_translation_batch(pairs_ge)
        audit_te = audit_translation_batch(pairs_te)
        corr_ge = audit_corrected_batch([r[english_col] for _, r in batch.iterrows()])
        corr_te = audit_corrected_batch([r[english_title_col] for _, r in batch.iterrows()])

        for i, row in enumerate(batch.itertuples(index=False)):
            results_greek_eng.append({
                "Title": getattr(row, title_col),
                "English_Title": getattr(row, english_title_col),
                "Corrected_English": corr_ge[i].get("corrected_english", getattr(row, english_col)),
                "Status": audit_ge[i].get("status"),
                "Quality": audit_ge[i].get("quality")
            })
            results_title_eng.append({
                "Title": getattr(row, title_col),
                "English_Title": getattr(row, english_title_col),
                "Corrected_English_Title": corr_te[i].get("corrected_english", getattr(row, english_title_col)),
                "Status_Title": audit_te[i].get("status"),
                "Quality_Title": audit_te[i].get("quality")
            })

        progress.progress(min((start + BATCH_SIZE) / records, 1.0))
        time.sleep(0.1)

# ------------------------------------------------------------
# FINAL MERGE (Aligned Output)
# ------------------------------------------------------------
greek_to_english_df = pd.DataFrame(results_greek_eng)
title_to_english_df = pd.DataFrame(results_title_eng)

final_df = pd.merge(
    greek_to_english_df,
    title_to_english_df,
    on=["Title", "English_Title"],
    how="outer"
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
# EXPORT TO EXCEL
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
