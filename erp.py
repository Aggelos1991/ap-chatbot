import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os, time

# ================= CONFIG =================
st.set_page_config(page_title="ERP Translation Audit", layout="wide")
st.title("🧠 ERP Translation Audit Dashboard — With Glossary + Batch Support")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

# ================= API KEY =================
api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.warning("Please enter your OpenAI API key to continue.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ================= BATCH SIZE =================
BATCH_SIZE = st.number_input("⚙️ Batch size (recommended 30–100)", value=50, min_value=10, max_value=200, step=10)

# ================= OPTIONAL ERP GLOSSARY =================
st.subheader("📘 Optional ERP Glossary")
glossary_text = ""
glossary_file = st.file_uploader("Upload ERP Glossary (CSV)", type=["csv"], key="glossary")

def load_glossary(df):
    """Convert ERP glossary CSV to text pairs Greek → English"""
    df.columns = [c.strip().lower() for c in df.columns]
    greek_col = next((c for c in df.columns if "greek" in c or "ελλην" in c), None)
    eng_col = next((c for c in df.columns if "english" in c or "approved" in c), None)
    if greek_col and eng_col:
        return "\n".join([f"{row[greek_col]} → {row[eng_col]}" for _, row in df.iterrows()])
    return ""

if glossary_file:
    glossary_df = pd.read_csv(glossary_file)
    glossary_text = load_glossary(glossary_df)
    st.success(f"✅ Loaded uploaded glossary with {len(glossary_df)} ERP terms.")
elif os.path.exists("erp_glossary.csv"):
    glossary_df = pd.read_csv("erp_glossary.csv")
    glossary_text = load_glossary(glossary_df)
    st.success(f"✅ Loaded local glossary with {len(glossary_df)} ERP terms.")
else:
    st.info("No glossary provided — running with AI-only terminology knowledge.")
    glossary_text = "(no glossary provided)"

# ================= UPLOAD TRANSLATION FILE =================
st.subheader("📂 Upload Translations File")
uploaded = st.file_uploader("Upload Excel or CSV containing translations", type=["xlsx", "csv"])

# ================= GPT HELPERS =================
def translate_with_gpt(text):
    if not text or pd.isna(text):
        return ""
    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": "You are an ERP translation expert specialized in accounting and Entersoft ERP terminology."},
                {"role": "user", "content": f"Translate this ERP field name from Greek to professional ERP English, using context:\n\nERP Glossary:\n{glossary_text}\n\nText: {text}"}
            ],
            temperature=0
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"

def evaluate_translation(greek, english):
    if not english or english.strip() == "":
        return 3, "Field_Not_Translated"
    prompt = f"Does the English '{english}' accurately translate the Greek '{greek}' in ERP/accounting context? Reply only with Yes or No."
    try:
        resp = client.chat.completions.create(model=MODEL, messages=[{"role":"user","content":prompt}], temperature=0)
        ans = resp.choices[0].message.content.strip().lower()
        return (1, "Translated_Correct") if "yes" in ans else (2, "Translated_Not_Accurate")
    except:
        return 2, "Translated_Not_Accurate"

def evaluate_quality(greek, corrected_english):
    if not corrected_english or corrected_english.strip() == "":
        return "Poor"
    prompt = f"Evaluate if '{corrected_english}' correctly translates '{greek}' in ERP/accounting context. Respond only with Excellent, Review, or Poor."
    try:
        resp = client.chat.completions.create(model=MODEL, messages=[{"role":"user","content":prompt}], temperature=0)
        q = resp.choices[0].message.content.strip().title()
        return "Excellent" if "Excellent" in q else "Review" if "Review" in q else "Poor"
    except:
        return "Review"

# ================= PROCESS =================
if uploaded and st.button("🚀 Run ERP Translation Audit"):
    df = pd.read_excel(uploaded) if uploaded.name.endswith(".xlsx") else pd.read_csv(uploaded)

    # Normalize column names
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    mapping = {}
    for col in df.columns:
        if "report" in col and "name" in col: mapping["Report_Name"] = col
        elif "report" in col and "desc" in col: mapping["Report_Description"] = col
        elif "field" in col and "name" in col: mapping["Field_Name"] = col
        elif "greek" in col or "ελλην" in col: mapping["Greek"] = col
        elif "english" in col: mapping["English"] = col

    if len(mapping) < 5:
        st.error("❌ Could not detect all required columns (Report_Name, Report_Description, Field_Name, Greek, English).")
        st.stop()

    total = len(df)
    progress_text = st.empty()
    progress_bar = st.progress(0)
    results = []

    for start in range(0, total, BATCH_SIZE):
        end = min(start + BATCH_SIZE, total)
        batch = df.iloc[start:end]
        progress_text.text(f"Processing batch {start+1}-{end} of {total}...")

        for _, row in batch.iterrows():
            greek = str(row[mapping["Greek"]]).strip()
            english = str(row[mapping["English"]]).strip()

            corrected = translate_with_gpt(greek)
            status, status_desc = evaluate_translation(greek, english)
            quality = evaluate_quality(greek, corrected)

            results.append({
                "Report_Name": row[mapping["Report_Name"]],
                "Report_Description": row[mapping["Report_Description"]],
                "Field_Name": row[mapping["Field_Name"]],
                "Greek": greek,
                "English": english,
                "Corrected_English": corrected,
                "Status": status,
                "Status_Description": status_desc,
                "Quality": quality
            })

        progress_bar.progress(end / total)
        time.sleep(0.1)

    progress_bar.empty()
    progress_text.empty()

    final_df = pd.DataFrame(results)
    st.success("✅ Audit completed successfully.")

    # Summary
    weak_rows = final_df[final_df["Quality"].isin(["Review", "Poor"])]
    if not weak_rows.empty:
        st.warning(f"⚠️ {len(weak_rows)} weak translations found (<Excellent).")

    st.dataframe(final_df, use_container_width=True)

    # Download
    output = io.BytesIO()
    final_df.to_excel(output, index=False)
    st.download_button(
        "💾 Download Final Excel (with Glossary + Batching)",
        data=output.getvalue(),
        file_name="ERP_Translation_Audit_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
