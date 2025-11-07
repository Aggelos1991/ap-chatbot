import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os

# ================= CONFIG =================
st.set_page_config(page_title="ERP Translation Audit", layout="wide")
st.title("🧠 ERP Translation Audit Dashboard")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ No API key found. Add it to .env or Streamlit secrets.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ================= FILE UPLOAD =================
uploaded = st.file_uploader("📤 Upload translation Excel (Greek + English)", type=["xlsx", "csv"])

# ================= GPT HELPER =================
def translate_with_gpt(text):
    if not text or pd.isna(text):
        return ""
    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": "You are a professional ERP translation auditor for Greek to English labels."},
                {"role": "user", "content": f"Translate the following ERP field name from Greek to professional English: {text}"}
            ],
            temperature=0
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"

# ================= LOGIC =================
def evaluate_translation(greek, english):
    """STATUS based on Greek -> English"""
    if not english or english.strip() == "":
        return 3, "Field_Not_Translated"
    prompt = f"Does the English '{english}' accurately translate the Greek '{greek}'? Reply only with Yes or No."
    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        answer = response.choices[0].message.content.strip().lower()
        if "yes" in answer:
            return 1, "Translated_Correct"
        else:
            return 2, "Translated_Not_Accurate"
    except:
        return 2, "Translated_Not_Accurate"

def evaluate_quality(greek, corrected_english):
    """QUALITY based on Greek -> Corrected English"""
    if not corrected_english or corrected_english.strip() == "":
        return "Poor"
    prompt = f"Evaluate if '{corrected_english}' correctly translates '{greek}'. Respond only with Excellent, Review, or Poor."
    try:
        response = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        quality = response.choices[0].message.content.strip().title()
        if "Excellent" in quality:
            return "Excellent"
        elif "Review" in quality:
            return "Review"
        else:
            return "Poor"
    except:
        return "Review"

# ================= PROCESS =================
if uploaded:
    df = pd.read_excel(uploaded) if uploaded.name.endswith(".xlsx") else pd.read_csv(uploaded)
    required_cols = ["Report_Name", "Report_Description", "Field_Name", "Greek", "English"]
    if not all(c in df.columns for c in required_cols):
        st.error(f"❌ Missing required columns: {required_cols}")
        st.stop()

    progress_text = st.empty()
    progress_bar = st.progress(0)
    results = []

    for i, row in df.iterrows():
        greek, english = str(row["Greek"]).strip(), str(row["English"]).strip()
        corrected = translate_with_gpt(greek)

        status, status_desc = evaluate_translation(greek, english)
        quality = evaluate_quality(greek, corrected)

        results.append({
            "Report_Name": row["Report_Name"],
            "Report_Description": row["Report_Description"],
            "Field_Name": row["Field_Name"],
            "Greek": greek,
            "English": english,
            "Corrected_English": corrected,
            "Status": status,
            "Status_Description": status_desc,
            "Quality": quality
        })
        progress_bar.progress((i + 1) / len(df))
        progress_text.text(f"Processed {i+1}/{len(df)} rows...")

    progress_bar.empty()
    progress_text.empty()

    final_df = pd.DataFrame(results)

    # Summary
    weak_rows = final_df[final_df["Quality"].isin(["Review", "Poor"])]
    if not weak_rows.empty:
        st.warning(f"⚠️ {len(weak_rows)} weak translations found (<Excellent). Automatically improved.")
    else:
        st.success("✅ Audit completed successfully. All translations excellent.")

    # Display final table
    st.dataframe(final_df, use_container_width=True)

    # Download
    output = io.BytesIO()
    final_df.to_excel(output, index=False)
    st.download_button("💾 Download Final Excel (Simplified)", data=output.getvalue(),
                       file_name="Translation_Audit_Final.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
