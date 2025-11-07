import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os, time

# ================= CONFIG =================
st.set_page_config(page_title="ERP Translation Audit", layout="wide")
st.title("🧠 ERP Translation Audit Dashboard")

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

# ================= MAIN PROCESS =================
if uploaded and st.button("🚀 Run ERP Translation Audit"):
    df = pd.read_excel(uploaded) if uploaded.name.endswith(".xlsx") else pd.read_csv(uploaded)
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

    # ============= BATCH LOOP =============
    for start in range(0, total, BATCH_SIZE):
        end = min(start + BATCH_SIZE, total)
        batch = df.iloc[start:end]
        progress_text.text(f"Processing batch {start+1}-{end} of {total}...")

        # ---- Build prompt for whole batch ----
        batch_text = ""
        for _, row in batch.iterrows():
            rn = str(row[mapping["Report_Name"]]).strip()
            rd = str(row[mapping["Report_Description"]]).strip()
            fn = str(row[mapping["Field_Name"]]).strip()
            gr = str(row[mapping["Greek"]]).strip()
            en = str(row[mapping["English"]]).strip()
            if not en or en.lower() == "nan": en = ""
            batch_text += f"{rn} | {rd} | {fn} | {gr} | {en}\n"

        prompt = f"""
You are a professional ERP localization auditor specialized in Entersoft ERP and accounting terminology.
Analyze each of the following entries (Report_Name | Report_Description | Field_Name | Greek | English).

Rules:
1️⃣ If English is blank, translate the Greek into professional ERP/accounting English.
2️⃣ If English exists, evaluate it against the Greek.
3️⃣ Always evaluate accuracy based on the *Corrected English* (the new version).
4️⃣ Return exactly one line per entry with these fields, separated by "|":
Report_Name | Report_Description | Field_Name | Greek | English | Corrected_English | Status | Status_Description | Quality

Statuses:
1 = Translated_Correct
2 = Translated_Not_Accurate
3 = Field_Not_Translated

Quality levels:
🟢 Excellent — fully correct ERP term
🟡 Review — acceptable but could improve
🔴 Poor — inaccurate or missing

ERP Glossary reference:
{glossary_text}

Now analyze:
{batch_text}
""".strip()

        # ---- GPT call (single batch) ----
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": "You are an ERP translation auditor."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0
            )

            text = resp.choices[0].message.content.strip()
            for line in text.splitlines():
                parts = [x.strip() for x in line.split("|")]
                if len(parts) >= 9:
                    results.append({
                        "Report_Name": parts[0],
                        "Report_Description": parts[1],
                        "Field_Name": parts[2],
                        "Greek": parts[3],
                        "English": parts[4],
                        "Corrected_English": parts[5],
                        "Status": parts[6],
                        "Status_Description": parts[7],
                        "Quality": parts[8]
                    })

        except Exception as e:
            st.warning(f"⚠️ Batch {start}-{end} failed: {e}")

        progress_bar.progress(end / total)
        time.sleep(0.1)

    # ============= FINAL OUTPUT =============
    out = pd.DataFrame(results)
    st.success("✅ Fast Audit completed successfully.")
    st.dataframe(out.head(30), use_container_width=True)

    # Summary
    weak_rows = out[out["Quality"].str.contains("Review|Poor", na=False)]
    if not weak_rows.empty:
        st.warning(f"⚠️ {len(weak_rows)} weak translations found (<Excellent).")

    # Download
    output = io.BytesIO()
    out.to_excel(output, index=False)
    st.download_button(
        "💾 Download Final Excel (Ultra Fast)",
        data=output.getvalue(),
        file_name="ERP_Translation_Audit_Fast.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
