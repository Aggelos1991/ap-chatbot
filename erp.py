import pandas as pd
import streamlit as st
from openai import OpenAI
from deep_translator import GoogleTranslator
import io, os, time, re

# ============ CONFIG ============
st.set_page_config(page_title="ERP Translation Audit — Hybrid Fast Mode", layout="wide")
st.title("🧠 ERP Translation Audit — Hybrid Ultra-Fast Edition")

try:
    from dotenv import load_dotenv
    load_dotenv()
except:
    pass

# ============ API KEY ============
api_key = os.getenv("OPENAI_API_KEY") or st.secrets.get("OPENAI_API_KEY")
if not api_key:
    api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.warning("Please enter your OpenAI API key to continue.")
    st.stop()

client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# ============ BATCH SIZE ============
BATCH_SIZE = st.number_input("⚙️ GPT batch size (recommended 50-100)", value=80, min_value=10, max_value=200, step=10)

# ============ GLOSSARY ============
st.subheader("📘 Optional ERP Glossary")
glossary_text = ""
g_file = st.file_uploader("Upload ERP Glossary (CSV)", type=["csv"], key="glossary")

def load_glossary(df):
    df.columns = [c.strip().lower() for c in df.columns]
    g = next((c for c in df.columns if "greek" in c or "ελλην" in c), None)
    e = next((c for c in df.columns if "english" in c or "approved" in c), None)
    if g and e:
        return "\n".join([f"{row[g]} → {row[e]}" for _, row in df.iterrows()])
    return ""

if g_file:
    df_g = pd.read_csv(g_file)
    glossary_text = load_glossary(df_g)
    st.success(f"✅ Loaded uploaded glossary with {len(df_g)} terms.")
elif os.path.exists("erp_glossary.csv"):
    df_g = pd.read_csv("erp_glossary.csv")
    glossary_text = load_glossary(df_g)
    st.success(f"✅ Loaded local glossary with {len(df_g)} terms.")
else:
    glossary_text = "(no glossary provided)"
    st.info("No glossary provided — AI will rely on ERP/accounting knowledge.")

# ============ FILE UPLOAD ============
st.subheader("📂 Upload ERP Translation Export")
uploaded = st.file_uploader("Upload Excel/CSV", type=["xlsx", "csv"])
if not uploaded:
    st.stop()

# ============ HELPERS ============
def fast_translate(text):
    """Translate Greek quickly using GoogleTranslator; fallback to original if fails"""
    try:
        if not text or pd.isna(text): return ""
        tr = GoogleTranslator(source="auto", target="en").translate(text)
        return re.sub(r"[\r\n]+", " ", tr).strip()
    except Exception:
        return text

def gpt_audit_batch(batch_df):
    joined = "\n".join([
        f"{r.Report_Name} | {r.Report_Description} | {r.Field_Name} | {r.Greek} | {r.Corrected_English}"
        for _, r in batch_df.iterrows()
    ])
    prompt = f"""
You are a senior ERP localization auditor specialized in Entersoft ERP and accounting terminology.
Compare each Greek ↔ Corrected English pair conceptually.
Return one line per record exactly as:
Report_Name | Report_Description | Field_Name | Greek | Corrected_English | Status | Status_Description | Quality
Statuses: 1=Translated_Correct, 2=Translated_Not_Accurate, 3=Field_Not_Translated
Quality: 🟢 Excellent | 🟡 Review | 🔴 Poor
ERP Glossary:
{glossary_text}
---
{joined}
""".strip()

    try:
        resp = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "system", "content": "You are an ERP translation auditor."},
                      {"role": "user", "content": prompt}],
            temperature=0
        )
        text = resp.choices[0].message.content.strip()
        rows = []
        for line in text.splitlines():
            parts = [p.strip() for p in line.split("|")]
            if len(parts) >= 8:
                rows.append({
                    "Report_Name": parts[0],
                    "Report_Description": parts[1],
                    "Field_Name": parts[2],
                    "Greek": parts[3],
                    "Corrected_English": parts[4],
                    "Status": parts[5],
                    "Status_Description": parts[6],
                    "Quality": parts[7]
                })
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"GPT batch failed: {e}")
        return pd.DataFrame()

# ============ MAIN PROCESS ============
if st.button("🚀 Run Hybrid Audit"):
    df = pd.read_excel(uploaded) if uploaded.name.endswith(".xlsx") else pd.read_csv(uploaded)
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

    # auto-map columns
    colmap = {}
    for c in df.columns:
        if "report" in c and "name" in c: colmap["Report_Name"] = c
        elif "report" in c and "desc" in c: colmap["Report_Description"] = c
        elif "field" in c and "name" in c: colmap["Field_Name"] = c
        elif "greek" in c or "ελλην" in c: colmap["Greek"] = c
        elif "english" in c: colmap["English"] = c
    if len(colmap) < 5:
        st.error("❌ Missing columns: Report_Name, Report_Description, Field_Name, Greek, English")
        st.stop()

    # FAST PASS translation
    st.info("⚡ Running local fast translation first...")
    df["Report_Name"] = df[colmap["Report_Name"]]
    df["Report_Description"] = df[colmap["Report_Description"]]
    df["Field_Name"] = df[colmap["Field_Name"]]
    df["Greek"] = df[colmap["Greek"]]
    df["English"] = df[colmap["English"]]
    df["Corrected_English"] = df["Greek"].apply(fast_translate)

    total = len(df)
    progress = st.progress(0)
    results = []
    for start in range(0, total, BATCH_SIZE):
        end = min(start + BATCH_SIZE, total)
        batch = df.iloc[start:end].copy()
        progress.progress(end / total)
        audited = gpt_audit_batch(batch)
        results.append(audited)
        time.sleep(0.05)
    progress.empty()

    # combine
    audited_df = pd.concat(results, ignore_index=True)
    merged = df.merge(audited_df, on=["Report_Name", "Report_Description", "Field_Name", "Greek", "Corrected_English"], how="left")

    st.success("✅ Hybrid Audit Completed Successfully")
    st.dataframe(merged.head(30), use_container_width=True)

    weak = merged[merged["Quality"].str.contains("Review|Poor", na=False)]
    if not weak.empty:
        st.warning(f"⚠️ {len(weak)} weak translations found (<Excellent)")

    # download
    output = io.BytesIO()
    merged.to_excel(output, index=False)
    st.download_button(
        "💾 Download Final Excel (Hybrid Audit)",
        data=output.getvalue(),
        file_name="ERP_Translation_Audit_Hybrid.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
