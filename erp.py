import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os, json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft ERP Translation Audit — Ultra Fast & Stable JSON Mode")

api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()
client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"

# Optional glossary
def load_glossary(df):
    df.columns = [c.strip().lower() for c in df.columns]
    g = next((c for c in df.columns if "greek" in c or "ελλην" in c), None)
    e = next((c for c in df.columns if "approved" in c or "english" in c), None)
    if g and e:
        return "\n".join([f"{row[g]} → {row[e]}" for _, row in df.iterrows()])
    return ""

glossary_text = ""
upl = st.file_uploader("📘 (Optional) Upload ERP glossary CSV", type=["csv"])
if upl:
    glossary_df = pd.read_csv(upl)
    glossary_text = load_glossary(glossary_df)
elif os.path.exists("erp_glossary.csv"):
    glossary_df = pd.read_csv("erp_glossary.csv")
    glossary_text = load_glossary(glossary_df)

upl_file = st.file_uploader("📂 Upload Excel (Report_Name | Report_Description | Field_Name | Greek | English)", type=["xlsx"])
if not upl_file:
    st.stop()
df = pd.read_excel(upl_file)

if st.checkbox("Run only first 30 rows (test mode)", value=True):
    df = df.head(30)

req_cols = {"Report_Name", "Report_Description", "Field_Name", "Greek", "English"}
if not req_cols.issubset(df.columns):
    st.error(f"❌ Excel must contain columns: {req_cols}")
    st.stop()

status_map = {
    "1": "Translated_Correct",
    "2": "Translated_Not_Accurate",
    "3": "Field_Not_Translated",
    "4": "Field_Not_Found_On_Report_View"
}

def get_quality_label(greek, corrected):
    try:
        qprompt = f"""
Rate the conceptual translation quality between these two fields in ERP/accounting context:
Greek: {greek}
English: {corrected}
Choose only one: "🟢 Excellent", "🟡 Review", or "🔴 Poor".
"""
        r = client.chat.completions.create(model=MODEL,
                                           messages=[{"role": "user", "content": qprompt}],
                                           temperature=0)
        return r.choices[0].message.content.strip()
    except:
        return "🟡 Review"

batch_size = st.slider("⚙️ Batch size (rows per GPT call):", 10, 200, 80, step=10)
st.caption("Larger batches = faster. Recommended: 80–100 rows per call.")

if st.button("🚀 Run Full Auto Audit"):
    results = []
    total = len(df)
    progress = st.progress(0)

    for start in range(0, total, batch_size):
        end = min(start + batch_size, total)
        batch = df.iloc[start:end]

        data = []
        for _, r in batch.iterrows():
            data.append({
                "Report_Name": str(r["Report_Name"]),
                "Report_Description": str(r["Report_Description"]),
                "Field_Name": str(r["Field_Name"]),
                "Greek": str(r["Greek"]),
                "English": str(r["English"])
            })

        prompt = f"""
You are an Entersoft ERP localization auditor and translator.

For each record in this JSON array, do the following:
1. If English is missing or wrong, translate Greek into proper ERP/accounting English (Corrected_English).
2. Determine Status:
   1=Translated_Correct
   2=Translated_Not_Accurate
   3=Field_Not_Translated
   4=Field_Not_Found_On_Report_View
3. Give a conceptual accuracy Score 0–100.

Return the full corrected data as valid JSON array with these keys:
["Report_Name","Report_Description","Field_Name","Greek","English","Corrected_English","Status","Score"]

Reference ERP glossary:
{glossary_text}

Data to analyze:
{json.dumps(data, ensure_ascii=False, indent=2)}
"""

        try:
            resp = client.chat.completions.create(model=MODEL,
                messages=[{"role":"user","content":prompt}],
                temperature=0)
            content = resp.choices[0].message.content.strip()
            parsed = json.loads(content)
            for item in parsed:
                item["Status"] = status_map.get(str(item.get("Status","")), item.get("Status",""))
                results.append(item)
        except Exception as e:
            st.warning(f"Batch {start}-{end} failed: {e}")

        progress.progress(end/total)

    out = pd.DataFrame(results)
    for col in ["Report_Name","Report_Description","Field_Name","Greek","English","Corrected_English","Status","Score"]:
        if col not in out.columns:
            out[col] = ""

    st.info("🔍 Evaluating final translation quality (Greek ↔ Corrected English)...")
    out["Quality"] = [get_quality_label(row["Greek"], row["Corrected_English"]) for _, row in out.iterrows()]

    display_cols = ["Report_Name","Report_Description","Field_Name","Greek","English","Corrected_English","Status","Quality"]
    st.session_state["audit_results"] = out
    st.success("✅ Audit completed successfully.")
    st.dataframe(out[display_cols].head(30))

if "audit_results" in st.session_state:
    out = st.session_state["audit_results"]
    wb = Workbook()
    ws = wb.active
    ws.title = "ERP Translation Audit"
    ws.append(list(out.columns))
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
    for _, r in out.iterrows():
        ws.append([r[col] for col in out.columns])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or "")) for c in col)+2,60)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.download_button("📥 Download Final Excel (All Corrected)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
