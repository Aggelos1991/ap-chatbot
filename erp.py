import pandas as pd
import streamlit as st
from openai import OpenAI
import io, os, json, time
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# ==========================================================
# CONFIG
# ==========================================================
st.set_page_config(page_title="Entersoft ERP Translation Audit", page_icon="🧠", layout="wide")
st.title("🧠 Entersoft ERP Translation Audit — Dual Field Expert Edition")

# ==========================================================
# OPENAI
# ==========================================================
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key:
    st.stop()
client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"
BATCH_SIZE = 50   # internal, fixed — no user slider

# ==========================================================
# OPTIONAL GLOSSARY
# ==========================================================
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
else:
    glossary_text = "(no glossary provided)"

# ==========================================================
# SOURCE EXCEL
# ==========================================================
upl_file = st.file_uploader("📂 Upload Excel (must include Greek, English, Title, English Title)", type=["xlsx"])
if not upl_file:
    st.info("Please upload your exported Excel file from SQL.")
    st.stop()

df = pd.read_excel(upl_file)
st.write(f"✅ File loaded successfully — {len(df)} rows detected.")

req_cols = {"Greek", "English", "Title", "English Title"}
if not req_cols.issubset(df.columns):
    st.error(f"❌ Excel must contain columns: {req_cols}")
    st.stop()

# ==========================================================
# ERP CONTEXT
# ==========================================================
ERP_CONTEXT = """
You are a senior ERP Localization Director with 20+ years of experience in translating, mapping,
and harmonizing enterprise systems such as Entersoft, SAP, and Oracle Financials.

You understand ERP structures — accounting, finance, logistics, and inventory.
You do NOT provide literal translations — use standard ERP English terms (SAP/Oracle style).

Rules:
1️⃣ Conceptual, not literal.
2️⃣ Title Case terms (Posting Date, Cost Center, Payment Method).
3️⃣ Never invent new fields.
4️⃣ Return only the corrected ERP English term.

Examples:
Καθαρή Αξία → Net Value
Πιστωτικό Τιμολόγιο → Credit Note
Ημερομηνία Καταχώρησης → Posting Date
Αποθήκη → Warehouse
Κέντρο Κόστους → Cost Center
Προμηθευτής → Supplier
Ποσό ΦΠΑ → VAT Amount
"""

# ==========================================================
# FUNCTIONS
# ==========================================================
def classify_status(greek, english):
    g, e = (greek or "").strip(), (english or "").strip()
    if not g:
        return "Field_Not_Found_On_Report_View"
    if not e:
        return "Field_Not_Translated"
    prompt = f"""
You are an ERP translation auditor.
Compare conceptually the following Greek and English field names.
Return one label:
Translated_Correct
Translated_Not_Accurate
Field_Not_Translated
Field_Not_Found_On_Report_View

Greek: {g}
English: {e}
"""
    try:
        r = client.chat.completions.create(model=MODEL, messages=[{"role": "user", "content": prompt}], temperature=0)
        result = r.choices[0].message.content.strip()
        allowed = {"Translated_Correct","Translated_Not_Accurate","Field_Not_Translated","Field_Not_Found_On_Report_View"}
        return result if result in allowed else "Translated_Not_Accurate"
    except Exception:
        return "Translated_Not_Accurate"

def quality_label(greek, corrected):
    g, c = (greek or "").strip(), (corrected or "").strip()
    if not g or not c:
        return "🟡 Review"
    prompt = f"""
Judge conceptual translation quality for ERP/accounting context.

Greek: {g}
English: {c}

Return one:
🟢 Excellent
🟡 Review
🔴 Poor
"""
    try:
        r = client.chat.completions.create(model=MODEL, messages=[{"role":"user","content":prompt}], temperature=0)
        out = r.choices[0].message.content.strip()
        return out if out in {"🟢 Excellent","🟡 Review","🔴 Poor"} else "🟢 Excellent"
    except Exception:
        return "🟡 Review"

# ==========================================================
# CACHE INIT
# ==========================================================
CACHE_FILE = "erp_translation_cache.json"
cache = {}
if os.path.exists(CACHE_FILE):
    try:
        cache = json.load(open(CACHE_FILE, "r"))
    except:
        cache = {}

# ==========================================================
# RUN AUTOMATIC AUDIT
# ==========================================================
st.write("🚀 Starting Smart-Batch Dual Audit... please wait.")

results = []
total = len(df)
progress = st.progress(0)
info = st.empty()

for start in range(0, total, BATCH_SIZE):
    end = min(start + BATCH_SIZE, total)
    batch = df.iloc[start:end]
    lines = []

    for _, r in batch.iterrows():
        for pair in [("Greek", "English"), ("Title", "English Title")]:
            src, tgt = str(r.get(pair[0], "")).strip(), str(r.get(pair[1], "")).strip()
            if src and src not in cache:
                lines.append(f"{src} | {tgt}")

    if lines:
        prompt = f"""{ERP_CONTEXT}

Translate or refine these ERP field pairs (Greek | English or Title | English Title).
Return in format:
Greek | Corrected_English

Glossary (optional reference):
{glossary_text}

{os.linesep.join(lines)}
"""
        try:
            r = client.chat.completions.create(model=MODEL, messages=[{"role": "user", "content": prompt}], temperature=0)
            for ln in r.choices[0].message.content.strip().splitlines():
                parts = [p.strip() for p in ln.split("|")]
                if len(parts) >= 2:
                    cache[parts[0]] = parts[1]
        except Exception as e:
            st.warning(f"Batch {start}-{end} failed: {e}")

    for _, r in batch.iterrows():
        row = {
            "Report_Name": str(r.get("Report_Name", "")).strip(),
            "Report_Description": str(r.get("Report_Description", "")).strip(),
            "Field_Name": str(r.get("Field_Name", "")).strip(),
            "Greek": str(r.get("Greek", "")).strip(),
            "English": str(r.get("English", "")).strip(),
            "Title": str(r.get("Title", "")).strip(),
            "English_Title": str(r.get("English Title", "")).strip()
        }

        # Greek-English
        row["Corrected_English"] = cache.get(row["Greek"], row["English"])
        row["Status"] = classify_status(row["Greek"], row["English"])
        row["Quality"] = quality_label(row["Greek"], row["Corrected_English"])

        # Title-English Title
        row["Corrected_English_Title"] = cache.get(row["Title"], row["English_Title"])
        row["Status_Title"] = classify_status(row["Title"], row["English_Title"])
        row["Quality_Title"] = quality_label(row["Title"], row["Corrected_English_Title"])

        results.append(row)

    progress.progress(end / total)
    info.write(f"Processed {end}/{total} rows...")
    json.dump(cache, open(CACHE_FILE, "w"), ensure_ascii=False, indent=2)
    time.sleep(0.3)

out = pd.DataFrame(results)
st.success("✅ Full dual audit complete.")
st.dataframe(out.head(30))

# ==========================================================
# EXPORT
# ==========================================================
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
    ws.column_dimensions[col[0].column_letter].width = min(
        max(len(str(c.value or "")) for c in col) + 2, 60
    )
buf = io.BytesIO()
wb.save(buf)
buf.seek(0)

st.download_button(
    "📥 Download Final Excel (ERP Dual Audit)",
    data=buf,
    file_name="erp_translation_audit_final.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
