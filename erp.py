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
st.title("🧠 Entersoft ERP Translation Audit — Final ERP Expert Edition")

# ==========================================================
# OPENAI
# ==========================================================
api_key = st.text_input("🔑 Enter your OpenAI API key:", type="password")
if not api_key: st.stop()
client = OpenAI(api_key=api_key)
MODEL = "gpt-4o-mini"   # 💡 Fast + cost-efficient

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
upl_file = st.file_uploader("📂 Upload Excel (Report_Name | Report_Description | Field_Name | Greek | English)", type=["xlsx"])
if not upl_file:
    st.info("Please upload your exported Excel file from SQL.")
    st.stop()

df = pd.read_excel(upl_file)
st.write(f"✅ File loaded successfully — {len(df)} rows detected.")

if st.checkbox("Run only first 30 rows (test mode)", value=True):
    df = df.head(30)
    st.warning("⚠️ Audit limited to 30 rows for testing.")

req_cols = {"Report_Name", "Report_Description", "Field_Name", "Greek", "English"}
if not req_cols.issubset(df.columns):
    st.error(f"❌ Excel must contain columns: {req_cols}")
    st.stop()

# ==========================================================
# ERP CONTEXT PROMPT
# ==========================================================
ERP_CONTEXT = """
You are a senior ERP localization consultant specialized in Entersoft ERP and accounting terminology.
Translate or refine each Greek field into clean, professional ERP-style English used in enterprise systems.

Guidelines:
• Be conceptual, not literal.
• Use standard ERP terms: Net Value, Posting Date, Credit Note, Cost Center, Ledger Account, VAT Amount, Warehouse, Supplier, Customer, Invoice Number, Payment Method, Transaction Date.
• Use Title Case.
• Return only the corrected ERP English term — no explanations.

EXAMPLES:
Καθαρή Αξία → Net Value
Ημερομηνία Καταχώρησης → Posting Date
Πιστωτικό Τιμολόγιο → Credit Note
Κέντρο Κόστους → Cost Center
Πελάτης → Customer
Προμηθευτής → Supplier
Λογαριασμός Γενικής Λογιστικής → Ledger Account
Ποσό ΦΠΑ → VAT Amount
Αποθήκη → Warehouse
"""

# ==========================================================
# HELPER FUNCTIONS
# ==========================================================
def classify_status(greek, english):
    g, e = (greek or "").strip(), (english or "").strip()
    if not g or g.lower() == "nan":
        return "Field_Not_Found_On_Report_View"
    if not e or e.lower() == "nan" or e == "":
        return "Field_Not_Translated"

    prompt = f"""
You are an ERP translation auditor.
Compare the following Greek and English field names conceptually (ignore alphabet or spelling).
Return one exact label:
Translated_Correct
Translated_Not_Accurate
Field_Not_Translated
Field_Not_Found_On_Report_View

Greek: {g}
English: {e}
"""
    try:
        r = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        result = r.choices[0].message.content.strip()
        allowed = {"Translated_Correct","Translated_Not_Accurate","Field_Not_Translated","Field_Not_Found_On_Report_View"}
        return result if result in allowed else "Translated_Not_Accurate"
    except Exception:
        return "Translated_Not_Accurate"

def correct_erp_english(greek, english_seed):
    g, seed = (greek or "").strip(), (english_seed or "").strip()
    prompt = f"""{ERP_CONTEXT}

Greek: {g}
Existing English: {seed}

Glossary Reference:
{glossary_text}
"""
    try:
        r = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        return r.choices[0].message.content.strip()
    except Exception:
        return seed or "(translation missing)"

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
        r = client.chat.completions.create(
            model=MODEL,
            messages=[{"role":"user","content":prompt}],
            temperature=0
        )
        out = r.choices[0].message.content.strip()
        return out if out in {"🟢 Excellent","🟡 Review","🔴 Poor"} else "🟢 Excellent"
    except Exception:
        return "🟡 Review"

# ==========================================================
# BATCH SIZE + CACHE
# ==========================================================
batch_size = st.slider("⚙️ Select batch size:", 20, 200, 50, step=10)
st.caption("💡 40–60 rows = best balance of accuracy and speed")

CACHE_FILE = "erp_translation_cache.json"
cache = {}
if os.path.exists(CACHE_FILE):
    try:
        cache = json.load(open(CACHE_FILE, "r"))
    except:
        cache = {}

# ==========================================================
# MAIN AUDIT (BATCHED)
# ==========================================================
if st.button("🚀 Run Smart-Batch Audit"):
    results = []
    total = len(df)
    progress = st.progress(0)
    info = st.empty()

    for start in range(0, total, batch_size):
        end = min(start + batch_size, total)
        batch = df.iloc[start:end]

        # --- build GPT input for untranslated ---
        lines = []
        for _, r in batch.iterrows():
            gr = str(r["Greek"]).strip()
            en = str(r["English"]).strip()
            if gr in cache or not gr:
                continue
            lines.append(f"{gr} | {en}")

        if lines:
            prompt = f"""{ERP_CONTEXT}

Translate or refine the following ERP field pairs (Greek | English).
Return one line per field in the format:
Greek | Corrected_English

Glossary (optional reference):
{glossary_text}

{os.linesep.join(lines)}
"""
            try:
                r = client.chat.completions.create(
                    model=MODEL,
                    messages=[{"role":"user","content":prompt}],
                    temperature=0
                )
                raw = r.choices[0].message.content.strip().splitlines()
                for ln in raw:
                    parts = [p.strip() for p in ln.split("|")]
                    if len(parts) >= 2:
                        cache[parts[0]] = parts[1]
            except Exception as e:
                st.warning(f"Batch {start}-{end} failed: {e}")

        # --- process batch ---
        for _, r in batch.iterrows():
            rn = str(r["Report_Name"]).strip()
            rd = str(r["Report_Description"]).strip()
            fn = str(r["Field_Name"]).strip()
            gr = str(r["Greek"]).strip()
            en_orig = str(r["English"]).strip()

            corrected = cache.get(gr, en_orig)
            status = classify_status(gr, en_orig)
            quality = quality_label(gr, corrected)

            results.append(dict(
                Report_Name=rn,
                Report_Description=rd,
                Field_Name=fn,
                Greek=gr,
                English=en_orig,
                Corrected_English=corrected,
                Status=status,
                Quality=quality
            ))

        progress.progress(end / total)
        info.write(f"Processed {end}/{total} rows...")
        json.dump(cache, open(CACHE_FILE, "w"), ensure_ascii=False, indent=2)
        time.sleep(0.3)

    out = pd.DataFrame(results)
    st.session_state["audit_results"] = out
    st.success("✅ Full audit complete (ERP expert quality + conceptual status + batching + caching).")
    st.dataframe(out.head(30))

# ==========================================================
# EXPORT
# ==========================================================
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
        ws.column_dimensions[col[0].column_letter].width = min(
            max(len(str(c.value or "")) for c in col) + 2, 60
        )
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    st.download_button(
        "📥 Download Final Excel (ERP Expert Smart-Batch)",
        data=buf,
        file_name="erp_translation_audit_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
