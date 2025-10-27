# --------------------------------------------------------------
# ReconRaptor – Vendor Reconciliation (FINAL STABLE, NO ERRORS)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows, get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from difflib import SequenceMatcher

# ==================== PAGE CONFIG & CSS ======================
st.set_page_config(page_title="ReconRaptor — Vendor Reconciliation", layout="wide")
st.markdown("""
<style>
.big-title {
    font-size: 3rem !important;
    font-weight: 700;
    text-align: center;
    background: linear-gradient(90deg, #1E88E5, #42A5F5);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 1rem;
}
.section-title {
    font-size: 1.8rem !important;
    font-weight: 600;
    color: #1565C0;
    border-bottom: 2px solid #42A5F5;
    padding-bottom: 0.5rem;
    margin-top: 2rem;
}
.metric-container {
    padding: 1.2rem !important;
    border-radius: 12px !important;
    margin-bottom: 1rem;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
}
.perfect-match {background:#2E7D32;color:#fff;font-weight:bold;}
.difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
.tier2-match {background:#26A69A;color:#fff;font-weight:bold;}
.tier3-match {background:#7E57C2;color:#fff;font-weight:bold;}
.missing-erp {background:#C62828;color:#fff;font-weight:bold;}
.missing-vendor {background:#AD1457;color:#fff;font-weight:bold;}
.payment-match {background:#004D40;color:#fff;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.3rem; color: #555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "": return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    if s.count(",") == 1 and s.count(".") == 1:
        if s.find(",") > s.find("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(",") == 1:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "", s.count(".") - 1)
    try:
        return float(s)
    except:
        return 0.0

def normalize_date(v):
    if pd.isna(v) or str(v).strip() == "": return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s = str(v).strip().lower()
    parts = re.split(r"[-_]", s)
    for p in reversed(parts):
        if re.fullmatch(r"\d{1,}", p) and not re.fullmatch(r"20[0-3]\d", p):
            s = p.lstrip("0")
            break
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    return s or "0"

# ==================== FIXED normalize_columns ====================
def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "document", "code", "doc", "ref", "num", "número", "no", "παραστατικό", "τιμολόγιο"],
        "credit": ["credit", "haber", "crédito", "nota", "πίστωση", "πιστωτικό"],
        "debit": ["debit", "debe", "importe", "amount", "valor", "χρέωση", "ποσό"],
        "reason": ["reason", "motivo", "descripcion", "αιτιολογία", "περιγραφή"],
        "date": ["date", "fecha", "ημερομηνία", "issue date", "posting date"]
    }

    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    # smart invoice
    invoice_matched = False
    for col, low in cols_lower.items():
        if any(a in low for a in mapping["invoice"]):
            rename_map[col] = f"invoice_{tag}"
            invoice_matched = True
            break
    if not invoice_matched:
        df[f"invoice_{tag}"] = ""

    # others
    for key, aliases in mapping.items():
        if key == "invoice": continue
        for col, low in cols_lower.items():
            if col in rename_map: continue
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"

    out = df.rename(columns=rename_map)

    # guarantee required columns
    for col in [f"invoice_{tag}", f"debit_{tag}", f"credit_{tag}", f"date_{tag}", f"reason_{tag}"]:
        if col not in out.columns:
            out[col] = "" if "date" in col or "invoice" in col or "reason" in col else 0.0

    out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

# ====================== STYLING =========================
def style(df, css):
    return df.style.apply(lambda _: [css] * len(_), axis=1)

# ==================== MATCHING ==========================
def match_invoices(erp_df, ven_df):
    if "invoice_erp" not in erp_df or "invoice_ven" not in ven_df:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    matched = []
    used_vendor = set()
    erp_df["__amt"] = abs(pd.to_numeric(erp_df.get("debit_erp", 0), errors="coerce") - pd.to_numeric(erp_df.get("credit_erp", 0), errors="coerce"))
    ven_df["__amt"] = abs(pd.to_numeric(ven_df.get("debit_ven", 0), errors="coerce") - pd.to_numeric(ven_df.get("credit_ven", 0), errors="coerce"))

    for e_idx, e in erp_df.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        for v_idx, v in ven_df.iterrows():
            if v_idx in used_vendor: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            if e_inv == v_inv and abs(e_amt - v_amt) <= 0.05:
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": abs(e_amt - v_amt),
                    "Status": "Perfect Match"
                })
                used_vendor.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    erp_unmatched = erp_df[~erp_df["invoice_erp"].isin(matched_df["ERP Invoice"] if not matched_df.empty else [])][["invoice_erp", "__amt"]].rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    ven_unmatched = ven_df[~ven_df["invoice_ven"].isin(matched_df["Vendor Invoice"] if not matched_df.empty else [])][["invoice_ven", "__amt"]].rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    return matched_df, erp_unmatched, ven_unmatched

# ==================== EXCEL EXPORT =========================
def export_excel(t1, miss_erp, miss_ven):
    wb = Workbook()
    ws1 = wb.active; ws1.title = "Matches"
    if not t1.empty:
        for r in dataframe_to_rows(t1, index=False, header=True): ws1.append(r)
    ws2 = wb.create_sheet("Missing ERP")
    if not miss_erp.empty:
        for r in dataframe_to_rows(miss_erp, index=False, header=True): ws2.append(r)
    ws3 = wb.create_sheet("Missing Vendor")
    if not miss_ven.empty:
        for r in dataframe_to_rows(miss_ven, index=False, header=True): ws3.append(r)
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ==================== UI =========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)

    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    with st.spinner("Analyzing invoices..."):
        tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)

    st.success("Reconciliation Complete!")
    st.dataframe(tier1 if not tier1.empty else pd.DataFrame({"Info": ["No exact matches found."]}))
    st.dataframe(miss_erp if not miss_erp.empty else pd.DataFrame({"Info": ["No unmatched ERP invoices."]}))
    st.dataframe(miss_ven if not miss_ven.empty else pd.DataFrame({"Info": ["No unmatched Vendor invoices."]}))

    excel_buf = export_excel(tier1, miss_erp, miss_ven)
    st.download_button("Download Excel Report", excel_buf, "ReconRaptor_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
