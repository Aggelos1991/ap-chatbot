# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (FINAL • Cobro fix • Tier de-dup • FIXED)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows, get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from difflib import SequenceMatcher
import numpy as np

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
.perfect-match   {background:#2E7D32;color:#fff;font-weight:bold;}
.difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
.tier2-match     {background:#26A69A;color:#fff;font-weight:bold;}
.tier3-match     {background:#7E57C2;color:#fff;font-weight:bold;}
.missing-erp     {background:#C62828;color:#fff;font-weight:bold;}
.missing-vendor  {background:#AD1457;color:#fff;font-weight:bold;}
.payment-match   {background:#004D40;color:#fff;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.3rem; color: #555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "":
        return 0.0
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
    if pd.isna(v) or str(v).strip() == "":
        return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    for fmt in [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%m/%d/%Y", "%m-%d-%Y",
        "%Y/%m/%d", "%Y-%m-%d",
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
        "%m/%d/%y", "%m-%d-%y",
        "%Y.%m.%d",
    ]:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d):
                return d.strftime("%Y-%m-%d")
        except:
            continue
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(d):
        d = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v:
        return ""
    s = str(v).strip().lower()
    parts = re.split(r"[-_.\s]", s)
    for p in reversed(parts):
        if re.fullmatch(r"\d{1,}", p) and not re.fullmatch(r"20[0-3]\d", p):
            s = p.lstrip("0")
            break
    s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no|apd|vs)\W*", "", s)
    s = re.sub(r"20\d{2}", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    s = re.sub(r"[^\d]", "", s)
    return s or "0"

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "numero", "document", "doc", "ref"],
        "credit":  ["credit", "haber", "credito", "abono"],
        "debit":   ["debit", "debe", "cargo", "importe", "amount", "valor", "total"],
        "reason":  ["reason", "motivo", "concepto", "descripcion", "detalle"],
        "date":    ["date", "fecha", "data", "issue date", "posting date"]
    }
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"
    out = df.rename(columns=rename_map)
    for req in ["debit", "credit"]:
        c = f"{req}_{tag}"
        if c not in out.columns:
            out[c] = 0.0
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

def style(df, css):
    return df.style.apply(lambda _: [css] * len(_), axis=1)

# --- match_invoices, tier2_match, tier3_match, extract_payments functions unchanged from your last version ---

# ==================== UI ==========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")

if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)

        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")
        st.write("🧩 ERP columns detected:", list(erp_df.columns))
        st.write("🧩 Vendor columns detected:", list(ven_df.columns))

        with st.spinner("Analyzing invoices..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)

            used_erp_inv = set(tier1["ERP Invoice"].astype(str)) if not tier1.empty else set()
            used_ven_inv = set(tier1["Vendor Invoice"].astype(str)) if not tier1.empty else set()

            if not miss_erp.empty:
                miss_erp = miss_erp[~miss_erp["Invoice"].astype(str).isin(used_erp_inv)]
            if not miss_ven.empty:
                miss_ven = miss_ven[~miss_ven["Invoice"].astype(str).isin(used_ven_inv)]

            tier2, _, _, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            if not tier2.empty:
                used_erp_inv |= set(tier2["ERP Invoice"].astype(str))
                used_ven_inv |= set(tier2["Vendor Invoice"].astype(str))
                if not miss_erp2.empty:
                    miss_erp2 = miss_erp2[~miss_erp2["Invoice"].astype(str).isin(used_erp_inv)]
                if not miss_ven2.empty:
                    miss_ven2 = miss_ven2[~miss_ven2["Invoice"].astype(str).isin(used_ven_inv)]
            else:
                miss_erp2, miss_ven2 = miss_erp, miss_ven

            tier3, _, _, final_erp_miss, final_ven_miss = tier3_match(miss_erp2, miss_ven2)
            if not tier3.empty:
                used_erp_inv |= set(tier3["ERP Invoice"].astype(str))
                used_ven_inv |= set(tier3["Vendor Invoice"].astype(str))
                if not final_erp_miss.empty:
                    final_erp_miss = final_erp_miss[~final_erp_miss["Invoice"].astype(str).isin(used_erp_inv)]
                if not final_ven_miss.empty:
                    final_ven_miss = final_ven_miss[~final_ven_miss["Invoice"].astype(str).isin(used_ven_inv)]

            # ✅ FIXED: Dedented cleanup
            all_matched_erp = set()
            all_matched_vendor = set()

            if not tier1.empty:
                if "ERP Invoice" in tier1.columns:
                    all_matched_erp |= set(tier1["ERP Invoice"].astype(str))
                if "Vendor Invoice" in tier1.columns:
                    all_matched_vendor |= set(tier1["Vendor Invoice"].astype(str))

            if not tier2.empty:
                if "ERP Invoice" in tier2.columns:
                    all_matched_erp |= set(tier2["ERP Invoice"].astype(str))
                if "Vendor Invoice" in tier2.columns:
                    all_matched_vendor |= set(tier2["Vendor Invoice"].astype(str))

            if not tier3.empty:
                if "ERP Invoice" in tier3.columns:
                    all_matched_erp |= set(tier3["ERP Invoice"].astype(str))
                if "Vendor Invoice" in tier3.columns:
                    all_matched_vendor |= set(tier3["Vendor Invoice"].astype(str))

            if not final_erp_miss.empty:
                final_erp_miss = final_erp_miss[~final_erp_miss["Invoice"].astype(str).isin(all_matched_erp)]
            if not final_ven_miss.empty:
                final_ven_miss = final_ven_miss[~final_ven_miss["Invoice"].astype(str).isin(all_matched_vendor)]

            erp_pay, ven_pay, pay_match = extract_payments(erp_df, ven_df)

        st.success("Reconciliation Complete!")

        # --- rest of display + metrics exactly same as yours (unchanged) ---

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check that your files contain columns like: invoice, debit/credit, date, reason")
