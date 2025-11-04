# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (FINAL • Balance Metric Integrated)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
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
.perfect-match   {background:#2E7D32;color:#fff;font-weight:bold;}
.difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
.tier2-match     {background:#26A69A;color:#fff;font-weight:bold;}
.tier3-match     {background:#7E57C2;color:#fff;font-weight:bold;}
.missing-erp     {background:#C62828;color:#fff;font-weight:bold;}
.missing-vendor  {background:#AD1457;color:#fff;font-weight:bold;}
.payment-match   {background:#004D40;color:#fff;font-weight:bold;}
.balance-box     {background:#1E88E5;color:white;font-weight:bold;}
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
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(d):
        d = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v:
        return ""
    s = str(v).strip().lower()
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    return s or "0"

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "num", "numero", "document", "doc", "ref"],
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

# ==================== BALANCE METRIC ==========================
def calculate_balance_difference(erp_df, ven_df):
    balance_col_erp = next((c for c in erp_df.columns if "balance" in c.lower()), None)
    possible_vendor_cols = ["balance", "saldo", "υπόλοιπο", "ypolipo", "υπολοιπο"]
    balance_col_ven = next((c for c in ven_df.columns if any(p in c.lower() for p in possible_vendor_cols)), None)
    if not balance_col_erp or not balance_col_ven:
        return None, None, None

    def parse_amount(v):
        s = str(v).strip().replace("€", "").replace(",", ".")
        s = re.sub(r"[^\d.\-]", "", s)
        try:
            return float(s)
        except:
            return 0.0

    erp_vals = [parse_amount(v) for v in erp_df[balance_col_erp] if str(v).strip()]
    ven_vals = [parse_amount(v) for v in ven_df[balance_col_ven] if str(v).strip()]
    if not erp_vals or not ven_vals:
        return None, None, None

    return erp_vals[-1], ven_vals[-1], round(erp_vals[-1] - ven_vals[-1], 2)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df, ven_df):
    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1)
    matched = []
    for _, e in erp_df.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0.0)), 2)
        for _, v in ven_df.iterrows():
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            if e_inv == v_inv:
                diff = abs(e_amt - v_amt)
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Perfect Match" if diff <= 0.01 else "Difference Match"
                })
                break
    matched_df = pd.DataFrame(matched)
    miss_erp = erp_df[~erp_df["invoice_erp"].isin(matched_df["ERP Invoice"] if not matched_df.empty else [])]
    miss_ven = ven_df[~ven_df["invoice_ven"].isin(matched_df["Vendor Invoice"] if not matched_df.empty else [])]
    miss_erp = miss_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    miss_ven = miss_ven.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    return matched_df, miss_erp, miss_ven

# ==================== UI ==========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        erp_df = normalize_columns(pd.read_excel(uploaded_erp, dtype=str), "erp")
        ven_df = normalize_columns(pd.read_excel(uploaded_vendor, dtype=str), "ven")

        with st.spinner("Analyzing invoices..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)

        st.success("Reconciliation Complete!")

        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        diff = tier1[tier1["Status"] == "Difference Match"] if not tier1.empty else pd.DataFrame()

        st.markdown('<h2 class="section-title">Reconciliation Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Perfect Matches", len(perf))
        with c2:
            st.metric("Differences", len(diff))
        with c3:
            st.metric("Unmatched ERP", len(miss_erp) + len(miss_ven))

        # ===== BALANCE METRIC =====
        last_balance_erp, last_balance_ven, balance_diff = calculate_balance_difference(erp_df, ven_df)
        if balance_diff is not None:
            st.markdown('<div class="metric-container balance-box">', unsafe_allow_html=True)
            st.markdown(f"### 💼 Balance Difference")
            st.markdown(f"**ERP Balance:** {last_balance_erp:,.2f}  |  **Vendor Balance:** {last_balance_ven:,.2f}  |  **Difference:** {balance_diff:,.2f}")
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### ✅ Perfect Matches")
        st.dataframe(perf, use_container_width=True)
        st.markdown("### ⚠️ Differences")
        st.dataframe(diff, use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Please check that your files include columns like 'Invoice', 'Debit', 'Credit', and 'Balance'.")
