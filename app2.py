# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (Enhanced Build)
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
.metric-box {
    border-radius: 12px;
    padding: 1.5rem;
    margin: 0.5rem;
    text-align: center;
    color: white;
    font-weight: 600;
    box-shadow: 0 4px 8px rgba(0,0,0,0.1);
}
.green {background: #2E7D32;}
.orange {background: #FF8F00;}
.teal {background: #26A69A;}
.purple {background: #7E57C2;}
.red {background: #C62828;}
.pink {background: #AD1457;}
.dark {background: #004D40;}
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align:center;font-size:1.3rem;color:#555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b): return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "":
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    if s.count(",") == 1 and s.count(".") == 1:
        s = s.replace(".", "").replace(",", ".") if s.find(",") > s.find(".") else s.replace(",", "")
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

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "numero", "número", "document",
                    "doc", "ref", "referencia", "nº factura", "num factura"],
        "credit": ["credit", "haber", "credito", "crédito", "abono"],
        "debit": ["debit", "debe", "cargo", "importe", "valor", "amount", "total"],
        "reason": ["reason", "motivo", "concepto", "descripcion", "detalle", "descripción"],
        "date": ["date", "fecha", "fech", "data", "issue date", "posting date"]
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

# ==================== CONSOLIDATION ==========================
def consolidate_by_invoice(df: pd.DataFrame, inv_col: str) -> pd.DataFrame:
    records = []
    if inv_col not in df.columns:
        return pd.DataFrame(columns=df.columns)
    tag = "erp" if "erp" in inv_col else "ven"
    debit_col = f"debit_{tag}"
    credit_col = f"credit_{tag}"

    for inv, group in df.groupby(inv_col, dropna=False):
        if group.empty:
            continue

        # sum all related lines (credit and debit)
        total_debit = group[debit_col].apply(normalize_number).sum()
        total_credit = group[credit_col].apply(normalize_number).sum()
        net = round(total_debit - total_credit, 2)

        base = group.iloc[0].copy()
        base["__amt"] = abs(net)
        base["__type"] = "INV" if net >= 0 else "CN"
        base[debit_col] = max(net, 0.0)
        base[credit_col] = -min(net, 0.0)
        records.append(base)

    return pd.DataFrame(records).reset_index(drop=True)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor = set()

    erp_use = consolidate_by_invoice(erp_df.copy(), "invoice_erp")
    ven_use = consolidate_by_invoice(ven_df.copy(), "invoice_ven")

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip().upper()
        e_amt = round(float(e.get("__amt", 0.0)), 2)
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            if e_inv == v_inv:
                diff = abs(e_amt - v_amt)
                if diff <= 0.01:
                    status = "Perfect Match"
                elif diff < 1.0:
                    status = "Difference Match"
                else:
                    continue
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": round(diff, 2),
                    "Status": status
                })
                used_vendor.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    matched_erp = set(matched_df["ERP Invoice"]) if not matched_df.empty else set()
    matched_ven = set(matched_df["Vendor Invoice"]) if not matched_df.empty else set()
    miss_erp = erp_use[~erp_use["invoice_erp"].isin(matched_erp)].copy()
    miss_ven = ven_use[~ven_use["invoice_ven"].isin(matched_ven)].copy()
    miss_erp = miss_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    miss_ven = miss_ven.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    return matched_df, miss_erp.reset_index(drop=True), miss_ven.reset_index(drop=True)

# ==================== UI ==========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    for col in ["invoice_erp", "invoice_ven"]:
        if col in erp_df.columns:
            erp_df[col] = erp_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})
        if col in ven_df.columns:
            ven_df[col] = ven_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})

    with st.spinner("Reconciling..."):
        tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)

    st.success("✅ Reconciliation complete!")

    # ---------- TABS ----------
    tab1, tab2, tab3 = st.tabs(["📊 Summary", "🧾 Matches", "💰 Payments"])

    # --- SUMMARY TAB ---
    with tab1:
        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        diff = tier1[tier1["Status"] == "Difference Match"] if not tier1.empty else pd.DataFrame()
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        col1.markdown(f"<div class='metric-box green'>Perfect<br><h2>{len(perf)}</h2></div>", unsafe_allow_html=True)
        col2.markdown(f"<div class='metric-box orange'>Differences<br><h2>{len(diff)}</h2></div>", unsafe_allow_html=True)
        col3.markdown(f"<div class='metric-box teal'>Tier-2<br><h2>0</h2></div>", unsafe_allow_html=True)
        col4.markdown(f"<div class='metric-box purple'>Tier-3<br><h2>0</h2></div>", unsafe_allow_html=True)
        col5.markdown(f"<div class='metric-box red'>Unmatched ERP<br><h2>{len(miss_erp)}</h2></div>", unsafe_allow_html=True)
        col6.markdown(f"<div class='metric-box pink'>Unmatched Vendor<br><h2>{len(miss_ven)}</h2></div>", unsafe_allow_html=True)

    # --- MATCHES TAB ---
    with tab2:
        st.markdown("### Tier-1 Matches")
        if not tier1.empty:
            st.dataframe(tier1, use_container_width=True)
        else:
            st.info("No matches found.")

        st.markdown("### Missing in ERP")
        st.dataframe(miss_ven, use_container_width=True)
        st.markdown("### Missing in Vendor")
        st.dataframe(miss_erp, use_container_width=True)

    # --- PAYMENTS TAB (placeholder) ---
    with tab3:
        st.markdown("<div class='metric-box dark'>Matched Payments<br><h2>0</h2></div>", unsafe_allow_html=True)
