# --------------------------------------------------------------
# ReconRaptor – Vendor Reconciliation (FINAL, Clean, Consolidated, Unlimited Difference)
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
st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)

# ==================== TITLES =========================
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
    return s

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "nº", "num", "numero", "número", "document", "doc", "ref", "referencia"],
        "credit":  ["credit", "haber", "credito", "crédito", "abono"],
        "debit":   ["debit", "debe", "cargo", "importe", "valor", "amount", "total"],
        "reason":  ["reason", "motivo", "concepto", "descripcion", "descripción", "detalle", "περιγραφή"],
        "date":    ["date", "fecha", "fech", "data", "issue date", "posting date", "ημερομηνία"]
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

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor = set()

    # ------- Tier-1: exact invoice string & amount comparison (no difference threshold) -------
    for e_idx, e in erp_df.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("debit_erp", 0)) - float(e.get("credit_erp", 0)), 2)
        for v_idx, v in ven_df.iterrows():
            if v_idx in used_vendor:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("debit_ven", 0)) - float(v.get("credit_ven", 0)), 2)
            if e_inv != v_inv:
                continue
            diff = abs(e_amt - v_amt)

            # ✅ “Difference Match” has no upper limit now
            if diff <= 0.01:
                status = "Perfect Match"
            else:
                status = "Difference Match"

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

    miss_erp = erp_df[~erp_df["invoice_erp"].isin(matched_ven)].copy()
    miss_ven = ven_df[~ven_df["invoice_ven"].isin(matched_erp)].copy()

    miss_erp = miss_erp.rename(columns={"invoice_erp": "Invoice", "date_erp": "Date"})
    miss_ven = miss_ven.rename(columns={"invoice_ven": "Invoice", "date_ven": "Date"})
    keep_cols = ["Invoice", "Date"]
    miss_erp = miss_erp[[c for c in keep_cols if c in miss_erp.columns]].reset_index(drop=True)
    miss_ven = miss_ven[[c for c in keep_cols if c in miss_ven.columns]].reset_index(drop=True)

    return matched_df, miss_erp, miss_ven

# ==================== UI =========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)

        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")

        with st.spinner("Reconciling..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)

        st.success("Reconciliation complete!")

        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        diff = tier1[tier1["Status"] == "Difference Match"] if not tier1.empty else pd.DataFrame()

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Perfect Matches")
            st.dataframe(perf, use_container_width=True)
        with c2:
            st.subheader("Differences (no threshold)")
            st.dataframe(diff, use_container_width=True)

        st.subheader("Missing in ERP")
        st.dataframe(miss_ven, use_container_width=True)
        st.subheader("Missing in Vendor")
        st.dataframe(miss_erp, use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")
