# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (FULL FINAL • Cobro fix • Tier de-dup • Balance Summary)
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
.balance-match   {background:#1565C0;color:#fff;font-weight:bold;}
</style>
""",
    unsafe_allow_html=True,
)

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
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v:
        return ""
    s = str(v).lower().strip()
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice", "factura", "fact", "document", "ref", "num"],
        "credit": ["credit", "haber", "abono"],
        "debit": ["debit", "debe", "importe", "amount", "valor", "total"],
        "reason": ["reason", "motivo", "concepto", "descripcion", "detalle"],
        "date": ["date", "fecha", "posting", "data"]
    }
    rename_map = {}
    for c in df.columns:
        low = str(c).lower()
        for k, aliases in mapping.items():
            if any(a in low for a in aliases):
                rename_map[c] = f"{k}_{tag}"
    out = df.rename(columns=rename_map)
    for c in ["debit", "credit"]:
        if f"{c}_{tag}" not in out.columns:
            out[f"{c}_{tag}"] = 0.0
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

def style(df, css): 
    return df.style.apply(lambda _: [css]*len(_), axis=1)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df, ven_df):
    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1)

    matched = []
    for _, e in erp_df.iterrows():
        for _, v in ven_df.iterrows():
            if str(e.get("invoice_erp", "")).strip() == str(v.get("invoice_ven", "")).strip():
                diff = abs(e["__amt"] - v["__amt"])
                matched.append({
                    "ERP Invoice": e["invoice_erp"],
                    "Vendor Invoice": v["invoice_ven"],
                    "ERP Amount": e["__amt"],
                    "Vendor Amount": v["__amt"],
                    "Difference": round(diff, 2),
                    "Status": "Perfect Match" if diff <= 0.01 else "Difference Match"
                })
                break
    matched_df = pd.DataFrame(matched)
    miss_erp = erp_df[~erp_df["invoice_erp"].isin(matched_df["ERP Invoice"] if not matched_df.empty else [])]
    miss_ven = ven_df[~ven_df["invoice_ven"].isin(matched_df["Vendor Invoice"] if not matched_df.empty else [])]
    miss_erp = miss_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    miss_ven = miss_ven.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
    return matched_df, miss_erp, miss_ven

def tier2_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty:
        return pd.DataFrame(), set(), set(), erp_miss, ven_miss
    matches, used_e, used_v = [], set(), set()
    for ei, e in erp_miss.iterrows():
        e_code, e_amt = clean_invoice_code(e["Invoice"]), round(e["Amount"], 2)
        for vi, v in ven_miss.iterrows():
            if vi in used_v: continue
            v_code, v_amt = clean_invoice_code(v["Invoice"]), round(v["Amount"], 2)
            diff, sim = abs(e_amt - v_amt), fuzzy_ratio(e_code, v_code)
            if diff <= 1 and sim >= 0.7:
                matches.append({"ERP Invoice": e["Invoice"], "Vendor Invoice": v["Invoice"], "ERP Amount": e_amt, "Vendor Amount": v_amt, "Difference": diff, "Match Type": "Tier-2"})
                used_e.add(ei); used_v.add(vi); break
    mdf = pd.DataFrame(matches)
    rem_e = erp_miss[~erp_miss.index.isin(used_e)]
    rem_v = ven_miss[~ven_miss.index.isin(used_v)]
    return mdf, used_e, used_v, rem_e, rem_v

def tier3_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty:
        return pd.DataFrame(), set(), set(), erp_miss, ven_miss
    matches, used_e, used_v = [], set(), set()
    for ei, e in erp_miss.iterrows():
        e_code, e_date, e_amt = clean_invoice_code(e["Invoice"]), e.get("Date", ""), round(e["Amount"], 2)
        for vi, v in ven_miss.iterrows():
            v_code, v_date, v_amt = clean_invoice_code(v["Invoice"]), v.get("Date", ""), round(v["Amount"], 2)
            if e_date == v_date and fuzzy_ratio(e_code, v_code) >= 0.75:
                diff = abs(e_amt - v_amt)
                matches.append({"ERP Invoice": e["Invoice"], "Vendor Invoice": v["Invoice"], "ERP Amount": e_amt, "Vendor Amount": v_amt, "Difference": diff, "Match Type": "Tier-3"})
                used_e.add(ei); used_v.add(vi); break
    mdf = pd.DataFrame(matches)
    rem_e = erp_miss[~erp_miss.index.isin(used_e)]
    rem_v = ven_miss[~ven_miss.index.isin(used_v)]
    return mdf, used_e, used_v, rem_e, rem_v

# ==================== UI ==========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        erp_df = normalize_columns(pd.read_excel(uploaded_erp, dtype=str), "erp")
        ven_df = normalize_columns(pd.read_excel(uploaded_vendor, dtype=str), "ven")

        with st.spinner("Reconciling..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)
            tier2, _, _, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            tier3, _, _, final_erp, final_ven = tier3_match(miss_erp2, miss_ven2)

        st.success("Reconciliation Complete!")

        st.markdown('<h2 class="section-title">Reconciliation Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)

        perf = tier1[tier1["Status"] == "Perfect Match"]
        diff = tier1[tier1["Status"] == "Difference Match"]

        def safe_sum(df, col): return float(df[col].sum()) if not df.empty and col in df.columns else 0.0

        with c1: st.metric("Perfect Matches", len(perf))
        with c2: st.metric("Differences", len(diff))
        with c3: st.metric("Tier-2", len(tier2))
        with c4: st.metric("Tier-3", len(tier3))
        with c5: st.metric("Unmatched ERP", len(final_erp))
        with c6: st.metric("Unmatched Vendor", len(final_ven))
        with c7: st.metric("Payments", 0)

        # ---- Balance Summary Metric ----
        with c8:
            possible_vendor_cols = ["balance", "saldo", "υπόλοιπο", "υπολοιπο", "ypolipo"]
            balance_col_erp = next((c for c in erp_df.columns if "balance" in c.lower()), None)
            balance_col_ven = next((c for c in ven_df.columns if any(p in c.lower() for p in possible_vendor_cols)), None)
            if balance_col_erp and balance_col_ven:
                def parse_amt(v):
                    s = str(v).strip().replace("€", "").replace(",", ".")
                    s = re.sub(r"[^\d.\-]", "", s)
                    try: return float(s)
                    except: return 0.0
                erp_vals = [parse_amt(v) for v in erp_df[balance_col_erp] if str(v).strip()]
                ven_vals = [parse_amt(v) for v in ven_df[balance_col_ven] if str(v).strip()]
                if erp_vals and ven_vals:
                    diff_val = round(erp_vals[-1] - ven_vals[-1], 2)
                    st.markdown('<div class="metric-container balance-match">', unsafe_allow_html=True)
                    st.metric("Balance Summary", "")
                    st.markdown(f"**ERP:** {erp_vals[-1]:,.2f}<br>**Vendor:** {ven_vals[-1]:,.2f}<br>**Diff:** {diff_val:,.2f}", unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

        # ---------- DISPLAY ----------
        st.markdown('<h2 class="section-title">Matched & Missing</h2>', unsafe_allow_html=True)
        st.dataframe(tier1, use_container_width=True)
        st.dataframe(tier2, use_container_width=True)
        st.dataframe(tier3, use_container_width=True)
        st.dataframe(final_erp, use_container_width=True)
        st.dataframe(final_ven, use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check that your files contain columns like: **invoice**, **debit/credit**, **date**, **reason**")
