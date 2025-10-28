# --------------------------------------------------------------
# ReconRaptor – Vendor Reconciliation (FINAL, Clean, Consolidated, Simplified Excel + Metrics)
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
        padding: 1rem !important;
        border-radius: 12px !important;
        margin-bottom: 0.5rem;
        box-shadow: 0 3px 5px rgba(0,0,0,0.1);
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

    # classify doc types
    def doc_type(row, tag):
        r = str(row.get(f"reason_{tag}", "")).lower()
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        pay_pat = [
            r"^πληρωμ", r"^payment", r"^remittance", r"^bank\s*transfer",
            r"^trf", r"^remesa", r"^pago", r"^paid"
        ]
        if any(re.search(p, r) for p in pay_pat):
            return "IGNORE"
        if any(k in r for k in ["credit", "nota", "abono", "cn", "πιστωτικό"]):
            return "CN"
        return "INV"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)

    erp_df["__amt"] = erp_df.apply(
        lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1
    )
    ven_df["__amt"] = ven_df.apply(
        lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1
    )

    erp_use = erp_df[erp_df["__type"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__type"] != "IGNORE"].copy()

    # Consolidate corrections per invoice
    def consolidate_by_invoice(df, inv_col):
        records = []
        if inv_col not in df.columns:
            return pd.DataFrame(columns=df.columns)
        for inv, group in df.groupby(inv_col, dropna=False):
            if group.empty:
                continue
            sum_inv = group.loc[group["__type"] == "INV", "__amt"].sum()
            sum_cn = group.loc[group["__type"] == "CN", "__amt"].sum()
            net = round(sum_inv - sum_cn, 2)
            base = group.iloc[0].copy()
            base["__amt"] = abs(net)
            base["__type"] = "INV" if net >= 0 else "CN"
            records.append(base)
        return pd.DataFrame(records).reset_index(drop=True)

    erp_use = consolidate_by_invoice(erp_use, "invoice_erp")
    ven_use = consolidate_by_invoice(ven_use, "invoice_ven")

    # ------- Tier-1: exact invoice string & no amount threshold for difference -------
    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e.get("__amt", 0.0)), 2)
        e_typ = e.get("__type", "INV")
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v.get("__amt", 0.0)), 2)
            v_typ = v.get("__type", "INV")
            if e_typ != v_typ or e_inv != v_inv:
                continue
            diff = abs(e_amt - v_amt)
            status = "Perfect Match" if diff <= 0.01 else "Difference Match"
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

    miss_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven)].copy()
    miss_ven = ven_use[~ven_use["invoice_ven"].isin(matched_erp)].copy()
    miss_erp = miss_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    miss_ven = miss_ven.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
    keep_cols = ["Invoice", "Amount", "Date"]
    miss_erp = miss_erp[[c for c in keep_cols if c in miss_erp.columns]].reset_index(drop=True)
    miss_ven = miss_ven[[c for c in keep_cols if c in miss_ven.columns]].reset_index(drop=True)

    return matched_df, miss_erp, miss_ven

# ==================== EXCEL EXPORT (ONLY UNMATCHED) =========================
def export_excel(miss_erp, miss_ven):
    wb = Workbook()
    ws = wb.active
    ws.title = "Unmatched"

    cur = 1
    if not miss_ven.empty:
        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=max(3, miss_ven.shape[1]))
        ws.cell(cur, 1, "Missing in ERP").font = Font(bold=True, size=13)
        cur += 2
        for r in dataframe_to_rows(miss_ven, index=False, header=True):
            ws.append(r)
        for c in ws[cur]:
            c.fill = PatternFill(start_color="C62828", end_color="C62828", fill_type="solid")
            c.font = Font(color="FFFFFF", bold=True)
        cur = ws.max_row + 3
    if not miss_erp.empty:
        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=max(3, miss_erp.shape[1]))
        ws.cell(cur, 1, "Missing in Vendor").font = Font(bold=True, size=13)
        cur += 2
        for r in dataframe_to_rows(miss_erp, index=False, header=True):
            ws.append(r)
        for c in ws[cur]:
            c.fill = PatternFill(start_color="AD1457", end_color="AD1457", fill_type="solid")
            c.font = Font(color="FFFFFF", bold=True)

    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

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

        with st.spinner("Analyzing invoices..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)

        st.success("Reconciliation Complete!")

        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        diff = tier1[tier1["Status"] == "Difference Match"] if not tier1.empty else pd.DataFrame()

        # ---------- METRICS ----------
        st.markdown('<h2 class="section-title">Reconciliation Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("Perfect Matches", len(perf))
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="metric-container difference-match">', unsafe_allow_html=True)
            st.metric("Differences", len(diff))
            st.markdown('</div>', unsafe_allow_html=True)
        with c3:
            st.markdown('<div class="metric-container missing-erp">', unsafe_allow_html=True)
            st.metric("Unmatched ERP", len(miss_erp))
            st.markdown('</div>', unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="metric-container missing-vendor">', unsafe_allow_html=True)
            st.metric("Unmatched Vendor", len(miss_ven))
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")

        # ---------- DISPLAY ----------
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Perfect Matches**")
            st.dataframe(perf, use_container_width=True)
        with col_b:
            st.markdown("**Differences**")
            st.dataframe(diff, use_container_width=True)

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown("**Missing in ERP**")
            st.dataframe(miss_ven, use_container_width=True)
        with col_m2:
            st.markdown("**Missing in Vendor**")
            st.dataframe(miss_erp, use_container_width=True)

        # ---------- EXPORT ----------
        st.markdown('<h2 class="section-title">Download Unmatched Report</h2>', unsafe_allow_html=True)
        excel_buf = export_excel(miss_erp, miss_ven)
        st.download_button(
            label="Download Unmatched Excel Report",
            data=excel_buf,
            file_name="ReconRaptor_Unmatched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check that your files contain columns like: **invoice**, **debit/credit**, **date**, **reason**")
