# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (Stable Working Build)
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

def style(df, css): return df.style.apply(lambda _: [css]*len(_), axis=1)

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
        net = group[debit_col].apply(normalize_number).sum() - group[credit_col].apply(normalize_number).sum()
        net = round(net, 2)
        base = group.iloc[0].copy()
        base["__amt"] = abs(net)
        base["__type"] = "INV" if net >= 0 else "CN"
        base[debit_col] = max(net, 0.0)
        base[credit_col] = -min(net, 0.0)
        records.append(base)
    return pd.DataFrame(records).reset_index(drop=True)

# ==================== MATCHING CORE (WORKING) ==========================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor = set()

    def doc_type(row, tag):
        r = str(row.get(f"reason_{tag}", "")).lower()
        pay_pat = [r"πληρωμ", r"payment", r"remittance", r"bank", r"transfer", r"trf", r"pago"]
        if any(p in r for p in pay_pat):
            return "IGNORE"
        if any(k in r for k in ["credit", "nota", "abono", "cn", "πιστω", "πίστωση"]):
            return "CN"
        return "INV"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)
    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0))
                                                 - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0))
                                                 - normalize_number(r.get("credit_ven", 0))), axis=1)

    erp_use = consolidate_by_invoice(erp_df[erp_df["__type"] != "IGNORE"].copy(), "invoice_erp")
    ven_use = consolidate_by_invoice(ven_df[ven_df["__type"] != "IGNORE"].copy(), "invoice_ven")

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

# ==================== TIER 2 ==========================
def tier2_match(erp_miss, ven_miss):
    if erp_miss.empty or ven_miss.empty:
        return pd.DataFrame(), set(), set(), erp_miss.copy(), ven_miss.copy()
    e = erp_miss.copy(); v = ven_miss.copy(); matches, used_e, used_v = [], set(), set()
    for ei, er in e.iterrows():
        if ei in used_e: continue
        e_inv = str(er.get("Invoice", "")); e_amt = round(float(er.get("Amount", 0.0)), 2)
        for vi, vr in v.iterrows():
            if vi in used_v: continue
            v_inv = str(vr.get("Invoice", "")); v_amt = round(float(vr.get("Amount", 0.0)), 2)
            diff = abs(e_amt - v_amt); sim = fuzzy_ratio(e_inv, v_inv)
            if diff <= 1.00 and sim >= 0.85:
                matches.append({"ERP Invoice": e_inv, "Vendor Invoice": v_inv,
                                "ERP Amount": e_amt, "Vendor Amount": v_amt,
                                "Difference": round(diff, 2),
                                "Fuzzy Score": round(sim, 2),
                                "Match Type": "Tier-2"})
                used_e.add(ei); used_v.add(vi); break
    mdf = pd.DataFrame(matches)
    rem_e = e[~e.index.isin(used_e)].copy()
    rem_v = v[~v.index.isin(used_v)].copy()
    return mdf, used_e, used_v, rem_e, rem_v

# ==================== EXPORT ==========================
def export_excel(t1, t2, miss_erp, miss_ven):
    wb = Workbook()
    def hdr(ws, row, color):
        for c in ws[row]:
            c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            c.font = Font(color="FFFFFF", bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

    ws1 = wb.active; ws1.title = "Tier1"
    if not t1.empty:
        for r in dataframe_to_rows(t1, index=False, header=True): ws1.append(r)
        hdr(ws1, 1, "1E88E5")
    ws2 = wb.create_sheet("Tier2")
    if not t2.empty:
        for r in dataframe_to_rows(t2, index=False, header=True): ws2.append(r)
        hdr(ws2, 1, "26A69A")
    ws3 = wb.create_sheet("Missing")
    cur = 1
    if not miss_ven.empty:
        ws3.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=3)
        ws3.cell(cur, 1, "Missing in ERP").font = Font(bold=True, size=14); cur += 2
        for r in dataframe_to_rows(miss_ven, index=False, header=True): ws3.append(r)
        hdr(ws3, cur, "C62828"); cur = ws3.max_row + 3
    if not miss_erp.empty:
        ws3.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=3)
        ws3.cell(cur, 1, "Missing in Vendor").font = Font(bold=True, size=14); cur += 2
        for r in dataframe_to_rows(miss_erp, index=False, header=True): ws3.append(r)
        hdr(ws3, cur, "AD1457")

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3
    buf = BytesIO(); wb.save(buf); buf.seek(0); return buf

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

        for col in ["invoice_erp", "invoice_ven"]:
            if col in erp_df.columns:
                erp_df[col] = erp_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})
            if col in ven_df.columns:
                ven_df[col] = ven_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})

        with st.spinner("Analyzing..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)
            tier2, _, _, final_erp_miss, final_ven_miss = tier2_match(miss_erp, miss_ven)

        st.success("Complete!")

        # ---------- METRICS ----------
        st.markdown('<h2 class="section-title">Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        diff = tier1[tier1["Status"] == "Difference Match"] if not tier1.empty else pd.DataFrame()

        with c1: st.metric("Perfect", len(perf))
        with c2: st.metric("Differences", len(diff))
        with c3: st.metric("Tier-2", len(tier2))
        with c4: st.metric("Unmatched ERP", len(final_erp_miss))
        with c5: st.metric("Unmatched Vendor", len(final_ven_miss))

        st.markdown("---")

        # ---------- DISPLAY ----------
        st.markdown('<h2 class="section-title">Tier-1: Exact Matches</h2>', unsafe_allow_html=True)
        if not perf.empty:
            st.dataframe(perf, use_container_width=True)
        if not diff.empty:
            st.dataframe(diff, use_container_width=True)
        if perf.empty and diff.empty:
            st.info("No Tier-1 matches found.")

        st.markdown('<h2 class="section-title">Tier-2: Fuzzy Matches</h2>', unsafe_allow_html=True)
        if not tier2.empty:
            st.dataframe(tier2, use_container_width=True)
        else:
            st.info("No Tier-2 matches.")

        st.markdown('<h2 class="section-title">Missing Invoices</h2>', unsafe_allow_html=True)
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown("**Missing in ERP**")
            if not final_ven_miss.empty: st.dataframe(final_ven_miss, use_container_width=True)
            else: st.success("All vendor invoices found in ERP")
        with col_m2:
            st.markdown("**Missing in Vendor**")
            if not final_erp_miss.empty: st.dataframe(final_erp_miss, use_container_width=True)
            else: st.success("All ERP invoices found in vendor")

        # ---------- EXPORT ----------
        st.markdown('<h2 class="section-title">Download Report</h2>', unsafe_allow_html=True)
        excel_buf = export_excel(tier1, tier2, final_erp_miss, final_ven_miss)
        st.download_button("Download Excel", data=excel_buf,
                           file_name="ReconRaptor_Report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Ensure your files include invoice, debit/credit, and date columns.")
