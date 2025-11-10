# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (FINAL STABLE AGGREGATION BUILD)
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
import numpy as np

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
        "invoice": ["invoice", "invoice number", "inv no", "factura", "fact", "nº", "num", "numero", "document", "doc", "ref", "αρ", "παραστ"],
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

# ==================== DUPLICATE AGGREGATION ==========================
def aggregate_duplicates(df, tag):
    """Combine duplicates and offset CNs"""
    if df is None or df.empty:
        st.info(f"ℹ️ Skipping {tag.upper()} aggregation — empty file.")
        return df

    inv_col = f"invoice_{tag}"
    if inv_col not in df.columns:
        st.warning(f"⚠️ '{inv_col}' not found in {tag.upper()} file. Columns: {', '.join(df.columns)}")
        return df

    df = df.copy()
    for c in [f"debit_{tag}", f"credit_{tag}"]:
        if c not in df.columns:
            df[c] = 0.0

    df["__net"] = df.apply(lambda r: normalize_number(r.get(f"debit_{tag}", 0)) - normalize_number(r.get(f"credit_{tag}", 0)), axis=1)

    def norm_code(v):
        if pd.isna(v) or str(v).strip() == "":
            return None
        s = re.sub(r"[^A-Z0-9]", "", str(v).upper())
        s = s.replace("INV", "").replace("FACT", "").replace("CN", "").replace("AB", "")
        return s.lstrip("0") or None

    df["__code"] = df[inv_col].apply(norm_code)
    df = df.dropna(subset=["__code"])
    if df.empty:
        return pd.DataFrame(columns=[inv_col, "Amount"])

    agg = df.groupby("__code", as_index=False)["__net"].sum()
    agg.rename(columns={"__code": inv_col, "__net": "Amount"}, inplace=True)
    agg = agg[agg["Amount"].round(2) != 0].copy()

    ref_cols = [inv_col]
    for opt in [f"reason_{tag}", f"date_{tag}"]:
        if opt in df.columns:
            ref_cols.append(opt)
    ref = df.groupby("__code").first().reset_index()[ref_cols]
    agg = pd.merge(agg, ref, left_on=inv_col, right_on="__code", how="left").drop(columns=["__code"], errors="ignore")

    st.markdown(
        f"<div style='background:#E3F2FD;padding:0.8rem;border-radius:8px;margin-bottom:0.5rem;'>"
        f"🧮 <b>{tag.upper()}</b> aggregation complete — "
        f"{len(df)} → {len(agg)} unique invoices.</div>",
        unsafe_allow_html=True
    )

    return agg.reset_index(drop=True)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df, ven_df):
    def doc_type(row, tag):
        txt = (str(row.get(f"reason_{tag}", "")) + " " + str(row.get(f"invoice_{tag}", ""))).lower()
        if any(k in txt for k in ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "ακυρωτικό"]):
            return "CN"
        if any(k in txt for k in ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]):
            return "INV"
        return "UNKNOWN"

    erp_df["__amt"] = erp_df.get("Amount", erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1))
    ven_df["__amt"] = ven_df.get("Amount", ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1))

    matched = []
    for _, e in erp_df.iterrows():
        e_inv, e_amt = str(e.get("invoice_erp", "")).strip(), round(float(e.get("__amt", 0.0)), 2)
        for _, v in ven_df.iterrows():
            v_inv, v_amt = str(v.get("invoice_ven", "")).strip(), round(float(v.get("__amt", 0.0)), 2)
            if e_inv == v_inv:
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
                break

    matched_df = pd.DataFrame(matched)
    miss_erp = erp_df[~erp_df["invoice_erp"].isin(matched_df["ERP Invoice"] if not matched_df.empty else [])]
    miss_ven = ven_df[~ven_df["invoice_ven"].isin(matched_df["Vendor Invoice"] if not matched_df.empty else [])]
    miss_erp = miss_erp.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})
    miss_ven = miss_ven.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    return matched_df, miss_erp, miss_ven

# ==================== EXCEL EXPORT =========================
def export_excel(miss_erp, miss_ven):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Missing")
    cur = 1
    if not miss_ven.empty:
        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=max(3, miss_ven.shape[1]))
        ws.cell(cur, 1, "Missing in ERP").font = Font(bold=True, size=14)
        cur += 2
        for r in dataframe_to_rows(miss_ven, index=False, header=True): ws.append(r)
        cur = ws.max_row + 3
    if not miss_erp.empty:
        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=max(3, miss_erp.shape[1]))
        ws.cell(cur, 1, "Missing in Vendor").font = Font(bold=True, size=14)
        cur += 2
        for r in dataframe_to_rows(miss_erp, index=False, header=True): ws.append(r)
    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(len(str(c.value or "")) for c in col) + 3
    buf = BytesIO(); wb.save(buf); buf.seek(0); return buf

# ==================== MAIN APP ==========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")
        st.write("🧩 ERP columns detected:", list(erp_df.columns))
        st.write("🧩 Vendor columns detected:", list(ven_df.columns))

        # --- Aggregate before matching ---
        erp_df = aggregate_duplicates(erp_df, "erp")
        ven_df = aggregate_duplicates(ven_df, "ven")

        matched, miss_erp, miss_ven = match_invoices(erp_df, ven_df)
        st.success("Reconciliation Complete!")

        st.dataframe(matched, use_container_width=True)
        st.dataframe(miss_erp, use_container_width=True)
        st.dataframe(miss_ven, use_container_width=True)

        excel_buf = export_excel(miss_erp, miss_ven)
        st.download_button("Download Excel Report", data=excel_buf, file_name="ReconRaptor_Report.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check that your files contain columns like: **invoice**, **debit/credit**, **date**, **reason**")
