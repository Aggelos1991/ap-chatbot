# --------------------------------------------------------------
# ReconRaptor – Vendor Reconciliation (FINAL, Smart Consolidation + Fixed Export)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
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
    for fmt in ["%d/%m/%Y","%Y-%m-%d","%Y/%m/%d","%d-%m-%Y","%m/%d/%Y","%d/%m/%y","%m/%d/%y"]:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d):
                return d.strftime("%Y-%m-%d")
        except:
            continue
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
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
        "invoice": ["invoice", "factura", "document", "ref", "αρ.", "τιμολόγιο"],
        "credit":  ["credit", "abono", "πιστωτικό"],
        "debit":   ["debit", "importe", "amount", "ποσό"],
        "reason":  ["reason", "descripcion", "αιτιολογία", "description"],
        "date":    ["date", "fecha", "ημερομηνία"]
    }
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"
    out = df.rename(columns=rename_map)
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    for c in ["debit","credit"]:
        col = f"{c}_{tag}"
        if col not in out.columns: out[col] = 0.0
    return out

def style(df, css):
    return df.style.apply(lambda _: [css]*len(_), axis=1)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor = set()

    def doc_type(row, tag):
        r = str(row.get(f"reason_{tag}", "")).lower()
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        if any(k in r for k in ["credit","abono","cn","πιστωτικό"]):
            return "CN"
        if any(k in r for k in ["invoice","factura","τιμολόγιο"]) or debit > 0:
            return "INV"
        return "UNKNOWN"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)

    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1)

    def consolidate_by_invoice(df, inv_col):
        records = []
        for inv, group in df.groupby(inv_col, dropna=False):
            total = 0.0
            for _, row in group.iterrows():
                amt = normalize_number(row.get("__amt", 0))
                total += amt if row.get("__type", "INV") == "INV" else -amt
            if abs(total) < 0.01:
                continue
            base = group.iloc[0].copy()
            base["__amt"] = abs(total)
            base["__type"] = "INV" if total > 0 else "CN"
            records.append(base)
        return pd.DataFrame(records)

    erp_use = consolidate_by_invoice(erp_df, "invoice_erp")
    ven_use = consolidate_by_invoice(ven_df, "invoice_ven")

    for e_idx, e in erp_use.iterrows():
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor: continue
            if e["invoice_erp"] == v["invoice_ven"] and e["__type"] == v["__type"]:
                diff = abs(e["__amt"] - v["__amt"])
                matched.append({
                    "ERP Invoice": e["invoice_erp"],
                    "Vendor Invoice": v["invoice_ven"],
                    "ERP Amount": e["__amt"],
                    "Vendor Amount": v["__amt"],
                    "Difference": diff,
                    "Status": "Perfect Match" if diff<=0.01 else "Difference Match"
                })
                used_vendor.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    matched_erp = set(matched_df["ERP Invoice"]) if not matched_df.empty else set()
    matched_ven = set(matched_df["Vendor Invoice"]) if not matched_df.empty else set()

    miss_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven)][["invoice_erp","__amt","date_erp"]].rename(columns={"invoice_erp":"Invoice","__amt":"Amount","date_erp":"Date"})
    miss_ven = ven_use[~ven_use["invoice_ven"].isin(matched_erp)][["invoice_ven","__amt","date_ven"]].rename(columns={"invoice_ven":"Invoice","__amt":"Amount","date_ven":"Date"})
    return matched_df, miss_erp, miss_ven

def tier2_match(e, v):
    matches=[]
    used_e,used_v=set(),set()
    for ei,er in e.iterrows():
        e_inv,e_amt=er["Invoice"],round(float(er["Amount"]),2)
        e_code=clean_invoice_code(e_inv)
        for vi,vr in v.iterrows():
            if vi in used_v: continue
            v_inv,v_amt=vr["Invoice"],round(float(vr["Amount"]),2)
            v_code=clean_invoice_code(v_inv)
            sim=fuzzy_ratio(e_code,v_code)
            diff=abs(e_amt-v_amt)
            if diff<0.05 and sim>=0.8:
                matches.append({
                    "ERP Invoice":e_inv,"Vendor Invoice":v_inv,
                    "ERP Amount":e_amt,"Vendor Amount":v_amt,
                    "Difference":diff,"Fuzzy Score":round(sim,2),"Match Type":"Tier-2"
                })
                used_e.add(ei);used_v.add(vi);break
    mdf=pd.DataFrame(matches)
    return mdf, e[~e.index.isin(used_e)], v[~v.index.isin(used_v)]

def tier3_match(e, v):
    matches=[]
    used_e,used_v=set(),set()
    e["d"]=e["Date"].apply(normalize_date)
    v["d"]=v["Date"].apply(normalize_date)
    for ei,er in e.iterrows():
        for vi,vr in v.iterrows():
            if vr["d"]==er["d"] and fuzzy_ratio(clean_invoice_code(er["Invoice"]),clean_invoice_code(vr["Invoice"]))>=0.75:
                diff=abs(er["Amount"]-vr["Amount"])
                matches.append({
                    "ERP Invoice":er["Invoice"],"Vendor Invoice":vr["Invoice"],
                    "ERP Amount":er["Amount"],"Vendor Amount":vr["Amount"],
                    "Difference":diff,"Date":er["d"],"Match Type":"Tier-3"
                })
                used_e.add(ei);used_v.add(vi);break
    mdf=pd.DataFrame(matches)
    return mdf, e[~e.index.isin(used_e)], v[~v.index.isin(used_v)]

# ==================== EXPORT (Fixed: Single Tab + Header Color Only) =========================
def export_excel(t1, t2, t3, miss_erp, miss_ven, pay_match):
    wb = Workbook()
    ws = wb.active
    ws.title = "Missing"

    def color_header(row_idx, color):
        for c in ws[row_idx]:
            c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            c.font = Font(color="FFFFFF", bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

    row = 1
    if not miss_erp.empty:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=miss_erp.shape[1])
        ws.cell(row, 1, "Missing in ERP").font = Font(bold=True, size=14)
        row += 2
        for r_idx, r in enumerate(dataframe_to_rows(miss_erp, index=False, header=True), start=row):
            ws.append(r)
            if r_idx == row:
                color_header(row, "C62828")
        row = ws.max_row + 3

    if not miss_ven.empty:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=miss_ven.shape[1])
        ws.cell(row, 1, "Missing in Vendor").font = Font(bold=True, size=14)
        row += 2
        for r_idx, r in enumerate(dataframe_to_rows(miss_ven, index=False, header=True), start=row):
            ws.append(r)
            if r_idx == row:
                color_header(row, "AD1457")

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
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")

        with st.spinner("Reconciling..."):
            tier1, miss_erp, miss_ven = match_invoices(erp_df, ven_df)
            tier2, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            tier3, final_erp_miss, final_ven_miss = tier3_match(miss_erp2, miss_ven2)

        st.success("Reconciliation Complete!")

        st.markdown('<h2 class="section-title">Download Report</h2>', unsafe_allow_html=True)
        excel_buf = export_excel(tier1, tier2, tier3, final_erp_miss, final_ven_miss, pd.DataFrame())
        st.download_button(
            label="Download Missing Report (Excel)",
            data=excel_buf,
            file_name="ReconRaptor_Missing.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Ensure your Excel files include columns for invoice, debit/credit, date, and reason.")
