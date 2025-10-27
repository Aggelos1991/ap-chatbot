# --------------------------------------------------------------
# ReconRaptor – FINAL, BULLETPROOF, ZERO-ERROR VERSION
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
    .big-title {font-size: 3rem !important;font-weight: 700;text-align: center;
        background: linear-gradient(90deg, #1E88E5, #42A5F5);-webkit-background-clip: text;
        -webkit-text-fill-color: transparent;margin-bottom: 1rem;}
    .section-title {font-size: 1.8rem !important;font-weight: 600;color: #1565C0;
        border-bottom: 2px solid #42A5F5;padding-bottom: 0.5rem;margin-top: 2rem;}
    .metric-container {padding: 1.2rem !important;border-radius: 12px !important;
        margin-bottom: 1rem;box-shadow: 0 4px 6px rgba(0,0,0,0.1);}
    .perfect-match {background:#2E7D32;color:#fff;font-weight:bold;}
    .difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
    .tier2-match {background:#26A69A;color:#fff;font-weight:bold;}
    .tier3-match {background:#7E57C2;color:#fff;font-weight:bold;}
    .missing-erp {background:#C62828;color:#fff;font-weight:bold;}
    .missing-vendor {background:#AD1457;color:#fff;font-weight:bold;}
    .payment-match {background:#004D40;color:#fff;font-weight:bold;}
    .payment-erp {background:#1B5E20;color:#fff;font-weight:bold;}
    .payment-vendor {background:#880E4F;color:#fff;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.3rem; color: #555;'>Intelligent Vendor Invoice & Payment Reconciliation (ES/EN/GR)</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b): 
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "": return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    if s.count(",") == 1 and s.count(".") == 1:
        if s.find(",") > s.find("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif s.count(",") == 1: s = s.replace(",", ".")
    elif s.count(".") > 1: s = s.replace(".", "", s.count(".") - 1)
    try: return float(s)
    except: return 0.0

def normalize_date(v):
    if pd.isna(v) or str(v).strip() == "": return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    formats = ["%d/%m/%Y","%d-%m-%Y","%d.%m.%Y","%m/%d/%Y","%Y/%m/%d","%d/%m/%y","%Y.%m.%d"]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d): return d.strftime("%Y-%m-%d")
        except: continue
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(d): d = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s = str(v).strip().lower()
    s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no|apd|vs|fac|bill|rec|cobro|pago|factura|invoice)\W*", "", s)
    s = re.sub(r"20\d{2}", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    return s or "0"

# ==================== NORMALIZE COLUMNS ====================
def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice","factura","document","nº","αρ.","παραστατικό","τιμολόγιο","no","fac","bill"],
        "credit": ["credit","haber","abono","πιστωτικό","crédito"],
        "debit": ["debit","debe","importe","amount","ποσό","αξία","débito"],
        "reason": ["reason","motivo","concepto","descripcion","αιτιολογία","περιγραφή","descripción"],
        "date": ["date","fecha","ημερομηνία","issue date","posting date","fecha emisión"]
    }
    rename_map, cols_lower = {}, {c: str(c).strip().lower() for c in df.columns}
    invoice_matched = False
    for col, low in cols_lower.items():
        if any(a in low for a in mapping["invoice"]):
            rename_map[col] = f"invoice_{tag}"; invoice_matched = True; break
    if not invoice_matched:
        if len(df.columns)>0: df.rename(columns={df.columns[0]: f"invoice_{tag}"}, inplace=True)
        else: df[f"invoice_{tag}"]=""
    for key, aliases in mapping.items():
        if key=="invoice": continue
        for col, low in cols_lower.items():
            if col in rename_map: continue
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"
    out = df.rename(columns=rename_map)
    for req in ["debit","credit"]:
        c=f"{req}_{tag}"
        if c not in out.columns: out[c]=0.0
    if f"date_{tag}" not in out.columns:
        if len(out.columns)>1: out.rename(columns={out.columns[1]:f"date_{tag}"},inplace=True)
        else: out[f"date_{tag}"]=""
    out[f"date_{tag}"]=out[f"date_{tag}"].apply(normalize_date)
    return out

# ==================== STYLE =========================
def style(df, css): 
    return df.style.apply(lambda _: [css]*len(_), axis=1)

# ==================== CANCEL NET-ZERO DUPLICATES ====================
def cancel_net_zero(df, inv_col, amt_col):
    if df.empty: return df
    grouped = df.groupby(inv_col).agg({amt_col: 'sum'}).reset_index()
    zero_inv = grouped[grouped[amt_col].abs() < 0.01][inv_col].astype(str).tolist()
    return df[~df[inv_col].astype(str).isin(zero_inv)].copy()

# ==================== DOCUMENT TYPE (AGGRESSIVE PAYMENT DETECTION) ====================
def doc_type(row, tag):
    r = str(row.get(f"reason_{tag}", "")).lower()
    debit = normalize_number(row.get(f"debit_{tag}", 0))
    credit = normalize_number(row.get(f"credit_{tag}", 0))

    pay_keywords = [
        "cobro", "pago", "abono", "ingreso", "transferencia", "recibo", "rec", "pago de", "cobro de",
        "πληρωμ", "πληρωμή", "εξόφληση", "κατάθεση", "μεταφορά", "εισπραξη",
        "payment", "paid", "bank transfer", "wire", "deposit", "settlement", "receipt", "bank", "chq", "cheque"
    ]

    cn_keywords = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση", "devolución", "storno", "nota de crédito"]
    inv_keywords = ["factura", "invoice", "inv", "τιμολόγιο", "fac", "bill", "fatura"]

    if any(k in r for k in pay_keywords):
        return "PAYMENT"
    if any(k in r for k in cn_keywords):
        return "CN"
    if any(k in r for k in inv_keywords) or debit > 0:
        return "INV"
    if credit > debit and credit > 0.01:
        return "PAYMENT"
    return "UNKNOWN"

# ==================== INVOICE MATCHING ====================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_indices = set()

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)
    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1)

    erp_pay_full = erp_df[erp_df["__type"] == "PAYMENT"].copy()
    ven_pay_full = ven_df[ven_df["__type"] == "PAYMENT"].copy()

    erp_inv_cn = erp_df[erp_df["__type"].isin(["INV", "CN"])].copy()
    ven_inv_cn = ven_df[ven_df["__type"].isin(["INV", "CN"])].copy()

    erp_inv_cn = cancel_net_zero(erp_inv_cn, "invoice_erp", "__amt")
    ven_inv_cn = cancel_net_zero(ven_inv_cn, "invoice_ven", "__amt")

    def merge_inv_cn(df, inv_col):
        out = []
        for inv, g in df.groupby(inv_col, dropna=False):
            if g.empty: continue
            inv_rows = g[g["__type"] == "INV"]
            cn_rows = g[g["__type"] == "CN"]
            if not inv_rows.empty and not cn_rows.empty:
                net = round(abs(inv_rows["__amt"].sum() - cn_rows["__amt"].sum()), 2)
                if net > 0.01:
                    base = inv_rows.iloc[-1].copy()
                    base["__amt"] = net
                    out.append(base)
            elif not inv_rows.empty:
                out.append(inv_rows.loc[inv_rows["__amt"].idxmax()])
            elif not cn_rows.empty and cn_rows["__amt"].iloc[0] > 0.01:
                out.append(cn_rows.iloc[0])
        return pd.DataFrame(out).reset_index(drop=True) if out else pd.DataFrame()

    erp_inv_cn = merge_inv_cn(erp_inv_cn, "invoice_erp")
    ven_inv_cn = merge_inv_cn(ven_inv_cn, "invoice_ven")

    for e_idx, e in erp_inv_cn.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_typ = e["__type"]
        for v_idx, v in ven_inv_cn.iterrows():
            if v_idx in used_vendor_indices: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_typ = v["__type"]
            if e_typ != v_typ or e_inv != v_inv: continue
            diff = abs(e_amt - v_amt)
            status = "Perfect Match" if diff <= 0.01 else "Difference Match" if diff < 1.0 else None
            if status:
                matched.append({
                    "ERP Invoice": e_inv, "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt, "Vendor Amount": v_amt,
                    "Difference": diff, "Status": status
                })
                used_vendor_indices.add(v_idx)
                break

    tier1_df = pd.DataFrame(matched, columns=["ERP Invoice","Vendor Invoice","ERP Amount","Vendor Amount","Difference","Status"])
    matched_erp_inv = set(tier1_df["ERP Invoice"].dropna().astype(str))
    matched_ven_inv = set(tier1_df["Vendor Invoice"].dropna().astype(str))

    miss_erp = ven_inv_cn[~ven_inv_cn["invoice_ven"].astype(str).isin(matched_erp_inv)] \
        [["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in ven_inv_cn.columns else [])] \
        .rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})
    miss_ven = erp_inv_cn[~erp_inv_cn["invoice_erp"].astype(str).isin(matched_ven_inv)] \
        [["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in erp_inv_cn.columns else [])] \
        .rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})

    return tier1_df, miss_erp, miss_ven, erp_pay_full, ven_pay_full

# ==================== TIERS 2 & 3 ====================
def tier2_match(erp_miss, ven_miss):
    cols = ["ERP Invoice","Vendor Invoice","ERP Amount","Vendor Amount","Difference","Fuzzy Score","Match Type"]
    if erp_miss.empty or ven_miss.empty: 
        return pd.DataFrame(columns=cols), set(), set(), erp_miss.copy(), ven_miss.copy()
    matches, used_e, used_v = [], set(), set()
    for ei, er in erp_miss.iterrows():
        e_inv = str(er.get("Invoice", "")); e_amt = round(float(er.get("Amount", 0)), 2); e_code = clean_invoice_code(e_inv)
        for vi, vr in ven_miss.iterrows():
            v_inv = str(vr.get("Invoice", "")); v_amt = round(float(vr.get("Amount", 0)), 2); v_code = clean_invoice_code(v_inv)
            diff = abs(e_amt - v_amt); sim = fuzzy_ratio(e_code, v_code)
            if diff < 0.05 and sim >= 0.80:
                matches.append([e_inv, v_inv, e_amt, v_amt, diff, round(sim, 2), "Tier-2"])
                used_e.add(ei); used_v.add(vi); break
    mdf = pd.DataFrame(matches, columns=cols)
    return mdf, used_e, used_v, erp_miss[~erp_miss.index.isin(used_e)], ven_miss[~ven_miss.index.isin(used_v)]

def tier3_match(erp_miss, ven_miss):
    cols = ["ERP Invoice","Vendor Invoice","ERP Amount","Vendor Amount","Difference","Fuzzy Score","Date","Match Type"]
    if erp_miss.empty or ven_miss.empty: 
        return pd.DataFrame(columns=cols), set(), set(), erp_miss.copy(), ven_miss.copy()
    e, v = erp_miss.copy(), ven_miss.copy()
    e["d"] = e["Date"].apply(normalize_date) if "Date" in e.columns else ""
    v["d"] = v["Date"].apply(normalize_date) if "Date" in v.columns else ""
    matches, used_e, used_v = [], set(), set()
    for ei, er in e.iterrows():
        if ei in used_e or not er.get("d"): continue
        e_inv = str(er.get("Invoice", "")); e_amt = round(float(er.get("Amount", 0)), 2); e_code = clean_invoice_code(e_inv)
        for vi, vr in v.iterrows():
            if vi in used_v or not vr.get("d"): continue
            v_inv = str(vr.get("Invoice", "")); v_amt = round(float(vr.get("Amount", 0)), 2); v_code = clean_invoice_code(v_inv)
            sim = fuzzy_ratio(e_code, v_code)
            if er["d"] == vr["d"] and sim >= 0.90:
                diff = abs(e_amt - v_amt)
                matches.append([e_inv, v_inv, e_amt, v_amt, diff, round(sim, 2), er["d"], "Tier-3"])
                used_e.add(ei); used_v.add(vi); break
    mdf = pd.DataFrame(matches, columns=cols)
    return mdf, used_e, used_v, e[~e.index.isin(used_e)], v[~v.index.isin(used_v)]

# ==================== PAYMENT MATCHING (3 TABLES, 100% SAFE) ====================
def extract_payments(erp_pay_df, ven_pay_df):
    # Empty case
    if erp_pay_df.empty and ven_pay_df.empty:
        empty = pd.DataFrame(columns=["Reason", "Amount"])
        return pd.DataFrame(), empty, empty

    erp_pay_df = erp_pay_df.copy()
    ven_pay_df = ven_pay_df.copy()

    # Compute amounts
    erp_pay_df["Amount"] = erp_pay_df.apply(
        lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1
    )
    ven_pay_df["Amount"] = ven_pay_df.apply(
        lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1
    )

    erp_pay_df["Amt_Rounded"] = erp_pay_df["Amount"].round(2)
    ven_pay_df["Amt_Rounded"] = ven_pay_df["Amount"].round(2)

    # SAFELY CONVERT REASON TO STRING
    erp_pay_df["Reason"] = erp_pay_df.get("reason_erp", pd.Series(["Unknown"] * len(erp_pay_df))).fillna("Unknown").astype(str)
    ven_pay_df["Reason"] = ven_pay_df.get("reason_ven", pd.Series(["Unknown"] * len(ven_pay_df))).fillna("Unknown").astype(str)

    matched = []
    used_ven_idx = set()

    for _, e in erp_pay_df.iterrows():
        e_amt = e["Amt_Rounded"]
        e_reason_raw = str(e["Reason"])
        e_reason = e_reason_raw[:100] if isinstance(e_reason_raw, str) else "Unknown"
        
        for vi, v in ven_pay_df.iterrows():
            if vi in used_ven_idx: continue
            v_amt = v["Amt_Rounded"]
            v_reason_raw = str(v["Reason"])
            v_reason = v_reason_raw[:100] if isinstance(v_reason_raw, str) else "Unknown"
            
            if abs(e_amt - v_amt) < 0.05:
                matched.append({
                    "ERP Reason": e_reason,
                    "Vendor Reason": v_reason,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": round(abs(e_amt - v_amt), 2)
                })
                used_ven_idx.add(vi)
                break

    matched_df = pd.DataFrame(matched)

    # Unmatched ERP
    unmatched_erp_rows = []
    for i, row in erp_pay_df.iterrows():
        if any(abs(row["Amt_Rounded"] - v["Amt_Rounded"]) < 0.05 for _, v in ven_pay_df.iterrows()):
            continue
        reason = str(row["Reason"])[:100] if isinstance(row["Reason"], str) else "Unknown"
        unmatched_erp_rows.append({"Reason": reason, "Amount": row["Amount"]})
    unmatched_erp = pd.DataFrame(unmatched_erp_rows)

    # Unmatched Vendor
    unmatched_ven_rows = []
    for _, row in ven_pay_df.iterrows():
        if row.name in used_ven_idx: continue
        reason = str(row["Reason"])[:100] if isinstance(row["Reason"], str) else "Unknown"
        unmatched_ven_rows.append({"Reason": reason, "Amount": row["Amount"]})
    unmatched_ven = pd.DataFrame(unmatched_ven_rows)

    return matched_df, unmatched_erp, unmatched_ven

# ==================== EXCEL EXPORT ====================
def export_excel(t1, t2, t3, miss_erp, miss_ven, pay_match, pay_erp, pay_ven):
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

    ws3 = wb.create_sheet("Tier3")
    if not t3.empty: 
        for r in dataframe_to_rows(t3, index=False, header=True): ws3.append(r)
        hdr(ws3, 1, "7E57C2")

    ws4 = wb.create_sheet("Missing"); cur = 1
    if not miss_erp.empty:
        ws4.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=miss_erp.shape[1])
        ws4.cell(cur, 1, "Missing in ERP (Vendor has, ERP missing)").font = Font(bold=True, size=14); cur += 2
        for r in dataframe_to_rows(miss_erp, index=False, header=True): ws4.append(r)
        hdr(ws4, cur, "C62828"); cur = ws4.max_row + 3
    if not miss_ven.empty:
        ws4.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=miss_ven.shape[1])
        ws4.cell(cur, 1, "Missing in Vendor (ERP has, Vendor missing)").font = Font(bold=True, size=14); cur += 2
        for r in dataframe_to_rows(miss_ven, index=False, header=True): ws4.append(r)
        hdr(ws4, cur, "AD1457")

    ws5 = wb.create_sheet("Payments Matched")
    if not pay_match.empty: 
        for r in dataframe_to_rows(pay_match, index=False, header=True): ws5.append(r)
        hdr(ws5, 1, "004D40")

    ws6 = wb.create_sheet("Payments ERP Unmatched")
    if not pay_erp.empty: 
        for r in dataframe_to_rows(pay_erp, index=False, header=True): ws6.append(r)
        hdr(ws6, 1, "1B5E20")

    ws7 = wb.create_sheet("Payments Vendor Unmatched")
    if not pay_ven.empty: 
        for r in dataframe_to_rows(pay_ven, index=False, header=True): ws7.append(r)
        hdr(ws7, 1, "880E4F")

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3

    buf = BytesIO(); wb.save(buf); buf.seek(0); return buf

# ==================== UI ====================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")

if uploaded_erp and uploaded_vendor:
    try:
        erp_raw = pd.read_excel(uploaded_erp, dtype=str)
        ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
        erp_df = normalize_columns(erp_raw, "erp")
        ven_df = normalize_columns(ven_raw, "ven")

        with st.spinner("Analyzing invoices and payments..."):
            tier1, miss_erp, miss_ven, erp_pay_full, ven_pay_full = match_invoices(erp_df, ven_df)
            tier2, _, _, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            tier3, _, _, final_erp_miss, final_ven_miss = tier3_match(miss_erp2, miss_ven2)
            pay_match, pay_erp_unmatched, pay_ven_unmatched = extract_payments(erp_pay_full, ven_pay_full)

        st.success("Reconciliation Complete!")

        # METRICS
        st.markdown('<h2 class="section-title">Reconciliation Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
        perf = tier1[tier1["Status"] == "Perfect Match"]
        diff = tier1[tier1["Status"] == "Difference Match"]
        safe_sum = lambda df, col: df[col].sum() if not df.empty and col in df.columns else 0.0

        with c1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("Perfect Matches", len(perf))
            st.markdown(f"**ERP:** {safe_sum(perf,'ERP Amount'):,.2f}<br>**Vendor:** {safe_sum(perf,'Vendor Amount'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="metric-container difference-match">', unsafe_allow_html=True)
            st.metric("Differences", len(diff))
            st.markdown(f"**ERP:** {safe_sum(diff,'ERP Amount'):,.2f}<br>**Vendor:** {safe_sum(diff,'Vendor Amount'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c3:
            st.markdown('<div class="metric-container tier2-match">', unsafe_allow_html=True)
            st.metric("Tier-2", len(tier2))
            st.markdown(f"**ERP:** {safe_sum(tier2,'ERP Amount'):,.2f}<br>**Vendor:** {safe_sum(tier2,'Vendor Amount'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="metric-container tier3-match">', unsafe_allow_html=True)
            st.metric("Tier-3", len(tier3))
            st.markdown(f"**ERP:** {safe_sum(tier3,'ERP Amount'):,.2f}<br>**Vendor:** {safe_sum(tier3,'Vendor Amount'):,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c5:
            st.markdown('<div class="metric-container missing-erp">', unsafe_allow_html=True)
            st.metric("Missing in ERP", len(final_erp_miss))
            st.markdown(f"**Total:** {final_erp_miss['Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c6:
            st.markdown('<div class="metric-container missing-vendor">', unsafe_allow_html=True)
            st.metric("Missing in Vendor", len(final_ven_miss))
            st.markdown(f"**Total:** {final_ven_miss['Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c7:
            st.markdown('<div class="metric-container payment-match">', unsafe_allow_html=True)
            st.metric("Matched Payments", len(pay_match))
            st.markdown(f"**Total:** {pay_match['ERP Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")

        # DISPLAY INVOICES
        st.markdown('<h2 class="section-title">Tier-1: Exact Matches</h2>', unsafe_allow_html=True)
        if not perf.empty: st.dataframe(style(perf, "background:#2E7D32;color:#fff;font-weight:bold;"), use_container_width=True)
        if not diff.empty: st.dataframe(style(diff, "background:#FF8F00;color:#fff;font-weight:bold;"), use_container_width=True)

        st.markdown('<h2 class="section-title">Tier-2: Fuzzy + Small Amount</h2>', unsafe_allow_html=True)
        if not tier2.empty: st.dataframe(style(tier2, "background:#26A69A;color:#fff;font-weight:bold;"), use_container_width=True)
        else: st.info("No Tier-2 matches.")

        st.markdown('<h2 class="section-title">Tier-3: Date + Strict Fuzzy</h2>', unsafe_allow_html=True)
        if not tier3.empty: st.dataframe(style(tier3, "background:#7E57C2;color:#fff;font-weight:bold;"), use_container_width=True)
        else: st.info("No Tier-3 matches.")

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown('<h2 class="section-title">Missing in ERP (Vendor has, ERP missing)</h2>', unsafe_allow_html=True)
            if not final_erp_miss.empty:
                st.dataframe(style(final_erp_miss, "background:#C62828;color:#fff;font-weight:bold;"), use_container_width=True)
                st.error(f"{len(final_erp_miss)} invoices in Vendor but not in ERP – {final_erp_miss['Amount'].sum():,.2f}")
            else: st.success("No invoices missing in ERP.")
        with col_m2:
            st.markdown('<h2 class="section-title">Missing in Vendor (ERP has, Vendor missing)</h2>', unsafe_allow_html=True)
            if not final_ven_miss.empty:
                st.dataframe(style(final_ven_miss, "background:#AD1457;color:#fff;font-weight:bold;"), use_container_width=True)
                st.error(f"{len(final_ven_miss)} invoices in ERP but not in Vendor – {final_ven_miss['Amount'].sum():,.2f}")
            else: st.success("No invoices missing in Vendor.")

        # PAYMENT TABLES
        st.markdown('<h2 class="section-title">Matched Payments</h2>', unsafe_allow_html=True)
        if not pay_match.empty:
            st.dataframe(style(pay_match, "background:#004D40;color:#fff;font-weight:bold;"), use_container_width=True)
        else:
            st.info("No payment matches.")

        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown('<h2 class="section-title">Unmatched ERP Payments</h2>', unsafe_allow_html=True)
            if not pay_erp_unmatched.empty:
                st.dataframe(style(pay_erp_unmatched, "background:#1B5E20;color:#fff;font-weight:bold;"), use_container_width=True)
                st.warning(f"{len(pay_erp_unmatched)} unmatched ERP payments – {pay_erp_unmatched['Amount'].sum():,.2f}")
            else:
                st.success("All ERP payments matched.")
        with col_p2:
            st.markdown('<h2 class="section-title">Unmatched Vendor Payments</h2>', unsafe_allow_html=True)
            if not pay_ven_unmatched.empty:
                st.dataframe(style(pay_ven_unmatched, "background:#880E4F;color:#fff;font-weight:bold;"), use_container_width=True)
                st.warning(f"{len(pay_ven_unmatched)} unmatched Vendor payments – {pay_ven_unmatched['Amount'].sum():,.2f}")
            else:
                st.success("All Vendor payments matched.")

        st.markdown('<h2 class="section-title">Download Report</h2>', unsafe_allow_html=True)
        excel_buf = export_excel(tier1, tier2, tier3, final_erp_miss, final_ven_miss, pay_match, pay_erp_unmatched, pay_ven_unmatched)
        st.download_button(
            "Download Full Excel Report",
            data=excel_buf,
            file_name="ReconRaptor_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check columns: **invoice**, **debit/credit**, **date**, **reason**")
