# --------------------------------------------------------------
# ReconRaptor – Vendor Reconciliation (FIXED, NO KEYERROR)
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
    .big-title {font-size:3rem !important;font-weight:700;text-align:center;
        background:linear-gradient(90deg,#1E88E5,#42A5F5);-webkit-background-clip:text;
        -webkit-text-fill-color:transparent;margin-bottom:1rem;}
    .section-title {font-size:1.8rem !important;font-weight:600;color:#1565C0;
        border-bottom:2px solid #42A5F5;padding-bottom:.5rem;margin-top:2rem;}
    .metric-container {padding:1.2rem !important;border-radius:12px !important;
        margin-bottom:1rem;box-shadow:0 4px 6px rgba(0,0,0,0.1);}
    .perfect-match {background:#2E7D32;color:#fff;font-weight:bold;}
    .difference-match{background:#FF8F00;color:#fff;font-weight:bold;}
    .tier2-match {background:#26A69A;color:#fff;font-weight:bold;}
    .tier3-match {background:#7E57C2;color:#fff;font-weight:bold;}
    .missing-erp {background:#C62828;color:#fff;font-weight:bold;}
    .missing-vendor {background:#AD1457;color:#fff;font-weight:bold;}
    .payment-match {background:#004D40;color:#fff;font-weight:bold;}
</style>
""",
    unsafe_allow_html=True,
)
st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align:center;font-size:1.3rem;color:#555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "": return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    s = s.replace(",", ".").replace("..", ".")
    try:
        return float(s)
    except:
        return 0.0

def normalize_date(v):
    if pd.isna(v) or str(v).strip() == "": return ""
    s = re.sub(r"[^\d\/\-\.]", "", str(v).strip())
    formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
        "%m/%d/%Y", "%m-%d-%Y",
        "%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d",
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y",
        "%m/%d/%y", "%m-%d-%y",
    ]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d):
                return d.strftime("%Y-%m-%d")
        except:
            continue
    d = pd.to_datetime(s, errors="coerce")
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s = str(v).strip().lower()
    s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no|apd|vs)\W*", "", s)
    s = re.sub(r"20\d{2}", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    return s or "0"

def normalize_invoice(v):
    return re.sub(r"\s+", "", str(v)).strip().upper()

# ==================== EXCLUDE MATCHED ====================
def exclude_matched(df: pd.DataFrame, col: str, matched_set: set) -> pd.DataFrame:
    if df.empty or col not in df.columns:
        return df.copy()
    norm = df[col].astype(str).apply(normalize_invoice)
    return df[~norm.isin(matched_set)].copy()

# ==================== NORMALIZE COLUMNS ====================
def normalize_columns(df: pd.DataFrame, tag: str) -> pd.DataFrame:
    mapping = {
        "invoice": ["invoice","factura","fact","nº","num","numero","número","document","doc",
                    "ref","referencia","nº factura","num factura","alternative document",
                    "document number","αρ.","αριθμός","νουμερο","νούμερο","no","παραστατικό"],
        "credit": ["credit","haber","credito","crédito","nota de crédito","abono","importe haber",
                   "πίστωση","πιστωτικό","πιστωτικό τιμολόγιο"],
        "debit": ["debit","debe","cargo","importe","amount","total","monto","χρέωση","αξία","ποσό"],
        "reason": ["reason","motivo","concepto","descripcion","detalle","razon","observaciones",
                   "αιτιολογία","περιγραφή","παρατηρήσεις","description"],
        "date": ["date","fecha","fech","data","issue date","ημερομηνία","ημ/νία"]
    }
    rename = {}
    lower_cols = {c: str(c).strip().lower() for c in df.columns}

    # ---- invoice (mandatory) ----
    for col, low in lower_cols.items():
        if any(k in low for k in mapping["invoice"]):
            rename[col] = f"invoice_{tag}"
            break
    else:
        st.error(f"Could not find **invoice** column for **{tag}** file. "
                 "Check that a column contains words like *invoice*, *factura*, *αρ.*, *no* …")
        st.stop()

    # ---- optional columns ----
    for key, aliases in mapping.items():
        if key == "invoice": continue
        for col, low in lower_cols.items():
            if col in rename: continue
            if any(a in low for a in aliases):
                rename[col] = f"{key}_{tag}"
                break

    df = df.rename(columns=rename)

    # ---- guarantee required columns ----
    for col in [f"debit_{tag}", f"credit_{tag}", f"date_{tag}", f"reason_{tag}"]:
        if col not in df.columns:
            df[col] = 0.0 if "debit" in col or "credit" in col else ""

    # ---- clean types ----
    df[f"debit_{tag}"] = df[f"debit_{tag}"].apply(normalize_number)
    df[f"credit_{tag}"] = df[f"credit_{tag}"].apply(normalize_number)
    df[f"date_{tag}"] = df[f"date_{tag}"].apply(normalize_date)

    return df

# ==================== NET INVOICES ====================
def net_invoices(df: pd.DataFrame, inv_col: str) -> pd.DataFrame:
    out = []
    for inv, g in df.groupby(inv_col, dropna=False):
        inv_str = str(inv).strip()
        if not inv_str or inv_str.lower() in {"none", "nan", ""}:
            continue
        inv_rows = g[g["__type"] == "INV"]
        cn_rows  = g[g["__type"] == "CN"]
        net = inv_rows["__amt"].sum() - cn_rows["__amt"].sum()
        net = round(net, 2)
        if abs(net) < 0.01:
            continue
        base = inv_rows.loc[inv_rows["__amt"].idxmax()] if not inv_rows.empty else cn_rows.iloc[0]
        base = base.copy()
        base["__amt"] = net
        base["__type"] = "INV" if net > 0 else "CN"
        out.append(base)
    return pd.DataFrame(out).reset_index(drop=True) if out else pd.DataFrame(columns=df.columns)

# ==================== TIER-1 ====================
def tier1_match(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    # ---- document type ----
    def doc_type(row, tag):
        r = str(row.get(f"reason_{tag}", "")).lower()
        if any(p in r for p in ["payment","πληρωμ","remittance","pago","transferencia","trf","paid","εξοφληση"]):
            return "IGNORE"
        if any(k in r for k in ["credit","nota","abono","cn","πιστωτικό"]):
            return "CN" if row.get(f"credit_{tag}", 0) > 0 else "INV"
        if any(k in r for k in ["factura","invoice","inv","τιμολόγιο"]) or row.get(f"debit_{tag}", 0) > 0:
            return "INV"
        return "UNKNOWN"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)

    erp_df["__amt"] = erp_df["debit_erp"] - erp_df["credit_erp"]
    ven_df["__amt"] = ven_df["debit_ven"] - ven_df["credit_ven"]

    erp_use = erp_df[erp_df["__type"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__type"] != "IGNORE"].copy()

    erp_use = erp_use[erp_use["invoice_erp"].notna() & (erp_use["invoice_erp"].str.strip() != "")]
    ven_use = ven_use[ven_use["invoice_ven"].notna() & (ven_use["invoice_ven"].str.strip() != "")]

    erp_use = net_invoices(erp_use, "invoice_erp")
    ven_use = net_invoices(ven_use, "invoice_ven")

    if erp_use.empty or ven_use.empty:
        empty = pd.DataFrame(columns=["ERP Invoice","Vendor Invoice","ERP Amount","Vendor Amount","Difference","Status"])
        return empty, set(), set()

    erp_use["__inv_norm"] = erp_use["invoice_erp"].apply(normalize_invoice)
    ven_use["__inv_norm"] = ven_use["invoice_ven"].apply(normalize_invoice)
    erp_use["__abs_amt"] = erp_use["__amt"].abs().round(2)
    ven_use["__abs_amt"] = ven_use["__amt"].abs().round(2)

    matched, used_e, used_v = [], set(), set()

    for ei, er in erp_use.iterrows():
        if ei in used_e: continue
        for vi, vr in ven_use.iterrows():
            if vi in used_v: continue
            if er["__inv_norm"] == vr["__inv_norm"] and abs(er["__abs_amt"] - vr["__abs_amt"]) < 0.01:
                matched.append({
                    "ERP Invoice": er["invoice_erp"],
                    "Vendor Invoice": vr["invoice_ven"],
                    "ERP Amount": er["__abs_amt"],
                    "Vendor Amount": vr["__abs_amt"],
                    "Difference": 0.0,
                    "Status": "Perfect Match"
                })
                used_e.add(ei)
                used_v.add(vi)
                break

    mdf = pd.DataFrame(matched)
    erp_set = {normalize_invoice(x) for x in mdf.get("ERP Invoice", [])}
    ven_set = {normalize_invoice(x) for x in mdf.get("Vendor Invoice", [])}
    return mdf, erp_set, ven_set

# ==================== TIER-2 ====================
def tier2_match(erp_miss: pd.DataFrame, ven_miss: pd.DataFrame):
    if erp_miss.empty or ven_miss.empty:
        return pd.DataFrame(), set(), set()

    matches, used_e, used_v = [], set(), set()
    for ei, er in erp_miss.iterrows():
        if ei in used_e: continue
        e_code = clean_invoice_code(er["Invoice"])
        e_amt = abs(round(float(er["Amount"]), 2))
        for vi, vr in ven_miss.iterrows():
            if vi in used_v: continue
            v_code = clean_invoice_code(vr["Invoice"])
            v_amt = abs(round(float(vr["Amount"]), 2))
            diff = abs(e_amt - v_amt)
            sim = fuzzy_ratio(e_code, v_code)
            if diff < 0.05 and sim >= 0.80:
                matches.append({
                    "ERP Invoice": er["Invoice"],
                    "Vendor Invoice": vr["Invoice"],
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Fuzzy Score": round(sim, 2),
                    "Match Type": "Tier-2"
                })
                used_e.add(ei)
                used_v.add(vi)
                break
    return pd.DataFrame(matches), used_e, used_v

# ==================== TIER-3 ====================
def tier3_match(erp_miss: pd.DataFrame, ven_miss: pd.DataFrame):
    if erp_miss.empty or ven_miss.empty:
        return pd.DataFrame(), set(), set()

    e = erp_miss.copy()
    v = ven_miss.copy()
    e["d"] = e["Date"].apply(lambda x: normalize_date(x) if pd.notna(x) else "")
    v["d"] = v["Date"].apply(lambda x: normalize_date(x) if pd.notna(x) else "")

    matches, used_e, used_v = [], set(), set()
    for ei, er in e.iterrows():
        if ei in used_e or not er["d"]: continue
        e_code = clean_invoice_code(er["Invoice"])
        e_amt = abs(round(float(er["Amount"]), 2))
        for vi, vr in v.iterrows():
            if vi in used_v or not vr["d"]: continue
            if er["d"] == vr["d"] and fuzzy_ratio(e_code, clean_invoice_code(vr["Invoice"])) >= 0.90:
                v_amt = abs(round(float(vr["Amount"]), 2))
                diff = abs(e_amt - v_amt)
                matches.append({
                    "ERP Invoice": er["Invoice"],
                    "Vendor Invoice": vr["Invoice"],
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Fuzzy Score": round(fuzzy_ratio(e_code, clean_invoice_code(vr["Invoice"])), 2),
                    "Date": er["d"],
                    "Match Type": "Tier-3"
                })
                used_e.add(ei)
                used_v.add(vi)
                break
    mdf = pd.DataFrame(matches)
    return mdf, used_e, used_v

# ==================== PAYMENTS ====================
def extract_payments(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    pay_kw = ["πληρωμ","payment","remittance","pago","transferencia","trf","paid","εξοφληση"]
    def is_pay(row, tag):
        txt = str(row.get(f"reason_{tag}", "")).lower()
        amt = abs(row.get(f"debit_{tag}", 0) - row.get(f"credit_{tag}", 0))
        return any(k in txt for k in pay_kw) and amt > 0

    erp_pay = erp_df[erp_df.apply(lambda r: is_pay(r, "erp"), axis=1)].copy()
    ven_pay = ven_df[ven_df.apply(lambda r: is_pay(r, "ven"), axis=1)].copy()

    for df, tag in [(erp_pay, "erp"), (ven_pay, "ven")]:
        if not df.empty:
            df["Amount"] = df.apply(lambda r: abs(r[f"debit_{tag}"] - r[f"credit_{tag}"]), axis=1)

    matched, used = [], set()
    for _, e in erp_pay.iterrows():
        for vi, v in ven_pay.iterrows():
            if vi in used: continue
            if abs(e["Amount"] - v["Amount"]) < 0.05:
                matched.append({
                    "ERP Reason": e.get("reason_erp", ""),
                    "Vendor Reason": v.get("reason_ven", ""),
                    "ERP Amount": round(e["Amount"], 2),
                    "Vendor Amount": round(v["Amount"], 2),
                    "Difference": round(abs(e["Amount"] - v["Amount"]), 2)
                })
                used.add(vi)
                break
    return erp_pay, ven_pay, pd.DataFrame(matched)

# ==================== EXCEL EXPORT ====================
def export_excel(t1, t2, t3, miss_erp, miss_ven, pay_match):
    wb = Workbook()
    def hdr(ws, row, color):
        for c in ws[row]:
            c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            c.font = Font(color="FFFFFF", bold=True)
            c.alignment = Alignment(horizontal="center")

    ws = wb.active; ws.title = "Tier1"
    if not t1.empty:
        for r in dataframe_to_rows(t1, index=False, header=True): ws.append(r)
        hdr(ws, 1, "1E88E5")

    ws2 = wb.create_sheet("Tier2")
    if not t2.empty:
        for r in dataframe_to_rows(t2, index=False, header=True): ws2.append(r)
        hdr(ws2, 1, "26A69A")

    ws3 = wb.create_sheet("Tier3")
    if not t3.empty:
        for r in dataframe_to_rows(t3, index=False, header=True): ws3.append(r)
        hdr(ws3, 1, "7E57C2")

    ws4 = wb.create_sheet("Missing")
    cur = 1
    if not miss_erp.empty:
        ws4.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=miss_erp.shape[1])
        ws4.cell(cur,1,"Missing in ERP").font = Font(bold=True,size=14)
        cur += 2
        for r in dataframe_to_rows(miss_erp, index=False, header=True): ws4.append(r)
        hdr(ws4, cur, "C62828")
        cur = ws4.max_row + 3
    if not miss_ven.empty:
        ws4.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=miss_ven.shape[1])
        ws4.cell(cur,1,"Missing in Vendor").font = Font(bold=True,size=14)
        cur += 2
        for r in dataframe_to_rows(miss_ven, index=False, header=True): ws4.append(r)
        hdr(ws4, cur, "AD1457")

    ws5 = wb.create_sheet("Payments")
    if not pay_match.empty:
        for r in dataframe_to_rows(pay_match, index=False, header=True): ws5.append(r)
        hdr(ws5, 1, "004D40")

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3

    buf = BytesIO()
    wb.save(buf); buf.seek(0)
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

        with st.spinner("Running reconciliation…"):
            # ---- Tier-1 ----
            tier1, matched_erp_set, matched_ven_set = tier1_match(erp_df, ven_df)

            # ---- Build “missing” tables (absolute amounts) ----
            miss_erp = erp_df[erp_df["invoice_erp"].notna()].copy()
            miss_ven = ven_df[ven_df["invoice_ven"].notna()].copy()

            miss_erp["Invoice"] = miss_erp["invoice_erp"]
            miss_ven["Invoice"] = miss_ven["invoice_ven"]
            miss_erp["Amount"] = miss_erp.apply(lambda r: abs(r["debit_erp"] - r["credit_erp"]), axis=1)
            miss_ven["Amount"] = miss_ven.apply(lambda r: abs(r["debit_ven"] - r["credit_ven"]), axis=1)
            miss_erp["Date"] = miss_erp.get("date_erp", "")
            miss_ven["Date"] = miss_ven.get("date_ven", "")

            miss_erp = exclude_matched(miss_erp, "Invoice", matched_ven_set)
            miss_ven = exclude_matched(miss_ven, "Invoice", matched_erp_set)

            # ---- Tier-2 ----
            tier2, used_e2, used_v2 = tier2_match(miss_erp, miss_ven)
            matched_erp_set.update(normalize_invoice(x) for x in tier2.get("ERP Invoice", []))
            matched_ven_set.update(normalize_invoice(x) for x in tier2.get("Vendor Invoice", []))
            miss_erp = miss_erp.loc[~miss_erp.index.isin(used_e2)]
            miss_ven = miss_ven.loc[~miss_ven.index.isin(used_v2)]

            # ---- Tier-3 ----
            tier3, used_e3, used_v3 = tier3_match(miss_erp, miss_ven)
            matched_erp_set.update(normalize_invoice(x) for x in tier3.get("ERP Invoice", []))
            matched_ven_set.update(normalize_invoice(x) for x in tier3.get("Vendor Invoice", []))

            final_erp_miss = miss_erp.loc[~miss_erp.index.isin(used_e3), ["Invoice","Amount","Date"]]
            final_ven_miss = miss_ven.loc[~miss_ven.index.isin(used_v3), ["Invoice","Amount","Date"]]

            # ---- Payments ----
            erp_pay, ven_pay, pay_match = extract_payments(erp_df, ven_df)

        st.success("Reconciliation complete!")

        # ---------- SUMMARY ----------
        st.markdown('<h2 class="section-title">Reconciliation Summary</h2>', unsafe_allow_html=True)
        c1,c2,c3,c4,c5,c6,c7 = st.columns(7)

        perf = tier1[tier1["Status"] == "Perfect Match"] if not tier1.empty else pd.DataFrame()
        with c1:
            st.markdown('<div class="metric-container perfect-match">', unsafe_allow_html=True)
            st.metric("Perfect Matches", len(perf))
            st.markdown(f"**Total:** {perf['ERP Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c3:
            st.markdown('<div class="metric-container tier2-match">', unsafe_allow_html=True)
            st.metric("Tier-2", len(tier2))
            st.markdown(f"**Total:** {tier2['ERP Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c4:
            st.markdown('<div class="metric-container tier3-match">', unsafe_allow_html=True)
            st.metric("Tier-3", len(tier3))
            st.markdown(f"**Total:** {tier3['ERP Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c5:
            st.markdown('<div class="metric-container missing-erp">', unsafe_allow_html=True)
            st.metric("Unmatched ERP", len(final_erp_miss))
            st.markdown(f"**Total:** {final_erp_miss['Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c6:
            st.markdown('<div class="metric-container missing-vendor">', unsafe_allow_html=True)
            st.metric("Unmatched Vendor", len(final_ven_miss))
            st.markdown(f"**Total:** {final_ven_miss['Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c7:
            st.markdown('<div class="metric-container payment-match">', unsafe_allow_html=True)
            st.metric("Matched Payments", len(pay_match))
            st.markdown(f"**Total:** {pay_match['ERP Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # ---------- DISPLAY ----------
        st.markdown("---")
        st.markdown('<h2 class="section-title">Tier-1: Exact Matches</h2>', unsafe_allow_html=True)
        if not tier1.empty:
            st.dataframe(tier1.style.apply(lambda _: ['background:#2E7D32;color:#fff;font-weight:bold;']*len(_), axis=1),
                         use_container_width=True)
        else:
            st.info("No Tier-1 matches.")

        st.markdown('<h2 class="section-title">Tier-2: Fuzzy + Small Diff</h2>', unsafe_allow_html=True)
        if not tier2.empty:
            st.dataframe(tier2.style.apply(lambda _: ['background:#26A69A;color:#fff;font-weight:bold;']*len(_), axis=1),
                         use_container_width=True)
        else:
            st.info("No Tier-2 matches.")

        st.markdown('<h2 class="section-title">Tier-3: Date + Strict Fuzzy</h2>', unsafe_allow_html=True)
        if not tier3.empty:
            st.dataframe(tier3.style.apply(lambda _: ['background:#7E57C2;color:#fff;font-weight:bold;']*len(_), axis=1),
                         use_container_width=True)
        else:
            st.info("No Tier-3 matches.")

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown('<h2 class="section-title">Missing in ERP</h2>', unsafe_allow_html=True)
            if not final_ven_miss.empty:
                st.dataframe(final_ven_miss.style.apply(lambda _: ['background:#AD1457;color:#fff;font-weight:bold;']*len(_), axis=1),
                             use_container_width=True)
            else:
                st.success("All vendor invoices found in ERP.")

        with col_m2:
            st.markdown('<h2 class="section-title">Missing in Vendor</h2>', unsafe_allow_html=True)
            if not final_erp_miss.empty:
                st.dataframe(final_erp_miss.style.apply(lambda _: ['background:#C62828;color:#fff;font-weight:bold;']*len(_), axis=1),
                             use_container_width=True)
            else:
                st.success("All ERP invoices found in vendor.")

        st.markdown('<h2 class="section-title">Payment Transactions</h2>', unsafe_allow_html=True)
        if not pay_match.empty:
            st.dataframe(pay_match.style.apply(lambda _: ['background:#004D40;color:#fff;font-weight:bold;']*len(_), axis=1),
                         use_container_width=True)

        # ---------- EXPORT ----------
        st.markdown('<h2 class="section-title">Download Report</h2>', unsafe_allow_html=True)
        excel_buf = export_excel(tier1, tier2, tier3, final_erp_miss, final_ven_miss, pay_match)
        st.download_button(
            label="Download Full Excel Report",
            data=excel_buf,
            file_name="ReconRaptor_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("**Required columns:** *invoice* (any language), *debit/credit* (or amount), optional *date* & *reason*.")
