# --------------------------------------------------------------
# ReconRaptor – Vendor Reconciliation (FINAL FIXED & OPTIMIZED)
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

# ==================== TITLES =========================
st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.3rem; color: #555;'>Intelligent Vendor Invoice Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b): 
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "": return 0.0
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
    if pd.isna(v) or str(v).strip() == "": return ""
    s = str(v).strip().replace(".", "/").replace("-", "/").replace(",", "/")
    formats = [
        "%d/%m/%Y","%d-%m-%Y","%d.%m.%Y","%m/%d/%Y","%Y/%m/%d",
        "%d/%m/%y","%Y.%m.%d"
    ]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors="coerce")
            if not pd.isna(d): 
                return d.strftime("%Y-%m-%d")
        except: 
            continue
    d = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(d): 
        d = pd.to_datetime(s, errors="coerce", dayfirst=False)
    return d.strftime("%Y-%m-%d") if not pd.isna(d) else ""

def clean_invoice_code(v):
    if not v: return ""
    s = str(v).strip().lower()
    parts = re.split(r"[-_]", s)
    for p in reversed(parts):
        if re.fullmatch(r"\d{1,}", p) and not re.fullmatch(r"20[0-3]\d", p):
            s = p.lstrip("0"); break
    s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no|apd|vs)\W*", "", s)
    s = re.sub(r"20\d{2}", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    s = re.sub(r"^0+", "", s)
    s = re.sub(r"[^\d]", "", s)
    return s or "0"

# ==================== NORMALIZE COLUMNS ====================
def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice","factura","document","nº","αρ.","παραστατικό","τιμολόγιο","no"],
        "credit": ["credit","haber","abono","πιστωτικό"],
        "debit": ["debit","debe","importe","amount","ποσό","αξία"],
        "reason": ["reason","motivo","concepto","descripcion","αιτιολογία","περιγραφή"],
        "date": ["date","fecha","ημερομηνία","issue date","posting date"]
    }
    rename_map, cols_lower = {}, {c: str(c).strip().lower() for c in df.columns}
    
    # Invoice
    invoice_matched = False
    for col, low in cols_lower.items():
        if any(a in low for a in mapping["invoice"]):
            rename_map[col] = f"invoice_{tag}"
            invoice_matched = True
            break
    if not invoice_matched:
        if len(df.columns) > 0:
            df.rename(columns={df.columns[0]: f"invoice_{tag}"}, inplace=True)
        else:
            df[f"invoice_{tag}"] = ""

    # Others
    for key, aliases in mapping.items():
        if key == "invoice": continue
        for col, low in cols_lower.items():
            if col in rename_map: continue
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"

    out = df.rename(columns=rename_map)
    for req in ["debit", "credit"]:
        c = f"{req}_{tag}"
        if c not in out.columns: 
            out[c] = 0.0
    if f"date_{tag}" not in out.columns:
        if len(out.columns) > 1:
            out.rename(columns={out.columns[1]: f"date_{tag}"}, inplace=True)
        else:
            out[f"date_{tag}"] = ""
    out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    return out

# ==================== STYLE =========================
def style(df, css): 
    return df.style.apply(lambda _: [css]*len(_), axis=1)

# ==================== MATCHING LOGIC (FIXED) ==========================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_indices = set()

    def doc_type(row, tag):
        r = str(row.get(f"reason_{tag}", "")).lower()
        pay_pat = [r"^πληρωμ", r"^payment", r"^bank\s*transfer", r"^pago", r"^transferencia", r"^paid"]
        if any(re.search(p, r) for p in pay_pat): 
            return "PAYMENT"
        if any(k in r for k in ["credit","nota","abono","cn","πιστωτικό","πίστωση"]): 
            return "CN"
        if any(k in r for k in ["factura","invoice","inv","τιμολόγιο"]) or normalize_number(row.get(f"debit_{tag}", 0)) > 0: 
            return "INV"
        return "UNKNOWN"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)
    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1)

    # Separate payments early
    erp_pay_full = erp_df[erp_df["__type"] == "PAYMENT"].copy()
    ven_pay_full = ven_df[ven_df["__type"] == "PAYMENT"].copy()

    # Work only on INV & CN
    erp_inv_cn = erp_df[erp_df["__type"].isin(["INV", "CN"])].copy()
    ven_inv_cn = ven_df[ven_df["__type"].isin(["INV", "CN"])].copy()

    def merge_inv_cn(df, inv_col):
        out = []
        for inv, g in df.groupby(inv_col, dropna=False):
            if g.empty: continue
            inv_rows = g[g["__type"] == "INV"]
            cn_rows = g[g["__type"] == "CN"]
            if not inv_rows.empty and not cn_rows.empty:
                net = round(abs(inv_rows["__amt"].sum() - cn_rows["__amt"].sum()), 2)
                base = inv_rows.iloc[-1].copy()
                base["__amt"] = net
                out.append(base)
            elif not inv_rows.empty:
                out.append(inv_rows.loc[inv_rows["__amt"].idxmax()])
            elif not cn_rows.empty:
                out.append(cn_rows.loc[cn_rows["__amt"].idxmax()])
        return pd.DataFrame(out).reset_index(drop=True) if out else pd.DataFrame()

    erp_inv_cn = merge_inv_cn(erp_inv_cn, "invoice_erp")
    ven_inv_cn = merge_inv_cn(ven_inv_cn, "invoice_ven")

    # Tier-1: Exact Invoice + Amount Match
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

    # Remove matched from remaining
    erp_remaining = erp_inv_cn[~erp_inv_cn["invoice_erp"].astype(str).isin(matched_ven_inv)]
    ven_remaining = ven_inv_cn[~ven_inv_cn["invoice_ven"].astype(str).isin(matched_erp_inv)]

    miss_erp = erp_remaining[["invoice_erp", "__amt"] + (["date_erp"] if "date_erp" in erp_remaining.columns else [])] \
        .rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    miss_ven = ven_remaining[["invoice_ven", "__amt"] + (["date_ven"] if "date_ven" in ven_remaining.columns else [])] \
        .rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})

    return tier1_df, miss_erp, miss_ven, erp_pay_full, ven_pay_full

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

def extract_payments(erp_pay_df, ven_pay_df, matched_erp_inv, matched_ven_inv):
    if erp_pay_df.empty or ven_pay_df.empty:
        return pd.DataFrame(columns=["ERP Reason", "Vendor Reason", "ERP Amount", "Vendor Amount", "Difference"])
    
    # Remove payment lines that reference matched invoices
    def has_matched_inv(reason, inv_set):
        if pd.isna(reason): return False
        txt = str(reason).lower()
        for inv in inv_set:
            if str(inv).lower() in txt:
                return True
        return False

    erp_pay_df = erp_pay_df[~erp_pay_df["reason_erp"].apply(lambda x: has_matched_inv(x, matched_ven_inv))]
    ven_pay_df = ven_pay_df[~ven_pay_df["reason_ven"].apply(lambda x: has_matched_inv(x, matched_erp_inv))]

    erp_pay_df["Amount"] = erp_pay_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_pay_df["Amount"] = ven_pay_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0))), axis=1)

    matched = []
    used = set()
    for _, e in erp_pay_df.iterrows():
        for vi, v in ven_pay_df.iterrows():
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
    return pd.DataFrame(matched)

# ==================== EXCEL EXPORT =========================
def export_excel(t1, t2, t3, miss_erp, miss_ven, pay_match):
    wb = Workbook()
    def hdr(ws, row, color):
        for c in ws[row]:
            c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            c.font = Font(color="FFFFFF", bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")

    ws1 = wb.active; ws1.title = "Tier1"
    if not t1.empty: 
        for r in dataframe_to_rows(t1, index=False, header=True): 
            ws1.append(r)
        hdr(ws1, 1, "1E88E5")

    ws2 = wb.create_sheet("Tier2")
    if not t2.empty: 
        for r in dataframe_to_rows(t2, index=False, header=True): 
            ws2.append(r)
        hdr(ws2, 1, "26A69A")

    ws3 = wb.create_sheet("Tier3")
    if not t3.empty: 
        for r in dataframe_to_rows(t3, index=False, header=True): 
            ws3.append(r)
        hdr(ws3, 1, "7E57C2")

    ws4 = wb.create_sheet("Missing"); cur = 1
    if not miss_erp.empty:
        ws4.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=miss_erp.shape[1])
        ws4.cell(cur, 1, "Missing in ERP").font = Font(bold=True, size=14); cur += 2
        for r in dataframe_to_rows(miss_erp, index=False, header=True): 
            ws4.append(r)
        hdr(ws4, cur, "C62828"); cur = ws4.max_row + 3
    if not miss_ven.empty:
        ws4.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=miss_ven.shape[1])
        ws4.cell(cur, 1, "Missing in Vendor").font = Font(bold=True, size=14); cur += 2
        for r in dataframe_to_rows(miss_ven, index=False, header=True): 
            ws4.append(r)
        hdr(ws4, cur, "AD1457")

    ws5 = wb.create_sheet("Payments")
    if not pay_match.empty: 
        for r in dataframe_to_rows(pay_match, index=False, header=True): 
            ws5.append(r)
        hdr(ws5, 1, "004D40")

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3

    buf = BytesIO(); wb.save(buf); buf.seek(0); return buf

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

        with st.spinner("Analyzing invoices and payments..."):
            tier1, miss_erp, miss_ven, erp_pay_full, ven_pay_full = match_invoices(erp_df, ven_df)
            tier2, _, _, miss_erp2, miss_ven2 = tier2_match(miss_erp, miss_ven)
            tier3, _, _, final_erp_miss, final_ven_miss = tier3_match(miss_erp2, miss_ven2)

            # Combine all matched invoices
            all_matched_erp = set(tier1["ERP Invoice"].dropna().astype(str)) | \
                              set(tier2["ERP Invoice"].dropna().astype(str)) | \
                              set(tier3["ERP Invoice"].dropna().astype(str))
            all_matched_ven = set(tier1["Vendor Invoice"].dropna().astype(str)) | \
                              set(tier2["Vendor Invoice"].dropna().astype(str)) | \
                              set(tier3["Vendor Invoice"].dropna().astype(str))

            pay_match = extract_payments(erp_pay_full, ven_pay_full, all_matched_erp, all_matched_ven)

        st.success("Reconciliation Complete!")

        # ---------- METRICS ----------
        st.markdown('<h2 class="section-title">Reconciliation Summary</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5, c6 = st.columns(6)
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
            st.metric("Unmatched ERP", len(final_erp_miss))
            st.markdown(f"**Total:** {final_erp_miss['Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c6:
            st.markdown('<div class="metric-container missing-vendor">', unsafe_allow_html=True)
            st.metric("Unmatched Vendor", len(final_ven_miss))
            st.markdown(f"**Total:** {final_ven_miss['Amount'].sum():,.2f}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")

        # ---------- DISPLAY ----------
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
            st.markdown('<h2 class="section-title">Missing in ERP</h2>', unsafe_allow_html=True)
            if not final_erp_miss.empty:
                st.dataframe(style(final_erp_miss, "background:#C62828;color:#fff;font-weight:bold;"), use_container_width=True)
                st.error(f"{len(final_erp_miss)} ERP invoices missing – {final_erp_miss['Amount'].sum():,.2f}")
            else: st.success("All ERP invoices found in vendor.")
        with col_m2:
            st.markdown('<h2 class="section-title">Missing in Vendor</h2>', unsafe_allow_html=True)
            if not final_ven_miss.empty:
                st.dataframe(style(final_ven_miss, "background:#AD1457;color:#fff;font-weight:bold;"), use_container_width=True)
                st.error(f"{len(final_ven_miss)} vendor invoices missing – {final_ven_miss['Amount'].sum():,.2f}")
            else: st.success("All vendor invoices found in ERP.")

        st.markdown('<h2 class="section-title">Payment Transactions</h2>', unsafe_allow_html=True)
        if not pay_match.empty:
            st.dataframe(style(pay_match, "background:#004D40;color:#fff;font-weight:bold;"), use_container_width=True)
            st.success(f"{len(pay_match)} payment transactions matched.")
        else:
            st.info("No unmatched payment transactions found.")

        st.markdown('<h2 class="section-title">Download Report</h2>', unsafe_allow_html=True)
        excel_buf = export_excel(tier1, tier2, tier3, final_erp_miss, final_ven_miss, pay_match)
        st.download_button(
            "Download Full Excel Report",
            data=excel_buf,
            file_name="ReconRaptor_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("Check that your files contain columns like: **invoice**, **debit/credit**, **date**, **reason**")
