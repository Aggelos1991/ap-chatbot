# --------------------------------------------------------------
# ReconRaptor – FINAL, TIER-1 PRESERVED, PAYMENTS FIXED
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
    .payment-match {background:#1565C0;color:#fff;font-weight:bold;}
    .payment-erp {background:#1976D2;color:#fff;font-weight:bold;}
    .payment-vendor {background:#0D47A1;color:#fff;font-weight:bold;}
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 1.3rem; color: #555;'>Intelligent Vendor Invoice & Payment Reconciliation</p>", unsafe_allow_html=True)

# ====================== HELPERS ==========================
def fuzzy_ratio(a, b): 
    return SequenceMatcher(None, str(a), str(b)).ratio()

def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "": return 0.0
    s = str(v).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if "," in s and "." in s and s.find(".") < s.find(","):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." in s and s.find(",") < s.find("."):
        s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        parts = s.split(".")
        s = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(s)
    except:
        return 0.0

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

# ==================== NORMALIZE COLUMNS (SAFE) ====================
def normalize_columns(df, tag):
    mapping = {
        "invoice": ["invoice","factura","document","nº","αρ.","παραστατικό","τιμολόγιο","no","fac","bill","code"],
        "credit": ["credit","haber","abono","πιστωτικό","crédito"],
        "debit": ["debit","debe","importe","amount","ποσό","αξία","débito","charg"],
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
                break
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

# ==================== DOCUMENT TYPE (SAFE: PAYMENT ONLY IF CREDIT > DEBIT + KEYWORD) ====================
def doc_type(row, tag):
    r = str(row.get(f"reason_{tag}", "")).lower()
    debit = normalize_number(row.get(f"debit_{tag}", 0))
    credit = normalize_number(row.get(f"credit_{tag}", 0))
    invoice = str(row.get(f"invoice_{tag}", "")).strip()

    # IGNORE NON-TRANSACTIONS
    if any(x in r for x in ["previous", "fiscal", "year", "balance", "carry", "forward", "opening"]):
        return "IGNORE"

    # PAYMENT: credit > debit AND payment keyword
    pay_keywords = ["cobro", "pago", "payment", "paid", "bank", "receipt", "πληρωμ", "κατάθεση"]
    if credit > debit and credit > 0.01 and any(k in r for k in pay_keywords):
        return "PAYMENT"

    # CREDIT NOTE
    cn_keywords = ["credit", "nota", "abono", "cn", "πιστωτικό", "devolución"]
    if any(k in r for k in cn_keywords):
        return "CN"

    # INVOICE: has invoice number OR debit > 0
    inv_keywords = ["factura", "invoice", "inv", "τιμολόγιο", "fac", "bill"]
    if invoice and invoice != "nan" and len(invoice) > 2:
        return "INV"
    if debit > 0.01:
        return "INV"

    return "UNKNOWN"

# ==================== INVOICE MATCHING (UNCHANGED) ====================
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

    erp_inv_cn = merge_inv reclaimed(erp_inv_cn, "invoice_erp")
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

# ... [TIERS 2 & 3, PAYMENT MATCHING, EXCEL, UI — SAME AS BEFORE] ...
