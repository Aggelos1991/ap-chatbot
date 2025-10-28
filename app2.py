# --------------------------------------------------------------
# ReconRaptor — Vendor Reconciliation (Final Fixed Build)
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font
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
.metric-box {
    border-radius: 12px;
    padding: 1.5rem;
    margin: 0.5rem;
    text-align: center;
    color: white;
    font-weight: 600;
    box-shadow: 0 4px 8px rgba(0,0,0,0.1);
}
.green {background: #2E7D32;}
.orange {background: #FF8F00;}
.teal {background: #26A69A;}
.purple {background: #7E57C2;}
.red {background: #C62828;}
.pink {background: #AD1457;}
.dark {background: #004D40;}
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
    s = str(v).strip()
    s = re.sub(r"[^\d/.-]", "", s)
    s = s.replace(".", "/").replace("-", "/")
    formats = ["%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m-%d-%Y"]
    for fmt in formats:
        try:
            d = pd.to_datetime(s, format=fmt, errors='raise')
            return d.strftime("%Y-%m-%d")
        except:
            continue
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
        "reason": ["reason", "motivo", "concepto", "descripcion", "detalle", "descripción", "περιγραφή"],
        "date": ["date", "fecha", "fech", "data", "issue date", "posting date", "ημερομηνία"]
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
    if f"reason_{tag}" not in out.columns:
        out[f"reason_{tag}"] = ""
    return out

# ==================== CONSOLIDATION ==========================
payment_keywords = [
    "remittances to suppliers", "payment receipts", "payments", "remesa", "pago", "recibo de pago",
    "πληρωμή", "απόδειξη πληρωμής", "εισπράξεις", "εξόφληση", "πληρωμές", "remittance", "pay", "receipt", "payment"
]

def consolidate_by_invoice(df: pd.DataFrame, inv_col: str, tag: str) -> pd.DataFrame:
    if inv_col not in df.columns:
        return pd.DataFrame(columns=df.columns)
    owing_col = f"debit_{tag}" if tag == "ven" else f"credit_{tag}"
    paid_col = f"credit_{tag}" if tag == "ven" else f"debit_{tag}"
    date_col = f"date_{tag}"
    reason_col = f"reason_{tag}"
    df['__id'] = df.index
    with_inv = df[df[inv_col].notna() & (df[inv_col] != '')].copy()
    without_inv = df[df[inv_col].isna() | (df[inv_col] == '')].copy()

    records = []

    for inv, group in with_inv.groupby(inv_col):
        owing = group[owing_col].apply(normalize_number).sum()
        paid = group[paid_col].apply(normalize_number).sum()
        net_owing = owing - paid
        if abs(net_owing) <= 0.01:
            continue
        base = group.iloc[0].copy()
        base["__amt"] = abs(net_owing)
        base["__type"] = "INV" if net_owing >= 0 else "CN"
        base['__group_ids'] = list(group['__id'])
        records.append(base)

    for idx, row in without_inv.iterrows():
        owing = normalize_number(row[owing_col])
        paid = normalize_number(row[paid_col])
        reason_text = str(row.get(reason_col, "")).lower()
        is_payment = paid > 0 and any(kw in reason_text for kw in payment_keywords)

        if is_payment:
            row["__amt"] = paid
            row["__type"] = "PAY"
            row['__group_ids'] = [row['__id']]
            records.append(row)
        elif owing > 0:
            row["__amt"] = owing
            row["__type"] = "OTHER"
            row['__group_ids'] = [row['__id']]
            records.append(row)

    cons_df = pd.DataFrame(records)
    return cons_df[cons_df['__type'].isin(['INV', 'CN', 'PAY'])].reset_index(drop=True)

# ==================== MATCHING CORE ==========================
def match_invoices(erp_use, ven_use):
    inv_erp = erp_use[erp_use['__type'].isin(['INV','CN'])].copy()
    pay_erp = erp_use[erp_use['__type']=='PAY'].copy()
    inv_ven = ven_use[ven_use['__type'].isin(['INV','CN'])].copy()
    pay_ven = ven_use[ven_use['__type']=='PAY'].copy()

    inv_erp['invoice_erp'] = inv_erp['invoice_erp'].astype(str).str.strip().str.upper()
    inv_ven['invoice_ven'] = inv_ven['invoice_ven'].astype(str).str.strip().str.upper()

    matched_perfect, matched_diff = [], []
    matched_tier2, matched_tier3 = [], []
    used_erp_inv = set()
    used_ven_inv = set()

    # Tier 1: Exact invoice number
    for e_idx, e in inv_erp.iterrows():
        if e_idx in used_erp_inv: continue
        e_inv = e["invoice_erp"]
        e_amt = round(e["__amt"], 2)
        e_date = e.get("date_erp", "")
        for v_idx, v in inv_ven.iterrows():
            if v_idx in used_ven_inv: continue
            v_inv = v["invoice_ven"]
            v_amt = round(v["__amt"], 2)
            v_date = v.get("date_ven", "")
            if e_inv != v_inv: continue
            diff = abs(e_amt - v_amt)
            if diff <= 0.01:
                matched_perfect.append({
                    "ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt,
                    "Vendor Amount": v_amt, "Difference": 0.0, "ERP Date": e_date,
                    "Vendor Date": v_date, "Status": "Perfect Match"
                })
                used_erp_inv.add(e_idx)
                used_ven_inv.add(v_idx)
                break
            elif diff < 1.0:
                matched_diff.append({
                    "ERP Invoice": e_inv, "Vendor Invoice": v_inv, "ERP Amount": e_amt,
                    "Vendor Amount": v_amt, "Difference": round(diff, 2), "ERP Date": e_date,
                    "Vendor Date": v_date, "Status": "Difference Match"
                })
                used_erp_inv.add(e_idx)
                used_ven_inv.add(v_idx)
                break

    remain_erp = inv_erp[~inv_erp.index.isin(used_erp_inv)]
    remain_ven = inv_ven[~inv_ven.index.isin(used_ven_inv)]

    # Tier 2: Fuzzy invoice (>=85%), diff <=600
    for e_idx, e in remain_erp.iterrows():
        if e_idx in used_erp_inv: continue
        e_inv = e["invoice_erp"]
        e_amt = round(e["__amt"], 2)
        e_date = e.get("date_erp", "")
        best_vidx = None
        best_ratio = 0
        best_diff = float('inf')
        for v_idx, v in remain_ven.iterrows():
            if v_idx in used_ven_inv: continue
            v_inv = v["invoice_ven"]
            v_amt = round(v["__amt"], 2)
            ratio = fuzzy_ratio(e_inv, v_inv)
            diff = abs(e_amt - v_amt)
            if ratio >= 0.85 and diff <= 600:
                if ratio > best_ratio or (ratio == best_ratio and diff < best_diff):
                    best_ratio = ratio
                    best_diff = diff
                    best_vidx = v_idx
        if best_vidx is not None:
            v = remain_ven.loc[best_vidx]
            matched_tier2.append({
                "ERP Invoice": e_inv, "Vendor Invoice": v["invoice_ven"],
                "ERP Amount": e_amt, "Vendor Amount": round(v["__amt"], 2),
                "Difference": round(best_diff, 2), "ERP Date": e_date,
                "Vendor Date": v.get("date_ven", ""), "Fuzzy Ratio": round(best_ratio, 2),
                "Status": "Tier-2 Fuzzy Diff"
            })
            used_erp_inv.add(e_idx)
            used_ven_inv.add(best_vidx)

    remain_erp = inv_erp[~inv_erp.index.isin(used_erp_inv)]
    remain_ven = inv_ven[~inv_ven.index.isin(used_ven_inv)]

    # Tier 3: Fuzzy (>=85%), exact amount + date (both must be present and match)
    for e_idx, e in remain_erp.iterrows():
        if e_idx in used_erp_inv: continue
        e_inv = e["invoice_erp"]
        e_amt = round(e["__amt"], 2)
        e_date = e.get("date_erp", "")
        if not e_date: continue  # Skip if no date
        best_vidx = None
        best_ratio = 0
        for v_idx, v in remain_ven.iterrows():
            if v_idx in used_ven_inv: continue
            v_inv = v["invoice_ven"]
            v_amt = round(v["__amt"], 2)
            v_date = v.get("date_ven", "")
            if not v_date: continue  # Skip if no date
            if e_date == v_date and abs(e_amt - v_amt) <= 0.01:
                ratio = fuzzy_ratio(e_inv, v_inv)
                if ratio >= 0.85 and ratio > best_ratio:
                    best_ratio = ratio
                    best_vidx = v_idx
        if best_vidx is not None:
            v = remain_ven.loc[best_vidx]
            matched_tier3.append({
                "ERP Invoice": e_inv, "Vendor Invoice": v["invoice_ven"],
                "ERP Amount": e_amt, "Vendor Amount": v_amt,
                "Difference": 0.0, "ERP Date": e_date,
                "Vendor Date": v.get("date_ven", ""), "Fuzzy Ratio": round(best_ratio, 2),
                "Status": "Tier-3 Fuzzy"
            })
            used_erp_inv.add(e_idx)
            used_ven_inv.add(best_vidx)

    # Unmatched
    unmatch_erp = inv_erp[~inv_erp.index.isin(used_erp_inv)].rename(
        columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"}
    )[['Invoice', 'Amount', 'Date']]
    unmatch_ven = inv_ven[~inv_ven.index.isin(used_ven_inv)].rename(
        columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"}
    )[['Invoice', 'Amount', 'Date']]

    # Payments - match by amount only (no invoice)
    matched_pay = []
    used_erp_pay = set()
    used_ven_pay = set()
    for e_idx, e in pay_erp.iterrows():
        if e_idx in used_erp_pay: continue
        e_amt = round(e["__amt"], 2)
        for v_idx, v in pay_ven.iterrows():
            if v_idx in used_ven_pay: continue
            v_amt = round(v["__amt"], 2)
            if abs(e_amt - v_amt) <= 0.01:
                matched_pay.append({
                    "ERP Amount": e_amt, "Vendor Amount": v_amt, "Difference": 0.0,
                    "ERP Date": e.get("date_erp", ""), "Vendor Date": v.get("date_ven", "")
                })
                used_erp_pay.add(e_idx)
                used_ven_pay.add(v_idx)
                break

    unmatch_pay_erp = pay_erp[~pay_erp.index.isin(used_erp_pay)].rename(
        columns={"__amt": "Amount", "date_erp": "Date"}
    )[['Amount', 'Date']]
    unmatch_pay_ven = pay_ven[~pay_ven.index.isin(used_ven_pay)].rename(
        columns={"__amt": "Amount", "date_ven": "Date"}
    )[['Amount', 'Date']]

    return (
        pd.DataFrame(matched_perfect), pd.DataFrame(matched_diff),
        pd.DataFrame(matched_tier2), pd.DataFrame(matched_tier3),
        unmatch_erp, unmatch_ven,
        pd.DataFrame(matched_pay), unmatch_pay_erp, unmatch_pay_ven,
        pay_erp, pay_ven
    )

# ==================== UI ==========================
st.markdown("### Upload Your Files")
uploaded_erp = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
uploaded_vendor = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="vendor")

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    for col in ["invoice_erp", "invoice_ven"]:
        if col in erp_df.columns:
            erp_df[col] = erp_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})
        if col in ven_df.columns:
            ven_df[col] = ven_df[col].astype(str).str.strip().str.upper().replace({"NAN": "", "NONE": "", "<NA>": ""})

    with st.spinner("Reconciling..."):
        erp_use = consolidate_by_invoice(erp_df, "invoice_erp", "erp")
        ven_use = consolidate_by_invoice(ven_df, "invoice_ven", "ven")
        (
            t1_perfect, t1_diff, t2_fuzzy, t3_fuzzy,
            miss_erp, miss_ven,
            pay_match, miss_pay_erp, miss_pay_ven,
            pay_erp, pay_ven
        ) = match_invoices(erp_use, ven_use)

    st.success("Reconciliation complete!")

    tab1, tab2, tab3 = st.tabs(["Summary", "Matches", "Payments"])

    # --- SUMMARY ---
    with tab1:
        c1, c2, c3, c4, c5, c6,c7 = st.columns(7)
        c1.markdown(f"<div class='metric-box green'>Perfect<br><h2>{len(t1_perfect)}</h2></div>", unsafe_allow_html=True)
        c2.markdown(f"<div class='metric-box orange'>Diff (±1)<br><h2>{len(t1_diff)}</h2></div>", unsafe_allow_html=True)
        c3.markdown(f"<div class='metric-box teal'>Tier-2<br><h2>{len(t2_fuzzy)}</h2></div>", unsafe_allow_html=True)
        c4.markdown(f"<div class='metric-box purple'>Tier-3<br><h2>{len(t3_fuzzy)}</h2></div>", unsafe_allow_html=True)
        c5.markdown(f"<div class='metric-box red'>Miss ERP<br><h2>{len(miss_erp)}</h2></div>", unsafe_allow_html=True)
        c6.markdown(f"<div class='metric-box pink'>Miss Ven<br><h2>{len(miss_ven)}</h2></div>", unsafe_allow_html=True)
        c7.markdown(f"<div class='metric-box dark'>Pay Match<br><h2>{len(pay_match)}</h2></div>", unsafe_allow_html=True)

    # --- MATCHES ---
    with tab2:
        st.markdown("### Tier-1 Perfect")
        if t1_perfect.empty:
            st.info("No perfect matches found.")
        else:
            st.dataframe(t1_perfect, use_container_width=True)

        st.markdown("### Tier-1 Difference")
        if t1_diff.empty:
            st.info("No difference matches found.")
        else:
            st.dataframe(t1_diff, use_container_width=True)

        st.markdown("### Tier-2 Fuzzy Diff (≤600€, ≥85%)")
        if t2_fuzzy.empty:
            st.info("No Tier-2 matches found.")
        else:
            st.dataframe(t2_fuzzy, use_container_width=True)

        st.markdown("### Tier-3 Fuzzy (Exact Date+Amt, ≥85%)")
        if t3_fuzzy.empty:
            st.info("No Tier-3 matches found.")
        else:
            st.dataframe(t3_fuzzy, use_container_width=True)

        st.markdown("### Missing in ERP (Present in Vendor)")
        st.dataframe(miss_ven, use_container_width=True)

        st.markdown("### Missing in Vendor (Present in ERP)")
        st.dataframe(miss_erp, use_container_width=True)

    # --- PAYMENTS ---
    with tab3:
        tot_erp = pay_erp['__amt'].sum() if not pay_erp.empty else 0.0
        tot_ven = pay_ven['__amt'].sum() if not pay_ven.empty else 0.0
        p1, p2, p3 = st.columns(3)
        p1.markdown(f"<div class='metric-box green'>Matched<br><h2>{len(pay_match)}</h2></div>", unsafe_allow_html=True)
        p2.markdown(f"<div class='metric-box teal'>ERP Total<br><h2>{tot_erp:.2f}</h2></div>", unsafe_allow_html=True)
        p3.markdown(f"<div class='metric-box purple'>Ven Total<br><h2>{tot_ven:.2f}</h2></div>", unsafe_allow_html=True)

        st.markdown("### Matched Payments")
        if pay_match.empty:
            st.info("No matched payments found.")
        else:
            st.dataframe(pay_match, use_container_width=True)

        st.markdown("### Unmatched ERP Payments")
        st.dataframe(miss_pay_erp, use_container_width=True)

        st.markdown("### Unmatched Vendor Payments")
        st.dataframe(miss_pay_ven, use_container_width=True)

        # Export only unmatched
        wb = Workbook()
        header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        sheets = {
            "Unmatched_ERP_Invoices": miss_erp,
            "Unmatched_Vendor_Invoices": miss_ven,
            "Unmatched_ERP_Payments": miss_pay_erp,
            "Unmatched_Vendor_Payments": miss_pay_ven
        }
        for name, df in sheets.items():
            if df.empty: continue
            ws = wb.create_sheet(name)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(bold=True)
            total = df['Amount'].sum()
            ws.append(["Total", total] + [""] * (ws.max_column - 2))
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
        out = BytesIO()
        wb.save(out)
        st.download_button(
            "Export Unmatched Only",
            data=out.getvalue(),
            file_name="unmatched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
