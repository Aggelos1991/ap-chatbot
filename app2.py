# =========================================================
# ReconRaptor — Vendor Reconciliation (STRICT DATE FINAL)
# =========================================================
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font
from difflib import SequenceMatcher

# =============== PAGE CONFIG & CSS ==================
st.set_page_config(page_title="ReconRaptor — Vendor Reconciliation", layout="wide")
st.markdown(
    """
<style>
.big-title{font-size:3rem!important;font-weight:700;text-align:center;
background:linear-gradient(90deg,#1E88E5,#42A5F5);
-webkit-background-clip:text;-webkit-text-fill-color:transparent;margin-bottom:1rem;}
.metric-box{border-radius:12px;padding:1.5rem;margin:0.5rem;text-align:center;
color:white;font-weight:600;box-shadow:0 4px 8px rgba(0,0,0,0.1);}
.green{background:#2E7D32;}.orange{background:#FF8F00;}
.teal{background:#26A69A;}.purple{background:#7E57C2;}
.red{background:#C62828;}.pink{background:#AD1457;}.dark{background:#004D40;}
</style>
""",
    unsafe_allow_html=True,
)
st.markdown('<h1 class="big-title">ReconRaptor</h1>', unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center;font-size:1.3rem;color:#555;'>Intelligent Vendor Invoice Reconciliation</p>",
    unsafe_allow_html=True,
)

# =============== HELPERS ==================
def fuzzy_ratio(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()


def normalize_number(v):
    if pd.isna(v) or str(v).strip() == "":
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", str(v).strip())
    # European 1.234,56  → 1234.56
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
    except Exception:
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
        "invoice": [
            "invoice",
            "factura",
            "fact",
            "nº",
            "num",
            "numero",
            "número",
            "document",
            "doc",
            "ref",
            "referencia",
            "nº factura",
            "num factura",
        ],
        "credit": ["credit", "haber", "credito", "crédito", "abono"],
        "debit": [
            "debit",
            "debe",
            "cargo",
            "importe",
            "valor",
            "amount",
            "total",
        ],
        "reason": [
            "reason",
            "motivo",
            "concepto",
            "descripcion",
            "detalle",
            "descripción",
            "περιγραφή",
        ],
        "date": [
            "date",
            "fecha",
            "fech",
            "data",
            "issue date",
            "posting date",
            "ημερομηνία",
        ],
    }
    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}
    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"
    out = df.rename(columns=rename_map)

    # guarantee required columns
    for req in ["debit", "credit"]:
        c = f"{req}_{tag}"
        if c not in out.columns:
            out[c] = 0.0
    if f"date_{tag}" in out.columns:
        out[f"date_{tag}"] = out[f"date_{tag}"].apply(normalize_date)
    if f"reason_{tag}" not in out.columns:
        out[f"reason_{tag}"] = ""
    return out


# =============== PAYMENT KEYWORDS ==================
payment_keywords = [
    "remittance",
    "payment",
    "transfer",
    "bank",
    "pagos",
    "pago",
    "πληρωμή",
    "εισπράξεις",
    "εξόφληση",
    "remesa",
    "receipt",
    "recibo",
    "efectivo",
]

credit_note_keywords = ["credit", "abono", "credito", "crédito", "haber"]


def is_payment_row(row, tag):
    """True if the row is a payment according to the file type."""
    reason = str(row.get(f"reason_{tag}", "")).lower()
    has_keyword = any(k in reason for k in payment_keywords)

    if tag == "erp":
        return has_keyword or normalize_number(row.get("debit_erp", 0)) > 0
    else:  # vendor
        return has_keyword or normalize_number(row.get("credit_ven", 0)) > 0


def is_credit_note_row(row, tag):
    """True if the row is a credit-note (not a payment)."""
    reason = str(row.get(f"reason_{tag}", "")).lower()
    return any(k in reason for k in credit_note_keywords)


# =============== CONSOLIDATION ==================
def consolidate_by_invoice(df, inv_col, tag):
    """
    Returns a dataframe with ONE row per logical transaction:
        INV  → net positive amount (owe)
        CN   → net negative amount (credit note)
        PAY  → pure payment (no invoice or keyword-detected)
    """
    if inv_col not in df.columns:
        return pd.DataFrame(columns=df.columns)

    owing_col = f"debit_{tag}" if tag == "ven" else f"credit_{tag}"
    paid_col = f"credit_{tag}" if tag == "ven" else f"debit_{tag}"
    reason_col = f"reason_{tag}"

    df = df.copy()
    df["__orig_idx"] = df.index

    # ---- 1. Rows WITH invoice number -------------------------------------------------
    with_inv = df[df[inv_col].notna() & (df[inv_col] != "")].copy()
    rec = []

    for inv, grp in with_inv.groupby(inv_col):
        owe = grp[owing_col].apply(normalize_number).sum()
        pay = grp[paid_col].apply(normalize_number).sum()
        net = owe - pay

        # keep even zero-net if it is a pure payment
        base = grp.iloc[0].copy()
        base["__amt"] = abs(net)
        if abs(net) > 0.01:
            base["__type"] = "INV" if net >= 0 else "CN"
        else:
            # net ≈ 0 → treat as PAY if a payment keyword exists
            base["__type"] = "PAY" if is_payment_row(base, tag) else "INV"
        rec.append(base)

    # ---- 2. Rows WITHOUT invoice number -----------------------------------------------
    without_inv = df[df[inv_col].isna() | (df[inv_col] == "")].copy()
    for _, r in without_inv.iterrows():
        owe = normalize_number(r[owing_col])
        pay = normalize_number(r[paid_col])
        row = r.copy()

        if is_payment_row(row, tag):
            row["__amt"] = pay if pay > 0 else owe
            row["__type"] = "PAY"
        elif is_credit_note_row(row, tag):
            row["__amt"] = owe if owe > 0 else pay
            row["__type"] = "CN"
        else:
            row["__amt"] = abs(owe - pay) or 0.0
            row["__type"] = "OTHER"
        rec.append(row)

    out = pd.DataFrame(rec)

    # keep only the three useful types
    out = out[out["__type"].isin(["INV", "CN", "PAY"])].reset_index(drop=True)
    return out


# =============== MATCHING ==================
def match_invoices(erp_use, ven_use):
    # Separate by type
    inv_erp = erp_use[erp_use["__type"].isin(["INV", "CN"])].copy()
    pay_erp = erp_use[erp_use["__type"] == "PAY"].copy()
    inv_ven = ven_use[ven_use["__type"].isin(["INV", "CN"])].copy()
    pay_ven = ven_use[ven_use["__type"] == "PAY"].copy()

    # ---------- INVOICE MATCHING ----------
    matched1, diff1, t2, t3 = [], [], [], []
    used_e_inv, used_v_inv = set(), set()

    # Tier-1: exact invoice + amount (or diff <=1)
    for ei, e in inv_erp.iterrows():
        if ei in used_e_inv:
            continue
        e_inv = str(e.get("invoice_erp", "")).strip().upper()
        e_amt = round(e["__amt"], 2)
        e_date = e.get("date_erp", "")

        for vi, v in inv_ven.iterrows():
            if vi in used_v_inv:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_amt = round(v["__amt"], 2)
            v_date = v.get("date_ven", "")

            if e_inv != v_inv:
                continue
            d = abs(e_amt - v_amt)
            if d <= 0.01:
                matched1.append(
                    {
                        "ERP Invoice": e_inv,
                        "Vendor Invoice": v_inv,
                        "ERP Amount": e_amt,
                        "Vendor Amount": v_amt,
                        "Difference": 0.0,
                        "ERP Date": e_date,
                        "Vendor Date": v_date,
                        "Status": "Perfect Match",
                    }
                )
                used_e_inv.add(ei)
                used_v_inv.add(vi)
                break
            elif d < 1.0:
                diff1.append(
                    {
                        "ERP Invoice": e_inv,
                        "Vendor Invoice": v_inv,
                        "ERP Amount": e_amt,
                        "Vendor Amount": v_amt,
                        "Difference": round(d, 2),
                        "ERP Date": e_date,
                        "Vendor Date": v_date,
                        "Status": "Diff ±1",
                    }
                )
                used_e_inv.add(ei)
                used_v_inv.add(vi)
                break

    # Tier-2: fuzzy invoice + amount diff <=600
    rem_e = inv_erp[~inv_erp.index.isin(used_e_inv)]
    rem_v = inv_ven[~inv_ven.index.isin(used_v_inv)]
    for ei, e in rem_e.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip().upper()
        e_amt = round(e["__amt"], 2)

        best, best_r, best_d = None, 0, float("inf")
        for vi, v in rem_v.iterrows():
            if vi in used_v_inv:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_amt = round(v["__amt"], 2)
            r = fuzzy_ratio(e_inv, v_inv)
            d = abs(e_amt - v_amt)
            if r >= 0.85 and d <= 600 and (r > best_r or (r == best_r and d < best_d)):
                best, best_r, best_d = vi, r, d
        if best is not None:
            v = rem_v.loc[best]
            t2.append(
                {
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": str(v["invoice_ven"]).strip().upper(),
                    "ERP Amount": e_amt,
                    "Vendor Amount": round(v["__amt"], 2),
                    "Difference": round(best_d, 2),
                    "ERP Date": e.get("date_erp", ""),
                    "Vendor Date": v.get("date_ven", ""),
                    "Fuzzy Ratio": round(best_r, 2),
                    "Status": "Tier-2 Fuzzy",
                }
            )
            used_e_inv.add(ei)
            used_v_inv.add(best)

    # Tier-3: same date + fuzzy >=0.85
    rem_e = inv_erp[~inv_erp.index.isin(used_e_inv)]
    rem_v = inv_ven[~inv_ven.index.isin(used_v_inv)]
    for ei, e in rem_e.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip().upper()
        e_amt = round(e["__amt"], 2)
        e_date = normalize_date(e.get("date_erp", ""))
        if not e_date:
            continue

        best, best_r = None, 0
        for vi, v in rem_v.iterrows():
            if vi in used_v_inv:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip().upper()
            v_amt = round(v["__amt"], 2)
            v_date = normalize_date(v.get("date_ven", ""))
            if v_date != e_date or abs(e_amt - v_amt) > 0.01:
                continue
            r = fuzzy_ratio(e_inv, v_inv)
            if r >= 0.85 and r > best_r:
                best, best_r = vi, r
        if best is not None:
            v = rem_v.loc[best]
            t3.append(
                {
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": str(v["invoice_ven"]).strip().upper(),
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": 0.0,
                    "ERP Date": e_date,
                    "Vendor Date": v_date,
                    "Fuzzy Ratio": round(best_r, 2),
                    "Status": "Tier-3 Strict (Same Date)",
                }
            )
            used_e_inv.add(ei)
            used_v_inv.add(best)

    # Missing invoices
    miss_e = (
        inv_erp[~inv_erp.index.isin(used_e_inv)]
        .rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})[
            ["Invoice", "Amount", "Date"]
        ]
        .reset_index(drop=True)
    )
    miss_v = (
        inv_ven[~inv_ven.index.isin(used_v_inv)]
        .rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})[
            ["Invoice", "Amount", "Date"]
        ]
        .reset_index(drop=True)
    )

    # ---------- PAYMENT MATCHING ----------
    matched_pay = []
    used_e_pay, used_v_pay = set(), set()
    for ei, e in pay_erp.iterrows():
        if ei in used_e_pay:
            continue
        e_amt = round(e["__amt"], 2)
        for vi, v in pay_ven.iterrows():
            if vi in used_v_pay:
                continue
            v_amt = round(v["__amt"], 2)
            if abs(e_amt - v_amt) <= 0.01:
                matched_pay.append(
                    {
                        "ERP Amount": e_amt,
                        "Vendor Amount": v_amt,
                        "Difference": 0.0,
                        "ERP Date": e.get("date_erp", ""),
                        "Vendor Date": v.get("date_ven", ""),
                    }
                )
                used_e_pay.add(ei)
                used_v_pay.add(vi)
                break

    miss_pe = (
        pay_erp[~pay_erp.index.isin(used_e_pay)]
        .rename(columns={"__amt": "Amount", "date_erp": "Date"})[["Amount", "Date"]]
        .reset_index(drop=True)
    )
    miss_pv = (
        pay_ven[~pay_ven.index.isin(used_v_pay)]
        .rename(columns={"__amt": "Amount", "date_ven": "Date"})[["Amount", "Date"]]
        .reset_index(drop=True)
    )

    return (
        pd.DataFrame(matched1),
        pd.DataFrame(diff1),
        pd.DataFrame(t2),
        pd.DataFrame(t3),
        miss_e,
        miss_v,
        pd.DataFrame(matched_pay),
        miss_pe,
        miss_pv,
        pay_erp,
        pay_ven,
    )


# =============== UI ==================
st.markdown("### Upload Your Files")
erp_file = st.file_uploader("ERP Export (Excel)", type=["xlsx"], key="erp")
ven_file = st.file_uploader("Vendor Statement (Excel)", type=["xlsx"], key="ven")

if erp_file and ven_file:
    erp_raw = pd.read_excel(erp_file, dtype=str)
    ven_raw = pd.read_excel(ven_file, dtype=str)

    erp = normalize_columns(erp_raw, "erp")
    ven = normalize_columns(ven_raw, "ven")

    # Normalise numeric columns
    for c in ["credit_erp", "debit_erp", "credit_ven", "debit_ven"]:
        if c in erp.columns:
            erp[c] = erp[c].apply(normalize_number)
        if c in ven.columns:
            ven[c] = ven[c].apply(normalize_number)

    # Normalise invoice numbers
    for col in ["invoice_erp", "invoice_ven"]:
        if col in erp.columns:
            erp[col] = erp[col].astype(str).str.strip().str.upper()
        if col in ven.columns:
            ven[col] = ven[col].astype(str).str.strip().str.upper()

    with st.spinner("Reconciling..."):
        e_use = consolidate_by_invoice(erp, "invoice_erp", "erp")
        v_use = consolidate_by_invoice(ven, "invoice_ven", "ven")
        (
            t1,
            t1d,
            t2,
            t3,
            missE,
            missV,
            payM,
            missPE,
            missPV,
            pE,
            pV,
        ) = match_invoices(e_use, v_use)

    st.success("Reconciliation complete!")

    tab1, tab2, tab3 = st.tabs(["Summary", "Matches", "Payments"])

    # ---------- SUMMARY ----------
    with tab1:
        c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
        c1.markdown(
            f"<div class='metric-box green'>Perfect<br><h2>{len(t1)}</h2></div>",
            unsafe_allow_html=True,
        )
        c2.markdown(
            f"<div class='metric-box orange'>Diff<br><h2>{len(t1d)}</h2></div>",
            unsafe_allow_html=True,
        )
        c3.markdown(
            f"<div class='metric-box teal'>Tier-2<br><h2>{len(t2)}</h2></div>",
            unsafe_allow_html=True,
        )
        c4.markdown(
            f"<div class='metric-box purple'>Tier-3<br><h2>{len(t3)}</h2></div>",
            unsafe_allow_html=True,
        )
        c5.markdown(
            f"<div class='metric-box red'>Miss ERP<br><h2>{len(missE)}</h2></div>",
            unsafe_allow_html=True,
        )
        c6.markdown(
            f"<div class='metric-box pink'>Miss Ven<br><h2>{len(missV)}</h2></div>",
            unsafe_allow_html=True,
        )
        c7.markdown(
            f"<div class='metric-box dark'>Pay Match<br><h2>{len(payM)}</h2></div>",
            unsafe_allow_html=True,
        )

    # ---------- MATCHES ----------
    with tab2:
        st.subheader("Tier-1 Perfect")
        st.dataframe(t1 if not t1.empty else pd.DataFrame(), use_container_width=True)

        st.subheader("Tier-1 Difference ±1€")
        st.dataframe(t1d if not t1d.empty else pd.DataFrame(), use_container_width=True)

        st.subheader("Tier-2 Fuzzy (≥85%, ≤600€)")
        st.dataframe(t2 if not t2.empty else pd.DataFrame(), use_container_width=True)

        st.subheader("Tier-3 Strict (Same Date, ≥85%)")
        st.dataframe(t3 if not t3.empty else pd.DataFrame(), use_container_width=True)

        st.subheader("Missing in ERP")
        st.dataframe(missV, use_container_width=True)

        st.subheader("Missing in Vendor")
        st.dataframe(missE, use_container_width=True)

    # ---------- PAYMENTS ----------
    with tab3:
        totE = pE["__amt"].sum() if not pE.empty else 0.0
        totV = pV["__amt"].sum() if not pV.empty else 0.0
        p1, p2, p3 = st.columns(3)
        p1.markdown(
            f"<div class='metric-box green'>Matched<br><h2>{len(payM)}</h2></div>",
            unsafe_allow_html=True,
        )
        p2.markdown(
            f"<div class='metric-box teal'>ERP Total<br><h2>{totE:.2f}</h2></div>",
            unsafe_allow_html=True,
        )
        p3.markdown(
            f"<div class='metric-box purple'>Ven Total<br><h2>{totV:.2f}</h2></div>",
            unsafe_allow_html=True,
        )

        st.subheader("Matched Payments")
        st.dataframe(payM if not payM.empty else pd.DataFrame(), use_container_width=True)

        st.subheader("Unmatched ERP Payments")
        st.dataframe(missPE, use_container_width=True)

        st.subheader("Unmatched Vendor Payments")
        st.dataframe(missPV, use_container_width=True)

        # ---- EXPORT UNMATCHED ONLY ----
        wb = Workbook()
        header_fill = PatternFill(
            start_color="ADD8E6", end_color="ADD8E6", fill_type="solid"
        )
        sheets = {
            "Unmatched_ERP": missE,
            "Unmatched_Vendor": missV,
            "Unmatched_ERP_Pay": missPE,
            "Unmatched_Ven_Pay": missPV,
        }
        for name, df in sheets.items():
            if df.empty:
                continue
            ws = wb.create_sheet(name)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(bold=True)
            total = df["Amount"].sum()
            ws.append(["Total", total] + [""] * (ws.max_column - 2))
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        out = BytesIO()
        wb.save(out)
        st.download_button(
            "Export Unmatched Only",
            data=out.getvalue(),
            file_name="unmatched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
