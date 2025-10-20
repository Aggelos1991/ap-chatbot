import streamlit as st
import pandas as pd
import re


# ======================================
# CONFIGURATION
# ======================================

st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ======================================
# HELPERS
# ======================================
def normalize_number(v):
    """Convert numeric strings like '1.234,56' or '1,234.56' safely to float."""
    if v is None or str(v).strip() == "":
        return 0.0
    s = str(v).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
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


def normalize_columns(df, tag):
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura", "alternative document",
            "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber",
            "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto",
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion",
            "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            "ημερομηνία", "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού"
        ],
    }

    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"

    out = df.rename(columns=rename_map)
    for required in ["debit", "credit"]:
        cname = f"{required}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0
    return out


# ======================================
# CORE MATCHING
# ======================================
def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    # === detect doc type ===
    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        payment_patterns = [
            r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer",
            r"^trf", r"^remesa", r"^pago", r"^transferencia", r"(?i)^f[-\s]?\d{4,8}",
        ]
        if any(re.search(p, reason) for p in payment_patterns):
            return "IGNORE"
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση","ακυρωτικό","ακυρωτικό παραστατικό"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in credit_words):
            return "CN"
        elif any(k in reason for k in invoice_words) or credit > 0:
            return "INV"
        return "UNKNOWN"

    def detect_vendor_doc_type(row):
        reason = str(row.get("reason_ven", "")).lower()
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        payment_words = [
            "pago", "payment", "transfer", "bank", "saldo", "trf",
            "πληρωμή", "μεταφορά", "τράπεζα", "τραπεζικό έμβασμα"
        ]
        credit_words = [
            "credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση","ακυρωτικό","ακυρωτικό παραστατικό"
        ]
        invoice_words = [
            "factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"
        ]
        if any(k in reason for k in payment_words):
            return "IGNORE"
        elif any(k in reason for k in credit_words) or credit > 0:
            return "CN"
        elif any(k in reason for k in invoice_words) or debit > 0:
            return "INV"
        return "UNKNOWN"

    def calc_erp_amount(row):
        doc = row.get("__doctype", "")
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        if doc == "INV":
            return abs(credit)
        elif doc == "CN":
            return -abs(charge if charge > 0 else credit)
        return 0.0

    def calc_vendor_amount(row):
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))
        doc = row.get("__doctype", "")
        if doc == "INV":
            return abs(debit)
        elif doc == "CN":
            return -abs(credit if credit > 0 else debit)
        return 0.0

    erp_df["__doctype"] = erp_df.apply(detect_erp_doc_type, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    # === remove INV + CN = 0 ===
    def remove_neutralized(df, inv_col):
        if inv_col not in df.columns or "__amt" not in df.columns:
            return df
        df["__amt"] = pd.to_numeric(df["__amt"], errors="coerce").fillna(0)
        grouped = df.groupby(inv_col)["__amt"].sum().reset_index()
        neutralized = grouped[abs(grouped["__amt"]) < 0.05][inv_col].dropna().astype(str).tolist()
        return df[~df[inv_col].astype(str).isin(neutralized)].reset_index(drop=True)

    erp_use = remove_neutralized(erp_use, "invoice_erp")
    ven_use = remove_neutralized(ven_use, "invoice_ven")

    # === match ===
    def clean_invoice_code(v):
        if not v:
            return ""
        s = str(v).strip().lower()
        s = re.sub(r"[^a-z0-9]", "", s)
        s = re.sub(r"^0+", "", s)
        return s

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_code = clean_invoice_code(e_inv)
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_code = clean_invoice_code(v_inv)
            diff = round(e_amt - v_amt, 2)
            if e["__doctype"] == v["__doctype"] and (e_code == v_code) and abs(diff) < 0.05:
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Match"
                })
                used_vendor_rows.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    matched_erp = {m["ERP Invoice"] for _, m in matched_df.iterrows()}
    matched_ven = {m["Vendor Invoice"] for _, m in matched_df.iterrows()}

    missing_in_erp = ven_use[~ven_use["invoice_ven"].isin(matched_ven)][["invoice_ven", "__amt"]]
    missing_in_vendor = erp_use[~erp_use["invoice_erp"].isin(matched_erp)][["invoice_erp", "__amt"]]

    missing_in_erp = missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})

    # 🧹 Final neutralization cleanup
    def remove_residual_neutralized(df):
        if df.empty:
            return df
        df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
        grouped = df.groupby("Invoice")["Amount"].sum().reset_index()
        drop_list = grouped[abs(grouped["Amount"]) < 0.05]["Invoice"].tolist()
        return df[~df["Invoice"].isin(drop_list)].reset_index(drop=True)

    missing_in_erp = remove_residual_neutralized(missing_in_erp)
    missing_in_vendor = remove_residual_neutralized(missing_in_vendor)

    return matched_df, missing_in_erp, missing_in_vendor


# ======================================
# PAYMENTS
# ======================================
def extract_payments(erp_df, ven_df):
    payment_keywords = [
        "πληρωμή", "payment", "bank transfer", "transferencia bancaria",
        "transfer", "trf", "remesa", "pago", "deposit", "μεταφορά", "έμβασμα"
    ]
    exclude_keywords = [
        "τιμολόγιο", "invoice", "παραστατικό", "έξοδα", "expenses", "expense",
        "invoice of expenses", "expense invoice", "τιμολόγιο εξόδων",
        "διόρθωση", "διορθώσεις", "correction", "reclass", "adjustment",
        "μεταφορά υπολοίπου", "balance transfer"
    ]

    def is_real_payment(reason):
        text = str(reason or "").lower()
        return any(k in text for k in payment_keywords) and not any(b in text for b in exclude_keywords)

    erp_pay = erp_df[erp_df["reason_erp"].apply(is_real_payment)].copy() if "reason_erp" in erp_df else pd.DataFrame()
    ven_pay = ven_df[ven_df["reason_ven"].apply(is_real_payment)].copy() if "reason_ven" in ven_df else pd.DataFrame()

    for df, tag in [(erp_pay, "erp"), (ven_pay, "ven")]:
        if not df.empty:
            df["Amount"] = df.apply(
                lambda r: abs(normalize_number(r.get(f"debit_{tag}")) - normalize_number(r.get(f"credit_{tag}"))),
                axis=1
            )

    matched_pay = []
    used_vendor = set()
    for _, e in erp_pay.iterrows():
        for v_idx, v in ven_pay.iterrows():
            if v_idx in used_vendor:
                continue
            if abs(e["Amount"] - v["Amount"]) < 0.05:
                matched_pay.append({
                    "ERP Reason": e.get("reason_erp", ""),
                    "Vendor Reason": v.get("reason_ven", ""),
                    "ERP Amount": round(e["Amount"], 2),
                    "Vendor Amount": round(v["Amount"], 2),
                    "Difference": round(abs(e["Amount"] - v["Amount"]), 2)
                })
                used_vendor.add(v_idx)
                break

    return erp_pay, ven_pay, pd.DataFrame(matched_pay)


# ======================================
# STREAMLIT UI
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)
    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
        erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)
    st.success("✅ Reconciliation complete")

    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        return [''] * len(row)

    st.subheader("📊 Matched / Differences")
    st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)

    st.subheader("❌ Missing in ERP")
    st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)

    st.subheader("❌ Missing in Vendor")
    st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)

    st.subheader("🏦 Payment Transactions")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**💼 ERP Payments**")
        st.dataframe(erp_pay, use_container_width=True)
    with col2:
        st.markdown("**🧾 Vendor Payments**")
        st.dataframe(ven_pay, use_container_width=True)
    st.markdown("### ✅ Matched Payments")
    st.dataframe(matched_pay, use_container_width=True)


    # === Excel Export ===
    def export_reconciliation_excel(matched, erp_missing, ven_missing):
        import io
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            matched.to_excel(writer, index=False, sheet_name="Matched & Differences")
            erp_missing.to_excel(writer, index=False, sheet_name="Missing ERP")
            ven_missing.to_excel(writer, index=False, sheet_name="Missing Vendor")
        output.seek(0)
        return output

    st.markdown("### 📥 Download Reconciliation Excel Report")
    excel_output = export_reconciliation_excel(matched, erp_missing, ven_missing)
    st.download_button(
        label="⬇️ Download Excel Report",
        data=excel_output,
        file_name="Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
