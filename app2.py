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
            # Greek
            "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber",
            # Greek
            "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto",
            # Greek
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion",
            # Greek
            "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            # Greek (safe only)
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            # Greek
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

    # Ensure debit/credit exist
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

    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))

        # Unified multilingual keywords/patterns
        payment_patterns = [
            r"^πληρωμ",             # Greek "Πληρωμή"
            r"^απόδειξη\s*πληρωμ",  # Greek "Απόδειξη πληρωμής"
            r"^payment",            # English: "Payment"
            r"^bank\s*transfer",    # "Bank Transfer"
            r"^trf",                # "TRF ..."
            r"^remesa",             # Spanish
            r"^pago",               # Spanish
            r"^transferencia",      # Spanish
            r"(?i)^f[-\s]?\d{4,8}",
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

    def calc_erp_amount(row):
        doc = row.get("__doctype", "")
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        if doc == "INV":
            return abs(credit)
        elif doc == "CN":
            return -abs(charge if charge > 0 else credit)
        return 0.0

    def detect_vendor_doc_type(row):
        reason = str(row.get("reason_ven", "")).lower()
        debit = normalize_number(row.get("debit_ven"))
        credit = normalize_number(row.get("credit_ven"))

        # Unified multilingual keywords
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
    erp_df["__amt"] = erp_df.apply(calc_erp_amount, axis=1)
    ven_df["__doctype"] = ven_df.apply(detect_vendor_doc_type, axis=1)
    ven_df["__amt"] = ven_df.apply(calc_vendor_amount, axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    # ====== SCENARIO 1 & 2: MERGE MULTIPLE AND CREDIT NOTES ======
    merged_rows = []
    for inv, group in erp_use.groupby("invoice_erp", dropna=False):
        if group.empty:
            continue

        # If 3 or more entries → take the last (latest)
        if len(group) >= 3:
            group = group.tail(1)

        # If both INV and CN exist for same number → combine
        inv_rows = group[group["__doctype"] == "INV"]
        cn_rows = group[group["__doctype"] == "CN"]

        if not inv_rows.empty and not cn_rows.empty:
            total_inv = inv_rows["__amt"].sum()
            total_cn = cn_rows["__amt"].sum()
            net = round(total_inv + total_cn, 2)
            base_row = inv_rows.iloc[-1].copy()
            base_row["__amt"] = net
            merged_rows.append(base_row)
        else:
            merged_rows.append(group.iloc[-1])

    erp_use = pd.DataFrame(merged_rows).reset_index(drop=True)

    def extract_digits(v):
        return re.sub(r"\D", "", str(v or "")).lstrip("0")

    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    # Add missing cleaner so we can compute e_code / v_code
    def clean_invoice_code(v):
        """
        Normalize invoice strings for comparison:
        - drop common prefixes
        - remove year snippets (20xx)
        - strip non-alphanumerics
        - keep only digits and trim leading zeros
        """
        if not v:
            return ""
        s = str(v).strip().lower()
            # 🧩 Handle structured invoice patterns like 2025-FV-00001-001248-01
        parts = re.split(r"[-_]", s)
        for p in reversed(parts):
            # numeric block with ≥4 digits, skip if it's a year (2020–2039)
            if re.fullmatch(r"\d{4,}", p) and not re.fullmatch(r"20[0-3]\d", p):
                s = p.lstrip("0")  # trim leading zeros (001248 → 1248)
                break
        s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no)\W*", "", s)
        s = re.sub(r"20\d{2}", "", s)
        s = re.sub(r"[^a-z0-9]", "", s)
        s = re.sub(r"^0+", "", s)
        # keep only digits for the final compare (like earlier logic)
        s = re.sub(r"[^\d]", "", s)
        s = re.sub(
    r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα|apo|ref|doc|num|no|fa|sf|ba|vn)\W*", 
    "", 
    s
)            
        return s
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_digits = extract_digits(e_inv)
        e_code = clean_invoice_code(e_inv)  # <<< compute cleaned code

        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_digits = extract_digits(v_inv)
            v_code = clean_invoice_code(v_inv)  # <<< compute cleaned code
            diff = round(e_amt - v_amt, 2)
            amt_close = abs(diff) < 0.05

            # --- Υποψήφιοι έλεγχοι ομοιότητας ---
            same_full  = (e_inv == v_inv)
            same_clean = (e_code == v_code)

            len_diff = abs(len(e_code) - len(v_code))
            suffix_ok = (
                len(e_code) > 2 and len(v_code) > 2 and
                len_diff <= 2 and (
                    e_code.endswith(v_code) or v_code.endswith(e_code)
                )
            )

            same_type = (e["__doctype"] == v["__doctype"])

            # --- ΝΕΟΣ κανόνας αποδοχής ---
            if same_type and same_full:
                take_it = True
            elif same_type and (same_clean or suffix_ok) and amt_close:
                take_it = True
            else:
                take_it = False

            if take_it:
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": "Match" if amt_close else "Difference"
                })
                used_vendor_rows.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)
    matched_erp = {m["ERP Invoice"] for _, m in matched_df.iterrows()}
    matched_ven = {m["Vendor Invoice"] for _, m in matched_df.iterrows()}

    missing_in_erp = (
        ven_use[~ven_use["invoice_ven"].isin(matched_ven)][["invoice_ven", "__amt"]]
        if "invoice_ven" in ven_use else pd.DataFrame()
    )
    missing_in_vendor = (
        erp_use[~erp_use["invoice_erp"].isin(matched_erp)][["invoice_erp", "__amt"]]
        if "invoice_erp" in erp_use else pd.DataFrame()
    )

    missing_in_erp = missing_in_erp.rename(columns={"invoice_ven": "Invoice", "__amt": "Amount"})
    missing_in_vendor = missing_in_vendor.rename(columns={"invoice_erp": "Invoice", "__amt": "Amount"})

    return matched_df, missing_in_erp, missing_in_vendor


# ======================================
def extract_payments(erp_df: pd.DataFrame, ven_df: pd.DataFrame):
    # --- λέξεις που δείχνουν πληρωμή ---
    payment_keywords = [
        "πληρωμή", "payment", "bank transfer", "transferencia bancaria",
        "transfer", "trf", "remesa", "pago", "deposit", "μεταφορά", "έμβασμα"
    ]

    # --- λέξεις που δείχνουν ότι ΔΕΝ είναι πληρωμή ---
    exclude_keywords = [
        "τιμολόγιο", "invoice", "παραστατικό", "έξοδα", "expenses", "expense",
        "invoice of expenses", "expense invoice", "τιμολόγιο εξόδων",
        "διόρθωση", "διορθώσεις", "correction", "reclass", "adjustment",
        "μεταφορά υπολοίπου", "balance transfer"
    ]

    def is_real_payment(reason: str) -> bool:
        """Επιστρέφει True μόνο αν περιέχει λέξη πληρωμής και δεν περιέχει καμία εξαιρούμενη λέξη."""
        text = str(reason or "").lower()
        has_payment = any(k in text for k in payment_keywords)
        has_exclusion = any(bad in text for bad in exclude_keywords)
        return has_payment and not has_exclusion

    # --- Φιλτράρισμα ERP & Vendor ---
    erp_pay = (
        erp_df[erp_df["reason_erp"].apply(is_real_payment)].copy()
        if "reason_erp" in erp_df else pd.DataFrame()
    )
    ven_pay = (
        ven_df[ven_df["reason_ven"].apply(is_real_payment)].copy()
        if "reason_ven" in ven_df else pd.DataFrame()
    )

    # --- Υπολογισμός ποσών ---
    if not erp_pay.empty:
        erp_pay["Amount"] = erp_pay.apply(
            lambda r: abs(normalize_number(r.get("debit_erp")) - normalize_number(r.get("credit_erp"))),
            axis=1
        )
    if not ven_pay.empty:
        ven_pay["Amount"] = ven_pay.apply(
            lambda r: abs(normalize_number(r.get("debit_ven")) - normalize_number(r.get("credit_ven"))),
            axis=1
        )

    # --- Matching μεταξύ ERP & Vendor ---
    matched_payments = []
    used_vendor = set()
    for _, e in erp_pay.iterrows():
        for v_idx, v in ven_pay.iterrows():
            if v_idx in used_vendor:
                continue
            diff = abs(e["Amount"] - v["Amount"])
            if diff < 0.05:
                matched_payments.append({
                    "ERP Reason": e.get("reason_erp", ""),
                    "Vendor Reason": v.get("reason_ven", ""),
                    "ERP Amount": round(float(e["Amount"]), 2),
                    "Vendor Amount": round(float(v["Amount"]), 2),
                    "Difference": round(diff, 2)
                })
                used_vendor.add(v_idx)
                break

    return erp_pay, ven_pay, pd.DataFrame(matched_payments)

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

    # ====== HIGHLIGHTING ======
    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        return [''] * len(row)

    # ====== MATCHED ======
    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches found.")

    # ====== MISSING ======
    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    if not erp_missing.empty:
        st.dataframe(
            erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
            use_container_width=True
        )
    else:
        st.success("✅ No missing invoices in ERP.")

    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    if not ven_missing.empty:
        st.dataframe(
            ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"),
            use_container_width=True
        )
    else:
        st.success("✅ No missing invoices in Vendor.")

    # ====== PAYMENTS ======
    st.subheader("🏦 Payment Transactions (Identified in both sides)")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**💼 ERP Payments**")
        if not erp_pay.empty:
            st.dataframe(
                erp_pay.style.applymap(lambda _: "background-color: #004d40; color: white"),
                use_container_width=True
            )
            st.markdown(f"**Total ERP Payments:** {erp_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No ERP payments found.")

    with col2:
        st.markdown("**🧾 Vendor Payments**")
        if not ven_pay.empty:
            st.dataframe(
                ven_pay.style.applymap(lambda _: "background-color: #1565c0; color: white"),
                use_container_width=True
            )
            st.markdown(f"**Total Vendor Payments:** {ven_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No Vendor payments found.")

    st.markdown("### ✅ Matched Payments")
    if not matched_pay.empty:
        st.dataframe(
            matched_pay.style.applymap(lambda _: "background-color: #2e7d32; color: white"),
            use_container_width=True
        )
        total_erp = matched_pay["ERP Amount"].sum()
        total_vendor = matched_pay["Vendor Amount"].sum()
        diff_total = abs(total_erp - total_vendor)
        st.markdown(f"**Total Matched ERP Payments:** {total_erp:,.2f} EUR")
        st.markdown(f"**Total Matched Vendor Payments:** {total_vendor:,.2f} EUR")
        st.markdown(f"**Difference Between ERP and Vendor Payments:** {diff_total:,.2f} EUR")
    else:
        st.info("No matching payments found.")


    # ======================================
    # 📥 EXCEL EXPORT (2-SHEET VERSION)
    # ======================================

    # ======================================
     

    # ======================================
    # 🦖 EXCEL EXPORT (FANCY 2-SHEET VERSION)
    # ======================================
    import io
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    def export_reconciliation_excel(matched, erp_missing, ven_missing):
        """Generate a visually enhanced Excel with colors and formatting."""
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # 1️⃣ Matched & Differences sheet
            matched.to_excel(writer, index=False, sheet_name="Matched & Differences")
            ws1 = writer.sheets["Matched & Differences"]

            # Apply header style (green)
            header_fill = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            for cell in ws1[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Alternate row shading
            for i, row in enumerate(ws1.iter_rows(min_row=2), start=2):
                if i % 2 == 0:
                    for cell in row:
                        cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

            # Auto column width
            for col in ws1.columns:
                max_length = 0
                column = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws1.column_dimensions[column].width = max_length + 2

            # 2️⃣ Missing sheet
            erp_missing.to_excel(writer, index=False, sheet_name="Missing", startrow=3)
            start_col = len(erp_missing.columns) + 3
            ven_missing.to_excel(writer, index=False, sheet_name="Missing", startcol=start_col, startrow=3)

            ws2 = writer.sheets["Missing"]
            ws2["A1"] = "Missing in ERP"
            ws2["A1"].font = Font(bold=True, size=14, color="FFFFFF")
            ws2["A1"].alignment = Alignment(horizontal="center")
            ws2["A1"].fill = PatternFill(start_color="E53935", end_color="E53935", fill_type="solid")

            ws2.cell(row=1, column=start_col + 1).value = "Missing in Vendor"
            ws2.cell(row=1, column=start_col + 1).font = Font(bold=True, size=14, color="FFFFFF")
            ws2.cell(row=1, column=start_col + 1).alignment = Alignment(horizontal="center")
            ws2.cell(row=1, column=start_col + 1).fill = PatternFill(start_color="1E88E5", end_color="1E88E5", fill_type="solid")

            # Format headers for both tables
            header_fill_erp = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
            header_fill_ven = PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid")
            for cell in ws2[3]:
                cell.fill = header_fill_erp
                cell.font = Font(bold=True)
            for cell in ws2.iter_rows(min_row=3, min_col=start_col + 1, max_row=3):
                for c in cell:
                    c.fill = header_fill_ven
                    c.font = Font(bold=True)

            # Add totals
            erp_total = erp_missing["Amount"].sum() if not erp_missing.empty else 0
            ven_total = ven_missing["Amount"].sum() if not ven_missing.empty else 0
            erp_total_row = erp_missing.shape[0] + 4
            ven_total_col = start_col + 1

            ws2[f"A{erp_total_row}"] = "TOTAL:"
            ws2[f"A{erp_total_row}"].font = Font(bold=True, color="E53935")
            ws2[f"B{erp_total_row}"] = float(erp_total)
            ws2[f"B{erp_total_row}"].font = Font(bold=True)

            ws2.cell(row=erp_total_row, column=ven_total_col).value = "TOTAL:"
            ws2.cell(row=erp_total_row, column=ven_total_col).font = Font(bold=True, color="1E88E5")
            ws2.cell(row=erp_total_row, column=ven_total_col + 1).value = float(ven_total)
            ws2.cell(row=erp_total_row, column=ven_total_col + 1).font = Font(bold=True)

            # Borders and auto width
            thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                          top=Side(style="thin"), bottom=Side(style="thin"))
            for row in ws2.iter_rows(min_row=3):
                for cell in row:
                    cell.border = thin

            for col in ws2.columns:
                max_length = 0
                column = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws2.column_dimensions[column].width = max_length + 2

        output.seek(0)
        return output

    # ====== DOWNLOAD BUTTON ======
    st.markdown("### 📥 Download Excel Report")
    excel_output = export_reconciliation_excel(matched, erp_missing, ven_missing)
    st.download_button(
        label="⬇️ Download Reconciliation Report (Excel)",
        data=excel_output,
        file_name="Reconciliation_Report_Fancy.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
