def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    def detect_erp_doc_type(row):
        reason = str(row.get("reason_erp", "")).lower()
        charge = normalize_number(row.get("debit_erp"))
        credit = normalize_number(row.get("credit_erp"))
        payment_words = ["pago", "payment", "transfer", "bank", "saldo", "trf",
                         "πληρωμή", "μεταφορά", "τράπεζα", "τραπεζικό έμβασμα"]
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
        if any(k in reason for k in payment_words):
            return "IGNORE"
        elif any(k in reason for k in credit_words):
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
        payment_words = ["pago", "payment", "transfer", "bank", "saldo", "trf",
                         "πληρωμή", "μεταφορά", "τράπεζα", "τραπεζικό έμβασμα"]
        credit_words = ["credit", "nota", "abono", "cn", "πιστωτικό", "πίστωση"]
        invoice_words = ["factura", "invoice", "inv", "τιμολόγιο", "παραστατικό"]
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
        if len(group) >= 3:
            group = group.tail(1)
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

    # ====== NEW SMART MATCHING LOGIC ======
    def clean_invoice_code(v):
        """Cleans invoice numbers by removing prefixes, years, special chars, and leading zeros."""
        if not v:
            return ""
        s = str(v).strip().lower()

        # Remove common Greek + Latin prefixes (real AP patterns)
        s = re.sub(r"^(αρ|τιμ|pf|ab|inv|tim|cn|ar|pa|πφ|πα)\W*", "", s)

        # Remove any year pattern (start or end)
        s = re.sub(r"(^20\d{2}[\W/\\-]*)|([\W/\\-]*20\d{2}$)", "", s)

        # Remove all special chars
        s = re.sub(r"[^a-z0-9]", "", s)

        # Remove leading zeros
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
            amt_close = abs(diff) < 0.05

            same_full = e_inv == v_inv
            same_clean = e_code == v_code
            length_diff = abs(len(e_code) - len(v_code))
            partial_match = (
                len(e_code) > 2 and len(v_code) > 2 and
                length_diff <= 2 and (
                    e_code.endswith(v_code) or v_code.endswith(e_code)
                )
            )

            # Compute fuzzy similarity ratio between cleaned codes
            score = SequenceMatcher(None, e_code, v_code).ratio()

            if same_full or same_clean or partial_match or score > 0.85:
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
