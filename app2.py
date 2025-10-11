def match_invoices(erp_df, ven_df):
    import re

    # ---------- Helpers ----------
    def norm_inv(s):
        """Normalize invoice number to comparable format (A-Z0-9 only)."""
        return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()

    def safe_col(df, name, default=0):
        """Return column safely."""
        return df[name] if name in df.columns else [default] * len(df)

    def normalize_amount(v):
        try:
            return float(str(v).replace(",", "."))
        except:
            return 0.0

    def is_payment(desc):
        desc = str(desc).lower()
        return any(k in desc for k in ["pago", "payment", "transfer", "transferencia", "liquidación", "partial"])

    def is_cn_vendor(row):
        desc = str(row.get("description_ven", "")).lower()
        credit = normalize_amount(row.get("credit_ven", 0))
        return ("abono" in desc or "credit" in desc) or credit > 0

    def is_cn_erp(row):
        desc = str(row.get("description_erp", "")).lower()
        amt = normalize_amount(row.get("amount_erp", row.get("credit_erp", 0)))
        return ("credit" in desc) or (amt < 0)

    def ven_amount(row):
        return abs(normalize_amount(row.get("credit_ven", 0))) if is_cn_vendor(row) else abs(normalize_amount(row.get("debit_ven", 0)))

    def erp_amount(row):
        return abs(normalize_amount(row.get("amount_erp", row.get("credit_erp", 0))))

    # ---------- Prepare ----------
    erp_df = erp_df.copy().reset_index(drop=True)
    ven_df = ven_df.copy().reset_index(drop=True)

    # Find TRN columns
    trn_erp = next((c for c in erp_df.columns if "trn_" in c), None)
    trn_ven = next((c for c in ven_df.columns if "trn_" in c), None)

    # Scope to same vendor TRN if available
    if trn_erp and trn_ven and not ven_df[trn_ven].dropna().empty:
        vendor_trn = str(ven_df[trn_ven].dropna().iloc[0]).strip()
        erp_df = erp_df[erp_df[trn_erp].astype(str).str.strip() == vendor_trn].copy()
        ven_df = ven_df[ven_df[trn_ven].astype(str).str.strip() == vendor_trn].copy()

    # Remove payments
    if "description_erp" in erp_df.columns:
        erp_df = erp_df[~erp_df["description_erp"].apply(is_payment)].copy()
    if "description_ven" in ven_df.columns:
        ven_df = ven_df[~ven_df["description_ven"].apply(is_payment)].copy()

    # Guard empty uploads
    if erp_df.empty or ven_df.empty:
        return pd.DataFrame(), erp_df, ven_df

    # Split into invoices / CN
    erp_inv = erp_df[~erp_df.apply(is_cn_erp, axis=1)].copy()
    erp_cn  = erp_df[ erp_df.apply(is_cn_erp, axis=1)].copy()
    ven_inv = ven_df[~ven_df.apply(is_cn_vendor, axis=1)].copy()
    ven_cn  = ven_df[ ven_df.apply(is_cn_vendor, axis=1)].copy()

    # Add normalized invoice column
    for df, col in [(erp_inv, "invoice_erp"), (erp_cn, "invoice_erp"),
                    (ven_inv, "invoice_ven"), (ven_cn, "invoice_ven")]:
        df["_norm_inv"] = df[col].astype(str).map(norm_inv) if col in df.columns else ""

    matched_rows = []

    # ---------- Strict matching ----------
    def perform_match(erp_pool, ven_pool):
        used = set()
        for _, e in erp_pool.iterrows():
            e_norm = e.get("_norm_inv", "")
            if not e_norm:
                continue
            cand = ven_pool[(ven_pool["_norm_inv"] == e_norm) & (~ven_pool.index.isin(used))]
            if cand.empty:
                continue
            v = cand.iloc[0]
            used.add(v.name)
            e_amt, v_amt = erp_amount(e), ven_amount(v)
            diff = round(e_amt - v_amt, 2)
            status = "Match" if abs(diff) <= 0.05 else "Difference"
            matched_rows.append({
                "Vendor/Supplier": e.get("vendor_erp", ""),
                "TRN/AFM": e.get(trn_erp, ""),
                "ERP Invoice": e.get("invoice_erp", ""),
                "Vendor Invoice": v.get("invoice_ven", ""),
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status,
                "Description": v.get("description_ven", ""),
            })
        return used

    used_inv = perform_match(erp_inv, ven_inv)
    used_cn  = perform_match(erp_cn,  ven_cn)

    matched_df = pd.DataFrame(matched_rows)

    # ---------- Missing ----------
    erp_used = set(matched_df["ERP Invoice"].astype(str)) if not matched_df.empty else set()
    ven_used = set(matched_df["Vendor Invoice"].astype(str)) if not matched_df.empty else set()
    erp_missing = erp_df[~erp_df["invoice_erp"].astype(str).isin(erp_used)].reset_index(drop=True)
    ven_missing = ven_df[~ven_df["invoice_ven"].astype(str).isin(ven_used)].reset_index(drop=True)

    return matched_df, erp_missing, ven_missing
