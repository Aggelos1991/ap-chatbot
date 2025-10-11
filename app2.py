def match_invoices(erp_df, ven_df):
    import re

    def norm_inv(s: str) -> str:
        # normalize invoice number to strict comparable token
        return re.sub(r"[^A-Za-z0-9]", "", str(s or "")).upper()

    def is_payment(desc: str) -> bool:
        w = str(desc).lower()
        return any(k in w for k in ["pago", "payment", "transfer", "transferencia", "liquidación", "partial"])

    def is_cn_vendor(row) -> bool:
        desc = str(row.get("description_ven", "")).lower()
        credit = normalize_number(row.get("credit_ven", 0))
        return ("abono" in desc or "credit" in desc) or credit > 0

    def is_cn_erp(row) -> bool:
        desc = str(row.get("description_erp", "")).lower()
        amt = normalize_number(row.get("amount_erp", row.get("credit_erp", 0)))
        return ("credit" in desc) or (amt < 0)

    def ven_amount(row) -> float:
        return abs(normalize_number(row.get("credit_ven", 0))) if is_cn_vendor(row) else abs(normalize_number(row.get("debit_ven", 0)))

    def erp_amount(row) -> float:
        return abs(normalize_number(row.get("amount_erp", row.get("credit_erp", 0))))

    # --- reset & TRN scoping ---
    erp_df = erp_df.reset_index(drop=True)
    ven_df = ven_df.reset_index(drop=True)

    trn_erp = next((c for c in erp_df.columns if "trn_" in c), None)
    trn_ven = next((c for c in ven_df.columns if "trn_" in c), None)
    if trn_erp and trn_ven and not ven_df[trn_ven].dropna().empty:
        trn = str(ven_df[trn_ven].dropna().iloc[0]).strip()
        erp_df = erp_df[erp_df[trn_erp].astype(str).str.strip() == trn].copy()
        ven_df = ven_df[ven_df[trn_ven].astype(str).str.strip() == trn].copy()

    # --- remove payments from BOTH sides ---
    if "description_erp" in erp_df.columns:
        erp_df = erp_df[~erp_df["description_erp"].apply(is_payment)].copy()
    if "description_ven" in ven_df.columns:
        ven_df = ven_df[~ven_df["description_ven"].apply(is_payment)].copy()

    # --- split into invoices vs credit notes ---
    erp_inv = erp_df[~erp_df.apply(is_cn_erp, axis=1)].copy()
    erp_cn  = erp_df[ erp_df.apply(is_cn_erp, axis=1)].copy()
    ven_inv = ven_df[~ven_df.apply(is_cn_vendor, axis=1)].copy()
    ven_cn  = ven_df[ ven_df.apply(is_cn_vendor, axis=1)].copy()

    # normalize invoice columns once
    for df, col in [(erp_inv, "invoice_erp"), (erp_cn, "invoice_erp"),
                    (ven_inv, "invoice_ven"), (ven_cn, "invoice_ven")]:
        if col in df.columns:
            df["_norm_inv"] = df[col].astype(str).map(norm_inv)
        else:
            df["_norm_inv"] = ""

    matched_rows = []

    def perform_match(pool_erp: pd.DataFrame, pool_ven: pd.DataFrame):
        used_ven = set()
        for _, e in pool_erp.iterrows():
            e_norm = e["_norm_inv"]
            if not e_norm:
                continue
            # exact-only match on normalized invoice number
            cand = pool_ven[(pool_ven["_norm_inv"] == e_norm) & (~pool_ven.index.isin(used_ven))]
            if cand.empty:
                continue
            v = cand.iloc[0]
            used_ven.add(v.name)

            e_amt = erp_amount(e)
            v_amt = ven_amount(v)
            diff = round(e_amt - v_amt, 2)
            status = "Match" if abs(diff) <= 0.01 else "Difference"

            matched_rows.append({
                "Vendor/Supplier": e.get("vendor_erp", ""),
                "TRN/AFM": e.get(trn_erp, ""),
                "ERP Invoice": e.get("invoice_erp", ""),
                "Vendor Invoice": v.get("invoice_ven", ""),
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status,
                "Description": str(v.get("description_ven", "")),
            })
        matched_ven_ids = set(x for x in used_ven)
        return matched_ven_ids

    used_inv = perform_match(erp_inv, ven_inv)
    used_cn  = perform_match(erp_cn,  ven_cn)

    matched_df = pd.DataFrame(matched_rows)

    # missing = anything not matched within its own type
    erp_used_invoices = set(matched_df["ERP Invoice"].astype(str)) if not matched_df.empty else set()
    ven_used_invoices = set(matched_df["Vendor Invoice"].astype(str)) if not matched_df.empty else set()

    erp_missing = erp_df[~erp_df["invoice_erp"].astype(str).isin(erp_used_invoices)].reset_index(drop=True)
    ven_missing = ven_df[~ven_df["invoice_ven"].astype(str).isin(ven_used_invoices)].reset_index(drop=True)

    return matched_df, erp_missing, ven_missing
