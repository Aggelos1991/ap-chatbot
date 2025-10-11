def match_invoices(erp_df, ven_df):
    import re
    from fuzzywuzzy import fuzz

    erp_df = erp_df.reset_index().rename(columns={"index": "_id_erp"})
    ven_df = ven_df.reset_index().rename(columns={"index": "_id_ven"})

    # ---------- 1️⃣ Vendor TRN focus ----------
    if "trn_ven" not in ven_df.columns or ven_df["trn_ven"].dropna().empty:
        return pd.DataFrame([]), erp_df.iloc[0:0], ven_df.iloc[0:0]

    selected_trn = str(ven_df["trn_ven"].dropna().iloc[0]).strip()
    ven_df = ven_df[ven_df["trn_ven"].astype(str).str.strip() == selected_trn].copy()
    erp_df = erp_df[erp_df["trn_erp"].astype(str).str.strip() == selected_trn].copy()

    # ---------- 2️⃣ Ignore payments ----------
    skip_words = ["pago", "transferencia", "payment", "paid", "bank", "deposit", "wire", "transf", "πληρωμή"]
    if "description_ven" in ven_df.columns:
        ven_df = ven_df[
            ~ven_df["description_ven"].astype(str).str.lower().apply(
                lambda x: any(w in x for w in skip_words)
            )
        ].reset_index(drop=True)

    # ---------- Helpers ----------
    def is_cn_vendor(row):
        desc = str(row.get("description_ven", "")).lower()
        credit = normalize_number(row.get("credit_ven", 0))
        return ("abono" in desc or "credit" in desc) or credit > 0

    def vendor_amount(row):
        return normalize_number(row.get("credit_ven", 0)) if is_cn_vendor(row) else normalize_number(row.get("debit_ven", 0))

    def is_cn_erp(row):
        amt = normalize_number(row.get("amount_erp", row.get("credit_erp", 0)))
        return amt < 0

    def erp_amount(row):
        return abs(normalize_number(row.get("amount_erp", row.get("credit_erp", 0))))

    def last_digits(s, k=6):
        s = str(s)
        digits = re.findall(r"\d+", s)
        return "".join(digits)[-k:] if digits else ""

    def invoice_match(a, b):
        ta, tb = last_digits(a), last_digits(b)
        for n in (6, 5, 4, 3):
            if len(ta) >= n and len(tb) >= n and ta[-n:] == tb[-n:]:
                return True
        return fuzz.ratio(str(a), str(b)) >= 90

    # ---------- 3️⃣ Separate pools ----------
    erp_inv = erp_df[~erp_df.apply(is_cn_erp, axis=1)].copy()
    erp_cn = erp_df[erp_df.apply(is_cn_erp, axis=1)].copy()
    ven_inv = ven_df[~ven_df.apply(is_cn_vendor, axis=1)].copy()
    ven_cn = ven_df[ven_df.apply(is_cn_vendor, axis=1)].copy()

    # ---------- 4️⃣ Matching function ----------
    def perform_match(erp_pool, ven_pool, trn):
        matches = []
        used_e, used_v = set(), set()

        for i, e in erp_pool.iterrows():
            e_inv = str(e.get("invoice_erp", "")).strip()
            e_amt = abs(erp_amount(e))
            e_id = e["_id_erp"]

            for j, v in ven_pool.iterrows():
                if j in used_v:
                    continue
                v_inv = str(v.get("invoice_ven", "")).strip()
                v_amt = abs(vendor_amount(v))

                if not invoice_match(e_inv, v_inv):
                    continue

                diff = round(e_amt - v_amt, 2)
                status = "Match" if abs(diff) <= 0.01 else "Difference"

                matches.append({
                    "Vendor/Supplier": e.get("vendor_erp", ""),
                    "TRN/AFM": trn,
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": status,
                    "Description": str(v.get("description_ven", "")),
                })

                used_e.add(e_id)
                used_v.add(j)
                break  # move to next ERP invoice

        matched_erp = erp_pool.loc[erp_pool["_id_erp"].isin(used_e)]
        matched_ven = ven_pool.loc[ven_pool["_id_ven"].isin(used_v)]
        return matches, matched_erp, matched_ven

    # ---------- 5️⃣ Perform separate invoice and CN matches ----------
    inv_matches, inv_e, inv_v = perform_match(erp_inv, ven_inv, selected_trn)
    cn_matches, cn_e, cn_v = perform_match(erp_cn, ven_cn, selected_trn)

    matched = pd.DataFrame(inv_matches + cn_matches)

    # ---------- 6️⃣ Missing ----------
    erp_used = pd.concat([inv_e, cn_e], ignore_index=True)
    ven_used = pd.concat([inv_v, cn_v], ignore_index=True)
    erp_missing = erp_df[~erp_df["_id_erp"].isin(erp_used["_id_erp"])].reset_index(drop=True)
    ven_missing = ven_df[~ven_df["_id_ven"].isin(ven_used["_id_ven"])].reset_index(drop=True)

    return matched, erp_missing, ven_missing
