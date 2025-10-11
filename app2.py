def match_invoices(erp_df, ven_df):
    import re
    from fuzzywuzzy import fuzz

    # === Safe reset of indexes ===
    erp_df = erp_df.reset_index(drop=True).reset_index().rename(columns={"index": "_id_erp"})
    ven_df = ven_df.reset_index(drop=True).reset_index().rename(columns={"index": "_id_ven"})

    # === Guard: must have columns ===
    for df, name in [(erp_df, "ERP"), (ven_df, "Vendor")]:
        if not any("invoice" in c.lower() for c in df.columns):
            st.error(f"❌ Missing invoice column in {name} file")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # === Step 1: Identify TRN to reconcile ===
    selected_trn = None
    if "trn_ven" in ven_df.columns and not ven_df["trn_ven"].dropna().empty:
        selected_trn = str(ven_df["trn_ven"].dropna().iloc[0]).strip()

    if selected_trn:
        erp_df = erp_df[erp_df.get("trn_erp", "").astype(str).str.strip() == selected_trn]
        ven_df = ven_df[ven_df.get("trn_ven", "").astype(str).str.strip() == selected_trn]
    else:
        st.warning("⚠️ No TRN detected — reconciling entire dataset.")

    # === Step 2: Ignore payments ===
    skip_words = ["pago", "transferencia", "payment", "paid", "bank", "deposit", "wire", "transf", "πληρωμή"]
    if "description_ven" in ven_df.columns:
        ven_df = ven_df[
            ~ven_df["description_ven"].astype(str).str.lower().apply(
                lambda x: any(w in x for w in skip_words)
            )
        ].reset_index(drop=True)

    # === Step 3: Helpers ===
    def normalize_number(v):
        if pd.isna(v):
            return 0.0
        s = re.sub(r"[^\d,.\-]", "", str(v))
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except:
            return 0.0

    def is_cn_vendor(row):
        desc = str(row.get("description_ven", "")).lower()
        return "abono" in desc or "credit" in desc or normalize_number(row.get("credit_ven", 0)) > 0

    def is_cn_erp(row):
        amt = normalize_number(row.get("amount_erp", row.get("credit_erp", 0)))
        return amt < 0

    def vendor_amount(row):
        if is_cn_vendor(row):
            return abs(normalize_number(row.get("credit_ven", 0)))
        return abs(normalize_number(row.get("debit_ven", 0)))

    def erp_amount(row):
        return abs(normalize_number(row.get("amount_erp", row.get("credit_erp", 0))))

    def last_digits(s, k=6):
        s = "".join(re.findall(r"\d+", str(s)))
        return s[-k:] if s else ""

    def invoice_match(a, b):
        ta, tb = last_digits(a), last_digits(b)
        if ta and tb:
            for n in (6, 5, 4, 3):
                if ta[-n:] == tb[-n:]:
                    return True
        return fuzz.ratio(str(a), str(b)) >= 90

    # === Step 4: Separate invoice / CN pools ===
    try:
        erp_inv = erp_df[~erp_df.apply(is_cn_erp, axis=1)].copy()
        erp_cn = erp_df[erp_df.apply(is_cn_erp, axis=1)].copy()
        ven_inv = ven_df[~ven_df.apply(is_cn_vendor, axis=1)].copy()
        ven_cn = ven_df[ven_df.apply(is_cn_vendor, axis=1)].copy()
    except Exception:
        st.error("❌ Error while classifying CN vs Invoice. Check file columns.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # === Step 5: Matching logic ===
    def perform_match(erp_pool, ven_pool, trn):
        matches = []
        used_e, used_v = set(), set()

        for _, e in erp_pool.iterrows():
            e_id, e_inv = e["_id_erp"], str(e.get("invoice_erp", "")).strip()
            e_amt = erp_amount(e)

            for _, v in ven_pool.iterrows():
                v_id, v_inv = v["_id_ven"], str(v.get("invoice_ven", "")).strip()
                v_amt = vendor_amount(v)
                if invoice_match(e_inv, v_inv):
                    diff = round(e_amt - v_amt, 2)
                    matches.append({
                        "Vendor/Supplier": e.get("vendor_erp", ""),
                        "TRN/AFM": trn or "N/A",
                        "ERP Invoice": e_inv,
                        "Vendor Invoice": v_inv,
                        "ERP Amount": e_amt,
                        "Vendor Amount": v_amt,
                        "Difference": diff,
                        "Status": "Match" if abs(diff) <= 0.01 else "Difference",
                        "Description": str(v.get("description_ven", "")),
                    })
                    used_e.add(e_id)
                    used_v.add(v_id)
                    break

        matched_erp = erp_pool.loc[erp_pool["_id_erp"].isin(used_e)]
        matched_ven = ven_pool.loc[ven_pool["_id_ven"].isin(used_v)]
        return matches, matched_erp, matched_ven

    # === Step 6: Perform reconciliation ===
    inv_matches, inv_e, inv_v = perform_match(erp_inv, ven_inv, selected_trn)
    cn_matches, cn_e, cn_v = perform_match(erp_cn, ven_cn, selected_trn)

    matched = pd.DataFrame(inv_matches + cn_matches)

    # === Step 7: Identify missing ===
    erp_used = pd.concat([inv_e, cn_e], ignore_index=True)
    ven_used = pd.concat([inv_v, cn_v], ignore_index=True)

    erp_missing = erp_df[~erp_df["_id_erp"].isin(erp_used["_id_erp"])].reset_index(drop=True)
    ven_missing = ven_df[~ven_df["_id_ven"].isin(ven_used["_id_ven"])].reset_index(drop=True)

    return matched, erp_missing, ven_missing
