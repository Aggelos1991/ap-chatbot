def match_invoices(erp_df, ven_df):
    try:
        erp_df = erp_df.reset_index().rename(columns={"index": "_id_erp"})
        ven_df = ven_df.reset_index().rename(columns={"index": "_id_ven"})

        matched, matched_erp, matched_ven = [], set(), set()

        # Detect vendor TRN
        trn_col = next((c for c in ven_df.columns if "trn_" in c), None)
        if trn_col:
            vendor_trn = str(ven_df[trn_col].dropna().iloc[0])
            erp_df = erp_df[erp_df.get("trn_erp", "") == vendor_trn]

        # 🧹 STEP 1: Filter out payment / transfer records from BOTH sides
        payment_keywords = ["pago", "payment", "transfer", "bank", "liquidación", "abono parcial"]
        def is_payment(desc):
            return any(k in str(desc).lower() for k in payment_keywords)

        # Filter ERP
        if "description_erp" in erp_df.columns:
            erp_df = erp_df[~erp_df["description_erp"].apply(is_payment)]
        # Filter Vendor
        if "description_ven" in ven_df.columns:
            ven_df = ven_df[~ven_df["description_ven"].apply(is_payment)]

        # 🧮 STEP 2: Proceed with matching
        for _, e in erp_df.iterrows():
            e_inv = str(e.get("invoice_erp", "")).strip()
            e_amt = normalize_number(e.get("amount_erp", 0))
            if not e_inv:
                continue

            for _, v in ven_df.iterrows():
                v_inv = str(v.get("invoice_ven", "")).strip()
                v_desc = str(v.get("description_ven", "")).lower()

                # Determine vendor amount
                d_val = normalize_number(v.get("debit_ven", 0))
                c_val = normalize_number(v.get("credit_ven", 0))
                v_amt = c_val if "abono" in v_desc or "credit" in v_desc else d_val

                # Fuzzy invoice matching
                if e_inv[-4:] in v_inv or fuzz.ratio(e_inv, v_inv) > 80:
                    diff = round(e_amt - v_amt, 2)
                    status = "Match" if abs(diff) < 0.05 else "Difference"

                    matched.append({
                        "Vendor/Supplier": e.get("vendor_erp", ""),
                        "TRN/AFM": e.get("trn_erp", ""),
                        "ERP Invoice": e_inv,
                        "Vendor Invoice": v_inv,
                        "ERP Amount": e_amt,
                        "Vendor Amount": v_amt,
                        "Difference": diff,
                        "Status": status,
                        "Description": v_desc
                    })
                    matched_erp.add(e_inv)
                    matched_ven.add(v_inv)
                    break

        df_matched = pd.DataFrame(matched)
        matched_invoices = df_matched["ERP Invoice"].unique().tolist()
        matched_vendor = df_matched["Vendor Invoice"].unique().tolist()

        # 🧾 STEP 3: Identify only true missing invoices
        erp_missing = erp_df[~erp_df["invoice_erp"].astype(str).isin(matched_invoices)].reset_index(drop=True)
        ven_missing = ven_df[~ven_df["invoice_ven"].astype(str).isin(matched_vendor)].reset_index(drop=True)

        return df_matched, erp_missing, ven_missing

    except Exception as e:
        st.error(f"❌ Matching error: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
