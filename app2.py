def match_invoices(erp_df, ven_df):
    try:
        # Reset indexes for clarity
        erp_df = erp_df.reset_index().rename(columns={"index": "_id_erp"})
        ven_df = ven_df.reset_index().rename(columns={"index": "_id_ven"})

        matched, matched_erp, matched_ven = [], set(), set()

        # Detect vendor TRN for filtering
        trn_col = next((c for c in ven_df.columns if "trn_" in c), None)
        if trn_col:
            vendor_trn = str(ven_df[trn_col].dropna().iloc[0])
            erp_df = erp_df[erp_df.get("trn_erp", "") == vendor_trn]

        # 🧹 STEP 1: Remove payment lines on BOTH sides
        payment_keywords = [
            "pago parcial", "pago", "payment", "transfer", 
            "liquidación", "transferencia", "abono parcial", 
            "bank", "partial payment", "pago fraccionado"
        ]

        def is_payment(desc):
            return any(k in str(desc).lower() for k in payment_keywords)

        if "description_erp" in erp_df.columns:
            erp_df = erp_df[~erp_df["description_erp"].apply(is_payment)]
        if "description_ven" in ven_df.columns:
            ven_df = ven_df[~ven_df["description_ven"].apply(is_payment)]

        # 🧮 STEP 2: Matching process
        for _, e_row in erp_df.iterrows():
            e_inv = str(e_row.get("invoice_erp", "")).strip()
            e_amt = normalize_number(e_row.get("amount_erp", 0))
            if not e_inv:
                continue

            for _, v_row in ven_df.iterrows():
                v_inv = str(v_row.get("invoice_ven", "")).strip()
                v_desc = str(v_row.get("description_ven", "")).lower()

                # Determine debit/credit value
                d_val = normalize_number(v_row.get("debit_ven", 0))
                c_val = normalize_number(v_row.get("credit_ven", 0))
                v_amt = c_val if "abono" in v_desc or "credit" in v_desc else d_val

                # Fuzzy invoice match (last 4 digits or high ratio)
                if e_inv[-4:] in v_inv or fuzz.ratio(e_inv, v_inv) > 80:
                    diff = round(e_amt - v_amt, 2)
                    status = "Match" if abs(diff) < 0.05 else "Difference"

                    matched.append({
                        "Vendor/Supplier": e_row.get("vendor_erp", ""),
                        "TRN/AFM": e_row.get("trn_erp", ""),
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
        matched_vendor_invoices = df_matched["Vendor Invoice"].unique().tolist()

        # 🧾 STEP 3: Identify missing invoices
        erp_missing = erp_df[
            ~erp_df["invoice_erp"].astype(str).isin(matched_invoices)
        ].reset_index(drop=True)
        ven_missing = ven_df[
            ~ven_df["invoice_ven"].astype(str).isin(matched_vendor_invoices)
        ].reset_index(drop=True)

        return df_matched, erp_missing, ven_missing

    except Exception as e:
        st.error(f"❌ Matching error: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
