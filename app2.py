def match_invoices(erp_df, ven_df):
    try:
        # Reset index safely
        erp_df = erp_df.reset_index(drop=True)
        ven_df = ven_df.reset_index(drop=True)

        matched_rows = []
        matched_erp_invoices, matched_vendor_invoices = set(), set()

        # 🧾 Identify TRN/AFM if present
        trn_col_ven = next((c for c in ven_df.columns if "trn_" in c or "vat" in c or "cif" in c), None)
        trn_col_erp = next((c for c in erp_df.columns if "trn_" in c or "vat" in c), None)
        if trn_col_ven and trn_col_erp:
            common_trn = str(ven_df[trn_col_ven].dropna().iloc[0])
            erp_df = erp_df[erp_df[trn_col_erp].astype(str) == common_trn]

        # 🧹 STEP 1: Remove payment / transfer rows from both sides
        payment_keywords = [
            "pago", "payment", "transfer", "bank", "liquidación", "transferencia", "abono parcial", "partial"
        ]

        def filter_payments(df, desc_col):
            if desc_col in df.columns:
                return df[~df[desc_col].astype(str).str.lower().apply(
                    lambda x: any(k in x for k in payment_keywords)
                )]
            return df

        erp_df = filter_payments(erp_df, "description_erp")
        ven_df = filter_payments(ven_df, "description_ven")

        # 🧮 STEP 2: Start matching
        for _, e_row in erp_df.iterrows():
            e_inv = str(e_row.get("invoice_erp", "")).strip()
            e_amt = normalize_number(e_row.get("amount_erp", 0))
            if not e_inv:
                continue

            for _, v_row in ven_df.iterrows():
                v_inv = str(v_row.get("invoice_ven", "")).strip()
                v_desc = str(v_row.get("description_ven", "")).lower()

                # Compute Vendor amount correctly (Debit = invoice, Credit = CN)
                d_val = normalize_number(v_row.get("debit_ven", 0))
                c_val = normalize_number(v_row.get("credit_ven", 0))
                v_amt = c_val if "abono" in v_desc or "credit" in v_desc else d_val

                # Match by last 3–4 digits or fuzzy match
                if (
                    e_inv[-4:] == v_inv[-4:]
                    or e_inv[-3:] == v_inv[-3:]
                    or fuzz.ratio(e_inv, v_inv) > 82
                ):
                    diff = round(e_amt - v_amt, 2)
                    status = "Match" if abs(diff) < 0.05 else "Difference"

                    matched_rows.append({
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

                    matched_erp_invoices.add(e_inv)
                    matched_vendor_invoices.add(v_inv)
                    break  # stop checking vendor rows for this invoice

        # ✅ Build matched dataframe
        df_matched = pd.DataFrame(matched_rows)

        # 🧩 STEP 3: Find missing ones
        erp_missing = erp_df[
            ~erp_df["invoice_erp"].astype(str).isin(matched_erp_invoices)
        ].reset_index(drop=True)
        ven_missing = ven_df[
            ~ven_df["invoice_ven"].astype(str).isin(matched_vendor_invoices)
        ].reset_index(drop=True)

        return df_matched, erp_missing, ven_missing

    except Exception as e:
        st.error(f"❌ Matching failed: {str(e)}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
