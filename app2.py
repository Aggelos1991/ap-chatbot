def match_invoices(erp_df, ven_df):
    # Validate required columns
    if "invoice_erp" not in erp_df.columns or "invoice_ven" not in ven_df.columns:
        st.error("Missing invoice number column in one or both files. Check column names.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    matched = []
    used_vendor = set()

    def doc_type(row, tag):
        r = str(row.get(f"reason_{tag}", "")).lower()
        debit = normalize_number(row.get(f"debit_{tag}", 0))
        credit = normalize_number(row.get(f"credit_{tag}", 0))
        pay_pat = [r"^πληρωμ", r"^απόδειξη\s*πληρωμ", r"^payment", r"^bank\s*transfer",
                   r"^trf", r"^remesa", r"^pago", r"^pagado", r"^transferencia",
                   r"^εξοφληση", r"^paid"]
        if any(re.search(p, r) for p in pay_pat): return "IGNORE"
        if any(k in r for k in ["credit","nota","abono","cn","πιστωτικό","πίστωση","ακυρωτικό"]):
            return "CN" if credit > 0 else "INV"
        if any(k in r for k in ["factura","invoice","inv","τιμολόγιο","παραστατικό"]) or debit > 0:
            return "INV"
        return "UNKNOWN"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)

    erp_df["__amt"] = erp_df.apply(lambda r:
        normalize_number(r.get("debit_erp", 0)) - normalize_number(r.get("credit_erp", 0)), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r:
        normalize_number(r.get("debit_ven", 0)) - normalize_number(r.get("credit_ven", 0)), axis=1)

    erp_use = erp_df[erp_df["__type"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__type"] != "IGNORE"].copy()

    # === FILTER OUT MISSING INVOICE NUMBERS ===
    erp_use = erp_use[erp_use["invoice_erp"].notna() & (erp_use["invoice_erp"].str.strip() != "")]
    ven_use = ven_use[ven_use["invoice_ven"].notna() & (ven_use["invoice_ven"].str.strip() != "")]

    # === NET INVOICES + CREDIT NOTES ===
    def net_invoices(df, inv_col):
        out = []
        for inv, g in df.groupby(inv_col, dropna=False):
            inv_str = str(inv).strip()
            if not inv_str or inv_str == "nan" or inv_str.lower() == "none":
                continue
            inv_rows = g[g["__type"] == "INV"]
            cn_rows = g[g["__type"] == "CN"]
            net_amt = inv_rows["__amt"].sum() - cn_rows["__amt"].sum()
            net_amt = round(net_amt, 2)
            if abs(net_amt) < 0.01: continue
            base = inv_rows.loc[inv_rows["__amt"].idxmax()] if not inv_rows.empty else cn_rows.iloc[0]
            base = base.copy()
            base["__amt"] = net_amt
            base["__type"] = "INV" if net_amt > 0 else "CN"
            out.append(base)
        return pd.DataFrame(out).reset_index(drop=True)

    erp_use = net_invoices(erp_use, "invoice_erp")
    ven_use = net_invoices(ven_use, "invoice_ven")

    # === NORMALIZE INVOICE ===
    def normalize_invoice(v):
        return re.sub(r'\s+', '', str(v)).strip().upper()

    # === TIER-1: EXACT MATCH (TYPE-AGNOSTIC) ===
    for e_idx, e in erp_use.iterrows():
        e_inv_raw = str(e.get("invoice_erp", "")).strip()
        e_inv_norm = normalize_invoice(e_inv_raw)
        e_amt = abs(round(float(e["__amt"]), 2))

        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor: continue
            v_inv_raw = str(v.get("invoice_ven", "")).strip()
            v_inv_norm = normalize_invoice(v_inv_raw)
            v_amt = abs(round(float(v["__amt"]), 2))

            if e_inv_norm != v_inv_norm:
                continue
            if abs(e_amt - v_amt) > 0.01:
                continue

            matched.append({
                "ERP Invoice": e_inv_raw,
                "Vendor Invoice": v_inv_raw,
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": 0.0,
                "Status": "Perfect Match"
            })
            used_vendor.add(v_idx)
            break

    # ... rest unchanged (matched_df, miss_erp, miss_ven, etc.)
