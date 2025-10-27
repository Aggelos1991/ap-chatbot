def match_invoices(erp_df, ven_df):
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
        if any(k in r for k in ["credit", "nota", "abono", "cn", "πιστωτικό",
                                "πίστωση", "ακυρωτικό"]): return "CN"
        if any(k in r for k in ["factura", "invoice", "inv", "τιμολόγιο",
                                "παραστατικό"]) or debit > 0: return "INV"
        return "UNKNOWN"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)
    erp_df["__amt"] = erp_df.apply(lambda r: abs(normalize_number(r.get("debit_erp", 0)) -
                                                normalize_number(r.get("credit_erp", 0))), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven", 0)) -
                                                normalize_number(r.get("credit_ven", 0))), axis=1)

    erp_use = erp_df[erp_df["__type"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__type"] != "IGNORE"].copy()

    def merge_inv_cn(df, inv_col):
        out = []
        for inv, g in df.groupby(inv_col, dropna=False):
            if g.empty: continue
            inv_rows = g[g["__type"] == "INV"]
            cn_rows = g[g["__type"] == "CN"]
            if not inv_rows.empty and not cn_rows.empty:
                net = round(abs(inv_rows["__amt"].sum() - cn_rows["__amt"].sum()), 2)
                base = inv_rows.iloc[-1].copy()
                base["__amt"] = net
                out.append(base)
            else:
                out.append(g.loc[g["__amt"].idxmax()])
        return pd.DataFrame(out).reset_index(drop=True)

    erp_use = merge_inv_cn(erp_use, "invoice_erp")
    ven_use = merge_inv_cn(ven_use, "invoice_ven")

    for e_idx, e in erp_use.iterrows():
        e_inv = str(e.get("invoice_erp", "")).strip()
        e_amt = round(float(e["__amt"]), 2)
        e_typ = e["__type"]
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor: continue
            v_inv = str(v.get("invoice_ven", "")).strip()
            v_amt = round(float(v["__amt"]), 2)
            v_typ = v["__type"]
            if e_typ != v_typ or e_inv != v_inv: continue
            diff = abs(e_amt - v_amt)
            status = "Perfect Match" if diff <= 0.01 else "Difference Match" if diff < 1.0 else None
            if status:
                matched.append({
                    "ERP Invoice": e_inv,
                    "Vendor Invoice": v_inv,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": status
                })
                used_vendor.add(v_idx)
                break

    matched_df = pd.DataFrame(matched)

    # === FIXED: SAFE COLUMN ACCESS ===
    matched_erp = set(matched_df["ERP Invoice"]) if not matched_df.empty and "ERP Invoice" in matched_df.columns else set()
    matched_ven = set(matched_df["Vendor Invoice"]) if not matched_df.empty and "Vendor Invoice" in matched_df.columns else set()
    # =================================

    date_cols_erp = ["date_erp"] if "date_erp" in erp_use.columns else []
    date_cols_ven = ["date_ven"] if "date_ven" in ven_use.columns else []

    miss_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven)][["invoice_erp", "__amt"] + date_cols_erp] \
        .rename(columns={"invoice_erp": "Invoice", "__amt": "Amount", "date_erp": "Date"})
    miss_ven = ven_use[~ven_use["invoice_ven"].isin(matched_erp)][["invoice_ven", "__amt"] + date_cols_ven] \
        .rename(columns={"invoice_ven": "Invoice", "__amt": "Amount", "date_ven": "Date"})

    return matched_df, miss_erp, miss_ven
