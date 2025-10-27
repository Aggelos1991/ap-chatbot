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
        if any(k in r for k in ["credit","nota","abono","cn","πιστωτικό","πίστωση","ακυρωτικό"]):
            return "CN" if credit > 0 else "INV"
        if any(k in r for k in ["factura","invoice","inv","τιμολόγιο","παραστατικό"]) or debit > 0:
            return "INV"
        return "UNKNOWN"

    erp_df["__type"] = erp_df.apply(lambda r: doc_type(r, "erp"), axis=1)
    ven_df["__type"] = ven_df.apply(lambda r: doc_type(r, "ven"), axis=1)

    erp_df["__amt"] = erp_df.apply(lambda r: 
        normalize_number(r.get("debit_erp",0)) - normalize_number(r.get("credit_erp",0)), axis=1)
    ven_df["__amt"] = ven_df.apply(lambda r: 
        normalize_number(r.get("debit_ven",0)) - normalize_number(r.get("credit_ven",0)), axis=1)

    erp_use = erp_df[erp_df["__type"] != "IGNORE"].copy()
    ven_use = ven_df[ven_df["__type"] != "IGNORE"].copy()

    # === NET INVOICES + CREDIT NOTES ===
    def net_invoices(df, inv_col):
        out = []
        for inv, g in df.groupby(inv_col, dropna=False):
            if g.empty: continue
            inv_rows = g[g["__type"] == "INV"]
            cn_rows  = g[g["__type"] == "CN"]
            net_amt = inv_rows["__amt"].sum() - cn_rows["__amt"].sum()
            net_amt = round(net_amt, 2)
            if abs(net_amt) < 0.01:
                continue
            base = inv_rows.loc[inv_rows["__amt"].idxmax()] if not inv_rows.empty else cn_rows.iloc[0]
            base = base.copy()
            base["__amt"] = net_amt
            base["__type"] = "INV" if net_amt > 0 else "CN"
            out.append(base)
        return pd.DataFrame(out).reset_index(drop=True)

    erp_use = net_invoices(erp_use, "invoice_erp")
    ven_use = net_invoices(ven_use, "invoice_ven")

    # === TIER-1: EXACT MATCH (MATCH ON RAW, DISPLAY WITH PARENS) ===
    for e_idx, e in erp_use.iterrows():
        e_inv_raw = str(e.get("invoice_erp","")).strip()
        e_inv_display = f"({e_inv_raw})" if e["__type"] == "CN" else e_inv_raw
        e_amt = abs(round(float(e["__amt"]), 2))
        e_typ = e["__type"]
        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor: continue
            v_inv_raw = str(v.get("invoice_ven","")).strip()
            v_inv_display = f"({v_inv_raw})" if v["__type"] == "CN" else v_inv_raw
            v_amt = abs(round(float(v["__amt"]), 2))
            v_typ = v["__type"]

            # MATCH ON RAW INVOICE NUMBER
            if e_typ != v_typ or e_inv_raw != v_inv_raw: 
                continue

            diff = abs(e_amt - v_amt)
            status = "Perfect Match" if diff <= 0.01 else "Difference Match" if diff < 1.0 else None
            if status:
                matched.append({
                    "ERP Invoice": e_inv_display,
                    "Vendor Invoice": v_inv_display,
                    "ERP Amount": e_amt,
                    "Vendor Amount": v_amt,
                    "Difference": diff,
                    "Status": status
                })
                used_vendor.add(v_idx)
                break

    # SAFE RETURN
    cols = ["ERP Invoice","Vendor Invoice","ERP Amount","Vendor Amount","Difference","Status"]
    matched_df = pd.DataFrame(matched, columns=cols) if matched else pd.DataFrame(columns=cols)

    # Use raw invoice numbers for filtering
    matched_erp_raw = set(m["ERP Invoice"].strip("()") for m in matched)
    matched_ven_raw = set(m["Vendor Invoice"].strip("()") for m in matched)

    date_cols_erp = ["date_erp"] if "date_erp" in erp_use.columns else []
    date_cols_ven = ["date_ven"] if "date_ven" in ven_use.columns else []

    miss_erp = erp_use[~erp_use["invoice_erp"].isin(matched_ven_raw)].copy()
    miss_ven = ven_use[~ven_use["invoice_ven"].isin(matched_erp_raw)].copy()

    miss_erp["Invoice"] = miss_erp.apply(
        lambda r: f"({r['invoice_erp']})" if r["__type"] == "CN" else r["invoice_erp"], axis=1)
    miss_ven["Invoice"] = miss_ven.apply(
        lambda r: f"({r['invoice_ven']})" if r["__type"] == "CN" else r["invoice_ven"], axis=1)

    miss_erp = miss_erp[["Invoice", "__amt"] + date_cols_erp]
    miss_erp = miss_erp.rename(columns={"__amt": "Amount", "date_erp": "Date"})
    if "Date" not in miss_erp.columns: miss_erp["Date"] = ""

    miss_ven = miss_ven[["Invoice", "__amt"] + date_cols_ven]
    miss_ven = miss_ven.rename(columns={"__amt": "Amount", "date_ven": "Date"})
    if "Date" not in miss_ven.columns: miss_ven["Date"] = ""

    miss_erp = miss_erp[["Invoice","Amount","Date"]].reset_index(drop=True)
    miss_ven = miss_ven[["Invoice","Amount","Date"]].reset_index(drop=True)

    return matched_df, miss_erp, miss_ven
