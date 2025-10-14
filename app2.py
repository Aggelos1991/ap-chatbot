def match_invoices(erp_df, ven_df):
    matched = []
    used_vendor_rows = set()

    # ====== ERP PREP ======
    erp_df["__doctype"] = erp_df.apply(
        lambda r: "CN" if normalize_number(r.get("debit_erp")) > 0
        else ("INV" if normalize_number(r.get("credit_erp")) > 0 else "UNKNOWN"),
        axis=1
    )
    erp_df["__amt"] = erp_df.apply(
        lambda r: normalize_number(r["credit_erp"]) if r["__doctype"] == "INV"
        else (-normalize_number(r["debit_erp"]) if r["__doctype"] == "CN" else 0.0),
        axis=1
    )

    # ====== VENDOR PREP ======
    ven_df["__doctype"] = ven_df.apply(
        lambda r: "CN" if normalize_number(r.get("debit_ven")) < 0 else "INV",
        axis=1
    )
    ven_df["__amt"] = ven_df.apply(lambda r: abs(normalize_number(r.get("debit_ven"))), axis=1)

    erp_use = erp_df[erp_df["__doctype"].isin(["INV", "CN"])].copy()
    ven_use = ven_df[ven_df["__doctype"].isin(["INV", "CN"])].copy()

    # ====== (Optional) merge ERP INV+CN with same invoice to a net line ======
    merged_rows = []
    for inv, g in erp_use.groupby("invoice_erp", dropna=False):
        if len(g) == 1:
            merged_rows.append(g.iloc[0])
            continue
        inv_rows = g[g["__doctype"] == "INV"]
        cn_rows  = g[g["__doctype"] == "CN"]
        if not inv_rows.empty and not cn_rows.empty:
            base = inv_rows.iloc[0].copy()
            base["__amt"] = round(inv_rows["__amt"].sum() + cn_rows["__amt"].sum(), 2)
            merged_rows.append(base)
        else:
            merged_rows.extend(list(g.itertuples(index=False, name=None)))
            merged_rows = [pd.Series(r, index=g.columns) for r in merged_rows]  # normalize
    if merged_rows and not isinstance(merged_rows[0], pd.Series):
        erp_use = erp_use  # fallback if above branch didn’t run
    else:
        erp_use = pd.DataFrame(merged_rows, columns=erp_use.columns).reset_index(drop=True)

    # ====== NORMALIZE invoice “cores” (digits only) ======
    def core(s):
        s = str(s or "")
        only_d = re.sub(r"[^0-9]", "", s)
        return only_d

    erp_use["__core"] = erp_use["invoice_erp"].apply(core)
    ven_use["__core"] = ven_use["invoice_ven"].apply(core)

    # ====== MATCHING (exact → last3 → suffix) ======
    for e_idx, e in erp_use.iterrows():
        e_inv  = str(e["invoice_erp"]).strip()
        e_core = e["__core"]
        e_amt  = round(float(e["__amt"]), 2)
        e_date = e.get("date_erp")

        best_score = -1
        best_v = None

        for v_idx, v in ven_use.iterrows():
            if v_idx in used_vendor_rows:
                continue

            v_inv  = str(v["invoice_ven"]).strip()
            v_core = v["__core"]
            v_amt  = round(float(v["__amt"]), 2)
            v_date = v.get("date_ven")

            # Rule 1: full invoice string equal
            if e_inv.lower() == v_inv.lower():
                score = 200
            # Rule 2: last 3 digits equal (compare numeric cores)
            elif len(e_core) >= 3 and len(v_core) >= 3 and e_core[-3:] == v_core[-3:]:
                score = 150
            # Rule 3: suffix/prefix numeric (e.g. PSF000012 ↔ 12)
            elif (e_core and v_core) and (e_core.endswith(v_core) or v_core.endswith(e_core)):
                score = 130
            else:
                score = 0

            if score > best_score:
                best_score = score
                best_v = (v_idx, v_inv, v_amt, v_date)

        # accept only strong rules (>= 130)
        if best_v is not None and best_score >= 130:
            v_idx, v_inv, v_amt, v_date = best_v
            used_vendor_rows.add(v_idx)
            diff = round(e_amt - v_amt, 2)
            status = "Match" if abs(diff) < 0.05 else "Difference"

            matched.append({
                "Date (ERP)": e_date,
                "Date (Vendor)": v_date,
                "ERP Invoice": e_inv if e_inv else "(inferred)",
                "Vendor Invoice": v_inv,
                "ERP Amount": e_amt,
                "Vendor Amount": v_amt,
                "Difference": diff,
                "Status": status
            })
        # else: no pair selected → will show in Missing tables below

    # ====== MISSING TABLES (pure set difference) ======
    matched_erp_codes = set(m["ERP Invoice"] for m in matched)
    matched_ven_codes = set(m["Vendor Invoice"] for m in matched)

    # Missing in ERP = in vendor file but not matched
    ven_missing_rows = []
    for _, r in ven_use.iterrows():
        code = str(r.get("invoice_ven"))
        if code not in matched_ven_codes:
            ven_missing_rows.append({
                "Date": r.get("date_ven"),
                "Invoice": code,
                "Amount": round(float(r.get("__amt", 0.0)), 2),
            })
    missing_erp_final = pd.DataFrame(ven_missing_rows) if ven_missing_rows else pd.DataFrame(columns=["Date","Invoice","Amount"])

    # Missing in Vendor = in ERP file but not matched
    erp_missing_rows = []
    for _, r in erp_use.iterrows():
        code = str(r.get("invoice_erp"))
        if code not in matched_erp_codes:
            erp_missing_rows.append({
                "Date": r.get("date_erp"),
                "Invoice": code,
                "Amount": round(float(r.get("__amt", 0.0)), 2),
            })
    missing_vendor_final = pd.DataFrame(erp_missing_rows) if erp_missing_rows else pd.DataFrame(columns=["Date","Invoice","Amount"])

    # tidy types
    if isinstance(matched, list):
        matched = pd.DataFrame(matched)
    for df in [matched, missing_erp_final, missing_vendor_final]:
        if not df.empty and "Invoice" in df.columns:
            df["Invoice"] = df["Invoice"].astype(str).str.strip()

    return matched, missing_erp_final, missing_vendor_final
