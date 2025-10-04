def run_query(q: str, df: pd.DataFrame):
    if df is None or df.empty:
        return "⚠️ Please upload an Excel file first.", None

    ql = q.lower().strip()

    # Prep numeric/date/status
    working = df.copy()
    working["amount"] = pd.to_numeric(working["amount"], errors="coerce")
    working["due_date_parsed"] = pd.to_datetime(working["due_date"], errors="coerce")
    working["status"] = working["status"].astype(str).str.lower()

    # ---------- Multi-invoice or single invoice handling ----------
    invoice_hits = pd.DataFrame()
    # Find all invoice IDs in query (e.g. INV1001, INV-1002)
    invoice_ids = re.findall(r"[a-z]{2,}[- ]?\d{3,}", ql)
    if invoice_ids:
        all_hits = []
        for inv in invoice_ids:
            hits = find_invoices_in_query(working, inv)
            if not hits.empty:
                all_hits.append(hits)
        if all_hits:
            invoice_hits = pd.concat(all_hits).drop_duplicates(subset=["invoice_no"])

    # ---------- If we found invoices ----------
    if not invoice_hits.empty:
        # If user asks for vendor names
        if "vendor" in ql and "email" not in ql:
            vendors = (
                invoice_hits["vendor_name"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )
            return f"🏢 Vendors: {', '.join(vendors)}", None

        # If user asks for emails
        if "email" in ql or "emails" in ql:
            emails = (
                invoice_hits["vendor_email"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )
            if not emails:
                return "No vendor emails found for these invoices.", None
            return f"📧 Emails: {'; '.join(sorted(emails, key=str.lower))}", None

        # If user asks for amounts
        if "amount" in ql:
            amounts = [
                fmt_money(r["amount"], r["currency"]) for _, r in invoice_hits.iterrows()
            ]
            return f"💰 Amounts: {', '.join(amounts)}", None

        # Otherwise, generic invoice summaries
        lines = []
        for _, r in invoice_hits.iterrows():
            lines.append(
                f"Invoice **{r['invoice_no']}** — vendor **{r.get('vendor_name','-')}**, "
                f"status **{r.get('status','-')}**, amount **{fmt_money(r.get('amount'), r.get('currency'))}**, "
                f"due **{r.get('due_date','-')}**."
            )
        return "\n\n".join(lines), invoice_hits.reset_index(drop=True)

    # ---------- Otherwise fall back to filters (same as before) ----------
    # Status
    if any(w in ql for w in ["open", "unpaid", "pending"]):
        working = working[working["status"].str.contains("open|unpaid|pending", case=False, na=False)]
    elif "paid" in ql and not any(w in ql for w in ["unpaid", "not paid", "open", "pending"]):
        working = working[working["status"].str.contains("paid", case=False, na=False)]

    # Amount filters
    m_over = re.search(r"(over|above|greater than|>=)\s*([0-9][0-9,\.]*)", ql)
    if m_over:
        val = float(m_over.group(2).replace(",", ""))
        working = working[working["amount"] >= val]
    m_under = re.search(r"(under|below|less than|<=)\s*([0-9][0-9,\.]*)", ql)
    if m_under:
        val2 = float(m_under.group(2).replace(",", ""))
        working = working[working["amount"] <= val2]

    # Due date filters
    di = extract_date_from_text(ql)
    if di and pd.notna(working["due_date_parsed"]).any():
        mode = di["mode"]
        d1 = di.get("d1")
        d2 = di.get("d2")
        if mode == "before" and pd.notna(d1):
            working = working[working["due_date_parsed"] <= d1]
        elif mode == "after" and pd.notna(d1):
            working = working[working["due_date_parsed"] >= d1]
        elif mode == "on" and pd.notna(d1):
            working = working[working["due_date_parsed"].dt.date == d1.date()]
        elif mode == "between" and pd.notna(d1) and pd.notna(d2):
            working = working[(working["due_date_parsed"] >= d1) & (working["due_date_parsed"] <= d2)]

    # Email filters
    if "email" in ql or "emails" in ql:
        emails = (
            working["vendor_email"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
            .tolist()
        )
        if not emails:
            return "No vendor emails found for this query.", None
        return f"📧 Emails: {'; '.join(sorted(emails, key=str.lower))}", None

    if working.empty:
        return "No invoices match your filters.", None

    return f"Found {len(working)} invoices matching your query.", working.reset_index(drop=True)
