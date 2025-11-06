def extract_with_gpt(lines):
    BATCH_SIZE = 60
    all_records = []

    for i in range(0, len(lines), BATCH_SIZE):
        text_block = "\n".join(lines[i:i + BATCH_SIZE])
        prompt = f"""
You are a multilingual financial data extractor specialized in **vendor statements** (Spanish / Greek / English).

⚙️ TASK:
Read each line carefully and extract ONLY valid transaction lines (Factura, Abono, Pago, Transferencia, Nota de crédito).
IGNORE accounting entries like "Asiento", "Diario", "Regularización", or summary lines.

For every valid transaction, return a JSON array of objects with:
- "Alternative Document" → the true document/invoice number.
- "Date" → the date appearing on the same or nearby line.
- "Reason" → one of ["Invoice", "Payment", "Credit Note"].
- "Debit"
- "Credit"
- "Balance"

📘 RULES:
- Skip lines that contain only “Asiento”, “Diario”, “Apertura”, “Regularización”, or “Saldo anterior”.
- The document number can appear after words like "Num", "Número", "Documento", "Factura", "FAC", "FV", "CO", "AB", "Doc.", or inside "Concepto" or comments like "por factura 12345".
- Reject values like "Asiento 204", "Remesa", "Pago", "Transferencia" if they are **not** followed by a real invoice/credit reference.
- Prefer any code matching patterns: (F|FV|CO|AB|FA|FAC)\d{{3,}} or at least 5 consecutive digits.
- If text mentions "Pago", "Cobro", "Transferencia", "Remesa" → Reason = "Payment".
- If text mentions "Abono", "Nota de crédito", "Crédit", "Descuento", "Πίστωση" → Reason = "Credit Note".
- Otherwise, assume "Invoice".
- Do NOT invent any fields.
- Return only a pure JSON array, nothing else.

Text:
{text_block}
"""
        data = []
        for model in [PRIMARY_MODEL, BACKUP_MODEL]:
            try:
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0
                )
                content = response.choices[0].message.content.strip()
                if i == 0:
                    st.text_area(f"🧠 GPT Response (Batch 1 – {model})", content, height=250, key=f"debug_{model}")
                data = parse_gpt_response(content, i // BATCH_SIZE + 1)
                if data:
                    break
            except Exception as e:
                st.warning(f"GPT error ({model}): {e}")

        if not data:
            continue

        # === Data post-processing ===
        for row in data:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            if not alt_doc:
                continue

            # Skip garbage
            if re.search(r"asiento|diario|apertura|regularizaci", alt_doc, re.IGNORECASE):
                continue
            if re.match(r"^(pago|remesa|transfer|trf|bank)$", alt_doc, re.IGNORECASE):
                continue

            # Force detect real doc number (fallback)
            match = re.search(r"((F|FV|CO|AB|FAC|FA)\d{3,}|\d{5,})", alt_doc)
            if match:
                alt_doc = match.group(1)
            else:
                continue  # no valid doc pattern found

            debit_val = normalize_number(row.get("Debit", ""))
            credit_val = normalize_number(row.get("Credit", ""))
            balance_val = normalize_number(row.get("Balance", ""))
            reason = str(row.get("Reason", "")).strip().lower()

            # Reason correction
            if re.search(r"pago|cobro|transfer|remesa|trf|bank", str(row), re.IGNORECASE):
                reason = "Payment"
            elif re.search(r"abono|nota\s*de\s*crédito|crédit|descuento|πίστωση", str(row), re.IGNORECASE):
                reason = "Credit Note"
            else:
                reason = "Invoice"

            # Correct Debit/Credit side
            if reason == "Payment":
                if debit_val and not credit_val:
                    credit_val, debit_val = debit_val, 0
            elif reason in ["Invoice", "Credit Note"]:
                if credit_val and not debit_val:
                    debit_val, credit_val = credit_val, 0

            # Skip blank lines
            if debit_val == "" and credit_val == "":
                continue

            all_records.append({
                "Alternative Document": alt_doc,
                "Date": str(row.get("Date", "")).strip(),
                "Reason": reason.title(),
                "Debit": debit_val,
                "Credit": credit_val,
                "Balance": balance_val
            })

    return all_records
