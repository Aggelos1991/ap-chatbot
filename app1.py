def preprocess_text_for_ai(raw_text):
    """
    Preprocess vendor statement text:
    - Keep original line context (document, date, concept)
    - Tag only DEBE column values (the ones before SALDO)
    - Ignore left-side numeric fields and 0,00 values
    - Mark Credit Notes explicitly
    """
    txt = raw_text
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r",\s+", ",", txt)

    lines = txt.split("\n")
    clean_lines = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Skip irrelevant sections
        if re.search(r"(?i)\b(SALDO\s+ANTERIOR|BANCO|COBRO|EFECTO|REME|PAGO)\b", line):
            continue

        # Extract numeric values (EU or US format)
        nums = re.findall(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", line)
        if len(nums) < 2:
            continue

        # DEBE is the SECOND number from the RIGHT (before HABER/SALDO)
        amount = nums[-2]

        # Skip 0,00 or 0.00 lines
        if re.match(r"^0+[.,]0+$", amount):
            continue

        # Add credit note tagging
        if re.search(r"(?i)(ABONO|NOTA\s+DE\s+CR[EÉ]DITO|CREDIT\s+NOTE|C\.?N\.?)", line):
            line = re.sub(re.escape(amount), f"[CREDIT: -{amount}]", line, count=1)
        else:
            line = re.sub(re.escape(amount), f"[DEBE: {amount}]", line, count=1)

        # Keep line context for GPT
        clean_lines.append(line)

    return "\n".join(clean_lines)
