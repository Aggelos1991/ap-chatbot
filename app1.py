def preprocess_text_for_ai(raw_text):
    """
    Preprocess vendor statement text:
    - Keep original line context (dates, document numbers, concept)
    - Tag only DEBE (second-to-last numeric value)
    - Ignore SALDO/HABER
    - Tag Credit Notes explicitly
    """
    txt = raw_text
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r",\s+", ",", txt)

    lines = txt.split("\n")
    clean_lines = []

    for line in lines:
        if not line.strip():
            continue

        # Skip irrelevant lines
        if re.search(r"(?i)\b(SALDO\s+ANTERIOR|BANCO|COBRO|EFECTO|REME|PAGO)\b", line):
            continue

        # Extract all numeric values (1.234,56 / 1234,56 / 1,234.56)
        nums = re.findall(r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", line)
        if not nums:
            continue

        # Get DEBE = second-to-last number
        if len(nums) >= 2:
            amount = nums[-2]
        else:
            amount = nums[-1]

        # Tag DEBE or CREDIT but KEEP the whole line
        if re.search(r"(?i)(ABONO|NOTA\s+DE\s+CR[EÉ]DITO|CREDIT\s+NOTE|C\.?N\.?)", line):
            line = re.sub(re.escape(amount), f"[CREDIT: -{amount}]", line, count=1)
        else:
            line = re.sub(re.escape(amount), f"[DEBE: {amount}]", line, count=1)

        # Keep document numbers, dates, etc.
        clean_lines.append(line.strip())

    return "\n".join(clean_lines)
