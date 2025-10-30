3. Never output or use SALDO, TOTAL, IVA or ASIENTO lines.
4. Do not include headers or summaries.
5. Output only valid transactions in JSON array format below.

**OUTPUT FORMAT (JSON only):**
[
  {{
    "Alternative Document": "...",
    "Date": "dd/mm/yy",
    "Reason": "Invoice | Payment | Credit Note",
    "Debit": "DEBE amount",
    "Credit": "HABER amount"
  }}
]

Text to analyze:
{text_block}
"""

        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.0
            )
            content = response.choices[0].message.content.strip()

            # Debug preview
            if i == 0:
                st.text_area("GPT Response (Batch 1):", content, height=200, key="debug_1")

            json_match = re.search(r'\[.*\]', content, re.DOTALL)
            if not json_match:
                json_match = re.search(r'(\[.*?\])', content, re.DOTALL)

            if json_match:
                json_str = json_match.group(0)
                data = json.loads(json_str)

                for row in data:
                    alt_doc = str(row.get("Alternative Document", "")).strip()

                    # Skip non-transactional lines
                    if not alt_doc or re.search(r"(asiento|saldo|total|iva)", alt_doc, re.IGNORECASE):
                        continue

                    debit_raw = row.get("Debit", "")
                    credit_raw = row.get("Credit", "")

                    debit_val = normalize_number(debit_raw)
                    credit_val = normalize_number(credit_raw)
                    reason = row.get("Reason", "").strip()

                    # Apply final classification enforcement
                    if debit_val and not credit_val:
                        reason = "Invoice"
                    elif credit_val and not debit_val:
                        if re.search(r"abono|nota|crédit|descuento", str(row), re.IGNORECASE):
                            reason = "Credit Note"
                        else:
                            reason = "Payment"
                    else:
                        continue

                    all_records.append({
                        "Alternative Document": alt_doc,
                        "Date": str(row.get("Date", "")).strip(),
                        "Reason": reason,
                        "Debit": debit_val,
                        "Credit": credit_val
                    })
            else:
                st.warning(f"No JSON found in batch {i//BATCH_SIZE + 1}")

        except Exception as e:
            st.warning(f"GPT error batch {i//BATCH_SIZE + 1}: {e}")
            continue

    return all_records

# ==========================================================
# EXPORT
# ==========================================================
def to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ==========================================================
# STREAMLIT UI
# ==========================================================
uploaded_pdf = st.file_uploader("📂 Upload Vendor Statement (PDF)", type=["pdf"])

if uploaded_pdf:
    with st.spinner("📄 Extracting text from all pages..."):
        lines = extract_raw_lines(uploaded_pdf)

    st.success(f"✅ Found {len(lines)} lines of text!")
    st.text_area("📄 Preview (first 30 lines):", "\n".join(lines[:30]), height=300)

    if st.button("🤖 Run Hybrid Extraction", type="primary"):
        with st.spinner("Analyzing with GPT-4o-mini..."):
            data = extract_with_gpt(lines)

        if data:
            df = pd.DataFrame(data)
            st.success(f"✅ Extraction complete — {len(df)} valid records found!")
            st.dataframe(df, use_container_width=True, hide_index=True)

            try:
                total_debit = df["Debit"].apply(pd.to_numeric, errors="coerce").sum()
                total_credit = df["Credit"].apply(pd.to_numeric, errors="coerce").sum()
                net = round(total_debit - total_credit, 2)

                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Debit", f"{total_debit:,.2f}")
                col2.metric("💳 Total Credit", f"{total_credit:,.2f}")
                col3.metric("⚖️ Net", f"{net:,.2f}")

            except Exception as e:
                st.error(f"Totals error: {e}")

            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(data),
                file_name=f"vendor_statement_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No structured data detected. Check GPT response above.")
else:
    st.info("Please upload a vendor statement PDF to begin.")
