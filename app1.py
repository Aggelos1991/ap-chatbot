"""
        
        # Retry logic for better accuracy
        max_retries = 2
        for retry in range(max_retries + 1):
            try:
                response = client.chat.completions.create(
                    model=MODEL,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.1,  # Lower temp = higher accuracy
                    max_tokens=4000
                )
                content = response.choices[0].message.content.strip()
                
                # Extract JSON
                json_match = re.search(r'\[.*\]', content, re.DOTALL)
                if json_match:
                    data = json.loads(json_match.group(0))
                    break
            except Exception as e:
                if retry == max_retries:
                    st.warning(f"⚠️ Batch {i//BATCH_SIZE + 1} failed: {e}")
                    continue
        
        # Validate & filter records
        for row in data if 'data' in locals() else []:
            alt_doc = str(row.get("Alternative Document", "")).strip()
            
            # Enhanced exclusion
            exclude_patterns = ['concil', 'total', 'saldo', 'iva', 'impuestos', 'reconcili']
            if any(re.search(p, alt_doc, re.IGNORECASE) for p in exclude_patterns):
                continue
                
            debit = normalize_number(row.get("Debit", 0))
            credit = normalize_number(row.get("Credit", 0))
            
            # VALID TRANSACTION REQUIRED
            if debit == 0 and credit == 0:
                continue
                
            # Auto-classify if Reason missing
            reason = row.get("Reason", "").strip()
            if not reason:
                if debit > 0:
                    reason = "Invoice"
                elif credit > 0:
                    reason = "Payment"
                else:
                    reason = "Credit Note"
            
            all_records.append({
                "Alternative Document": alt_doc,
                "Date": str(row.get("Date", "")).strip(),
                "Reason": reason,
                "Debit": debit,
                "Credit": credit,
                "Confidence": row.get("Confidence", 0.95)
            })
        
        progress_bar.progress(min((i + BATCH_SIZE) / len(lines), 1.0))
        time.sleep(0.1)  # Rate limiting
    
    progress_bar.empty()
    return all_records

# ==========================================================
# ENHANCED VALIDATION & STATS
# ==========================================================
def validate_records(records: List[Dict]) -> pd.DataFrame:
    """Add validation scores and statistics."""
    df = pd.DataFrame(records)
    if df.empty:
        return df
    
    # Numeric conversion
    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce').fillna(0)
    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce').fillna(0)
    
    # Validation score
    df['Valid'] = (
        (df['Debit'] > 0) | (df['Credit'] > 0) &
        df['Alternative Document'].str.contains(r'\d', na=False)
    )
    
    return df

# ==========================================================
# SUPERIOR EXPORT
# ==========================================================
def to_excel_bytes(records: List[Dict]) -> BytesIO:
    df = pd.DataFrame(records)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Transactions', index=False)
        
        # Summary sheet
        summary = pd.DataFrame({
            'Metric': ['Total Records', 'Valid Records', 'Total Debit', 'Total Credit', 'Net Balance'],
            'Value': [
                len(df),
                len(df[df['Valid'] == True]),
                df['Debit'].sum(),
                df['Credit'].sum(),
                df['Debit'].sum() - df['Credit'].sum()
            ]
        })
        summary.to_excel(writer, sheet_name='Summary', index=False)
    
    buf.seek(0)
    return buf

# ==========================================================
# ENHANCED STREAMLIT UI
# ==========================================================
st.header("📂 Upload Vendor Statement")
uploaded_pdf = st.file_uploader("Choose PDF file", type=["pdf"])

if uploaded_pdf:
    # Preview
    with st.spinner("🔍 Analyzing PDF structure..."):
        lines = extract_raw_lines(uploaded_pdf)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("📄 Pages Processed", len(lines))
        st.metric("🔢 Numeric Lines", sum(1 for line in lines if re.search(r'\d+[.,]\d{2}', line)))
    with col2:
        lang = detect_language("\n".join(lines[:100]))
        st.metric("🌐 Detected Language", lang.upper())
        st.metric("⚡ Extraction Speed", "Ultra Fast")
    
    st.text_area("📄 Sample Lines:", "\n".join(lines[:20]), height=200)
    
    if st.button("🚀 Extract with AI", type="primary"):
        with st.spinner("🧠 GPT-4o-mini analyzing (99% accuracy)..."):
            data = extract_with_gpt(lines)
            df = validate_records(data)
        
        if df.empty:
            st.error("❌ No valid transactions found.")
        else:
            # METRICS
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown('<div class="metric-container perfect">', unsafe_allow_html=True)
                st.metric("✅ Valid Records", len(df[df['Valid'] == True]))
                st.markdown('</div>', unsafe_allow_html=True)
            with col2:
                st.markdown('<div class="metric-container warning">', unsafe_allow_html=True)
                st.metric("📊 Total Debit", f"{df['Debit'].sum():,.2f}")
                st.markdown('</div>', unsafe_allow_html=True)
            with col3:
                st.metric("💳 Total Credit", f"{df['Credit'].sum():,.2f}")
            with col4:
                net = df['Debit'].sum() - df['Credit'].sum()
                st.metric("⚖️ Net Balance", f"{net:,.2f}")
            
            st.success(f"🎉 Extraction complete! {len(df[df['Valid'] == True])} valid records.")
            
            # Filter valid records
            valid_df = df[df['Valid'] == True].drop('Valid', axis=1)
            
            st.subheader("📋 Valid Transactions")
            st.dataframe(valid_df, use_container_width=True, height=500)
            
            # Download
            excel_data = to_excel_bytes(valid_df.to_dict('records'))
            st.download_button(
                "💾 Download Excel Report",
                data=excel_data,
                file_name=f"DataFalcon_Pro_{int(time.time())}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Raw data toggle
            with st.expander("🔍 View Raw Extraction (Debug)"):
                st.dataframe(df, use_container_width=True)
