# age_invoicer_pro.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io

# === CONFIG ===
st.set_page_config(page_title="Age Invoicer Pro", layout="wide")
st.title("Age Invoicer Pro")
st.markdown("**Interactive overdue invoice analytics • Click any bar to drill down • Export raw data**")

# Initialize session state
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None
if 'top_n' not in st.session_state:
    st.session_state.top_n = 30

# === FILE UPLOAD ===
uploaded_file = st.file_uploader("Upload Excel File (Sheet: 'Outstanding Invoices IB')", type=['xlsx'])

if not uploaded_file:
    st.info("Upload your invoice file to begin.")
    st.stop()

try:
    # === READ EXCEL (Only needed columns) ===
    with pd.ExcelFile(uploaded_file) as xls:
        if 'Outstanding Invoices IB' not in xls.sheet_names:
            st.error("Sheet 'Outstanding Invoices IB' not found.")
            st.stop()

        # Columns: A(0), B(1), E(4), G(6), AD(29), AE(30), AF(31), AH(33), AJ(35), AN(39), BD(55), BJ(61)
        keep_cols = [0, 1, 4, 6, 29, 30, 31, 33, 35, 39, 55, 61]
        df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None, usecols=keep_cols)

    # === FIND HEADER ROW ===
    header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
    if header_row.empty:
        st.error("Header 'VENDOR' not found in column A.")
        st.stop()

    start_row = header_row[0] + 1
    df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

    # === ASSIGN COLUMNS ===
    df.columns = [
        'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
        'Vendor_Email', 'Account_Email',
        'AF', 'AH', 'AJ', 'AN', 'BD', 'BJ_Alt_Invoice'
    ]

    # === FILTER LOGIC ===
    yes_mask = (
        (df['AF'].astype(str).str.strip().str.upper() == 'YES') &
        (df['AH'].astype(str).str.strip().str.upper() == 'YES') &
        (df['AJ'].astype(str).str.strip().str.upper() == 'YES') &
        (df['AN'].astype(str).str.strip().str.upper() == 'YES')
    )

    bd_keywords = ['ENTERTAINMENT', 'FALSE', 'REGULAR', 'PRIORITY VENDOR', 'PRIORITY VENDOR OS&E']
    bd_mask = df['BD'].astype(str).str.upper().apply(
        lambda x: any(k in x for k in bd_keywords)
    )

    df = df[yes_mask & bd_mask].reset_index(drop=True)
    df = df.drop(columns=['AF', 'AH', 'AJ', 'AN', 'BD'])

    if df.empty:
        st.warning("No invoices match filter criteria.")
        st.stop()

    # === CLEAN DATA ===
    df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
    df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
    df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
    df = df[df['Open_Amount'] > 0].copy()

    if df.empty:
        st.warning("No valid open invoices after cleaning.")
        st.stop()

    # === OVERDUE STATUS ===
    today = pd.Timestamp.today().normalize()
    df['Overdue'] = df['Due_Date'] < today
    df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

    # === SUMMARY BY VENDOR (ENSURE BOTH COLUMNS) ===
    summary = (
        df.groupby(['Vendor_Name', 'Status'])['Open_Amount']
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    # Ensure both columns exist
    for col in ['Overdue', 'Not Overdue']:
        if col not in summary.columns:
            summary[col] = 0

    summary['Total'] = summary['Overdue'] + summary['Not Overdue']

    # === FILTERS ===
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        status_filter = st.selectbox(
            "Filter By",
            ["All Open", "Overdue Only", "Not Overdue Only"],
            key="status_filter"
        )
    with col2:
        vendor_options = ["Top 30"] + sorted(df['Vendor_Name'].unique().tolist())
        selected_vendor = st.selectbox("Select Vendor", vendor_options, key="vendor_select")
    with col3:
        st.session_state.top_n = st.selectbox("Top N", [10, 20, 30, 50], index=2)

    # === TOP N LOGIC ===
    top_n = st.session_state.top_n

    if status_filter == "All Open":
        top_df = summary.nlargest(top_n, 'Total').copy()
        title = f"Top {top_n} Vendors (All Open)"
    elif status_filter == "Overdue Only":
        top_df = summary.nlargest(top_n, 'Overdue').copy()
        top_df['Not Overdue'] = 0
        title = f"Top {top_n} Vendors (Overdue Only)"
    else:  # Not Overdue Only
        if summary['Not Overdue'].sum() == 0:
            st.warning("No 'Not Overdue' invoices found.")
            top_df = summary.head(0).copy()
            top_df['Overdue'] = 0
            top_df['Not Overdue'] = 0
        else:
            top_df = summary.nlargest(top_n, 'Not Overdue').copy()
            top_df['Overdue'] = 0
        title = f"Top {top_n} Vendors (Not Overdue Only)"

    # === SELECT SINGLE VENDOR IF CHOSEN ===
    if selected_vendor != "Top 30":
        base_df = summary[summary['Vendor_Name'] == selected_vendor].copy()
        if base_df.empty:
            st.error("Selected vendor not found.")
            st.stop()
    else:
        base_df = top_df

    # === MELT FOR PLOT ===
    plot_df = base_df.melt(
        id_vars='Vendor_Name',
        value_vars=['Overdue', 'Not Overdue'],
        var_name='Type',
        value_name='Amount'
    )
    plot_df = plot_df[plot_df['Amount'] > 0].copy()

    if plot_df.empty:
        st.info("No data to display for the selected filter.")
        st.stop()

    # Add total for labels
    total_map = base_df.set_index('Vendor_Name')['Total'].to_dict()
    plot_df['Total'] = plot_df['Vendor_Name'].map(total_map)

    # === PLOTLY BAR CHART ===
    fig = px.bar(
        plot_df,
        x='Amount',
        y='Vendor_Name',
        color='Type',
        orientation='h',
        title=title,
        color_discrete_map={'Overdue': '#B22222', 'Not Overdue': '#4682B4'},
        height=max(600, len(plot_df) * 40),
        text=None
    )

    # Add total labels
    totals = plot_df.groupby('Vendor_Name')['Amount'].sum().reset_index()
    fig.add_scatter(
        x=totals['Amount'],
        y=totals['Vendor_Name'],
        mode='text',
        text=totals['Amount'].apply(lambda x: f'€{x:,.0f}'),
        textposition='top center',
        textfont=dict(size=13, color='white', family='Arial Black', weight='bold'),
        showlegend=False,
        hoverinfo='skip'
    )

    fig.update_layout(
        xaxis_title="Amount (€)",
        yaxis_title="Vendor",
        legend_title="Status",
        barmode='stack',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=180, r=50, t=80, b=50),
        yaxis={'categoryorder': 'total ascending'}
    )

    # === INTERACTIVE CHART ===
    chart = st.plotly_chart(fig, use_container_width=True, key="main_chart", on_select="rerun")

    # === CAPTURE CLICK ===
    clicked_vendor = None
    if chart.selection and chart.selection['points']:
        point = chart.selection['points'][0]
        clicked_vendor = point['y']
        st.session_state.clicked_vendor = clicked_vendor

    show_vendor = st.session_state.clicked_vendor

    # === SHOW ONLY DATA IN CURRENT VIEW (TOP N + FILTER) ===
    if show_vendor:
        st.subheader(f"{show_vendor}")

        # Get data in current view (same filter + top N logic)
        view_vendors = base_df['Vendor_Name'].tolist()
        filtered_df = df[df['Vendor_Name'].isin(view_vendors)].copy()

        # Further filter by status if needed
        if status_filter == "Overdue Only":
            filtered_df = filtered_df[filtered_df['Overdue']]
        elif status_filter == "Not Overdue Only":
            filtered_df = filtered_df[~filtered_df['Overdue']]

        vendor_data = filtered_df[filtered_df['Vendor_Name'] == show_vendor].copy()

        if vendor_data.empty:
            st.info("No invoices in current view for this vendor.")
        else:
            display_cols = ['VAT_ID', 'Due_Date', 'Open_Amount', 'BJ_Alt_Invoice', 'Status', 'Vendor_Email', 'Account_Email']
            display_data = vendor_data[display_cols].copy()
            display_data['Due_Date'] = display_data['Due_Date'].dt.strftime('%Y-%m-%d')
            display_data['Open_Amount'] = display_data['Open_Amount'].map('€{:,.2f}'.format)
            st.dataframe(display_data, use_container_width=True)

            # Download this vendor
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                display_data.to_excel(writer, index=False, sheet_name='Invoices')
            buffer.seek(0)
            st.download_button(
                "Download This Vendor",
                data=buffer,
                file_name=f"{show_vendor.replace(' ', '_')}_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("**Click any bar** to view invoice details.")

    # === EXPORT RAW DATA (ALL FILTERED) ===
    st.markdown("---")
    st.subheader("Export Raw Data")

    def export_excel(df_export, filename):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_export.to_excel(writer, sheet_name='Raw_Data', index=False, startrow=1, header=False)
            workbook = writer.book
            worksheet = writer.sheets['Raw_Data']
            header_fmt = workbook.add_format({
                'bold': True, 'bg_color': '#1f4e79', 'font_color': 'white',
                'border': 1, 'font_name': 'Arial', 'font_size': 11
            })
            for col_num, value in enumerate(df_export.columns):
                worksheet.write(0, col_num, value, header_fmt)
            currency_fmt = workbook.add_format({'num_format': '€#,##0.00'})
            date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            worksheet.set_column('C:C', 15, currency_fmt)
            worksheet.set_column('B:B', 12, date_fmt)
            worksheet.freeze_panes(1, 0)
        buffer.seek(0)
        return buffer

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        buf = export_excel(df, "all.xlsx")
        st.download_button("All Open", data=buf, file_name="All_Open_Invoices.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_b:
        buf = export_excel(df[df['Overdue']], "overdue.xlsx")
        st.download_button("All Overdue", data=buf, file_name="All_Overdue_Invoices.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col_c:
        buf = export_excel(df[~df['Overdue']], "not.xlsx")
        st.download_button("All Not Overdue", data=buf, file_name="All_Not_Overdue_Invoices.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

except KeyError as e:
    st.error(f"Missing column: {e}. Check data structure.")
except Exception as e:
    st.error(f"Error: {type(e).__name__}: {e}")
    st.expander("Details").write(e, unsafe_allow_html=True)
