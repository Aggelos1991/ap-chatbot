import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")

# === Session State ===
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None
if 'top_n_option' not in st.session_state:
    st.session_state.top_n_option = "Top 30"

# === File Upload ===
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        # Read specific columns including BJ (column 61 → index 60)
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()
            keep_cols = [0, 1, 4, 6, 29, 30, 31, 33, 35, 39, 55, 61]  # BJ = col 61
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB',
                                 header=None, usecols=keep_cols)

        # Find header row where column A contains "VENDOR"
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        # Assign column names
        df.columns = [
            'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
            'Vendor_Email', 'Account_Email',
            'AF', 'AH', 'AJ', 'AN', 'BD', 'BJ_Alt_Invoice'
        ]

        # === FILTER: All YES + BD keywords ===
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
            st.warning("No invoices match the filter criteria.")
            st.stop()

        # === Data Cleaning ===
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')  # Fixed: was 'cocoerce'
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No open invoices found after cleaning.")
            st.stop()

        # === Overdue Logic ===
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # === Summary by Vendor & Status ===
        full_summary = df.groupby(['Vendor_Name', 'Status'])['Open_Amount']\
                        .sum().unstack(fill_value=0).reset_index()
        for col in ['Overdue', 'Not Overdue']:
            if col not in full_summary.columns:
                full_summary[col] = 0
        full_summary['Total'] = full_summary['Overdue'] + full_summary['Not Overdue']

        # === Filters ===
        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            status_filter = st.selectbox("Show",
                ["All Open", "Overdue Only", "Not Overdue Only"], key="status")
        with col2:
            vendor_list = ["Top N"] + sorted(df['Vendor_Name'].unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendor_list, key="vendor_select")
        with col3:
            st.session_state.top_n_option = st.selectbox(
                "Show", ["Top 20", "Top 30", "All Vendors"], 
                index=["Top 20", "Top 30", "All Vendors"].index(st.session_state.top_n_option)
            )

        # === Top N Logic ===
        if st.session_state.top_n_option == "Top 20":
            top_n = 20
            title_suffix = "Top 20"
        elif st.session_state.top_n_option == "Top 30":
            top_n = 30
            title_suffix = "Top 30"
        else:
            top_n = len(full_summary)
            title_suffix = "All"

        # === Apply Status Filter + Top N ===
        if status_filter == "All Open":
            top_df = full_summary.nlargest(top_n, 'Total')
            title = f"{title_suffix} Vendors (All Open)"
        elif status_filter == "Overdue Only":
            top_df = full_summary.nlargest(top_n, 'Overdue').copy()
            top_df['Not Overdue'] = 0
            title = f"{title_suffix} Vendors (Overdue Only)"
        else:  # Not Overdue Only
            if full_summary['Not Overdue'].sum() == 0:
                st.warning("No 'Not Overdue' invoices.")
                top_df = full_summary.head(0).copy()
                top_df['Overdue'] = 0
                top_df['Not Overdue'] = 0
            else:
                top_df = full_summary.nlargest(top_n, 'Not Overdue').copy()
                top_df['Overdue'] = 0
            title = f"{title_suffix} Vendors (Not Overdue Only)"

        # === Selected Vendor Override ===
        base_df = top_df if selected_vendor == "Top N" else \
                  full_summary[full_summary['Vendor_Name'] == selected_vendor]

        # === Prepare Plot Data ===
        plot_df = base_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df = plot_df[plot_df['Amount'] > 0]
        total_per_vendor = base_df.set_index('Vendor_Name')['Total'].to_dict()
        plot_df['Total'] = plot_df['Vendor_Name'].map(total_per_vendor)

        # === Horizontal Stacked Bar Chart ===
        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title,
            color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
            height=max(500, len(plot_df) * 45),
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
            textfont=dict(size=14, color='white', family='Arial Black'),
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
            margin=dict(l=150, r=50, t=80, b=50)
        )

        chart = st.plotly_chart(fig, use_container_width=True,
                                key="vendor_chart", on_select="rerun")

        # === Click Handler: Show Raw Invoices ===
        if chart.selection and chart.selection['points']:
            point = chart.selection['points'][0]
            st.session_state.clicked_vendor = point['y']

        show_vendor = st.session_state.clicked_vendor

        if show_vendor:
            st.subheader(f"Raw Invoices: {show_vendor}")

            # Respect current status filter
            visible_vendors = base_df['Vendor_Name'].tolist()
            if status_filter == "Overdue Only":
                status_mask = df['Overdue']
            elif status_filter == "Not Overdue Only":
                status_mask = ~df['Overdue']
            else:
                status_mask = pd.Series([True] * len(df), index=df.index)

            filtered = df[df['Vendor_Name'].isin(visible_vendors) & status_mask]
            raw_details = filtered[filtered['Vendor_Name'] == show_vendor].copy()

            raw_details = raw_details[[
                'VAT_ID', 'Due_Date', 'Open_Amount',
                'BJ_Alt_Invoice', 'Status', 'Vendor_Email', 'Account_Email'
            ]]

            raw_details['Due_Date'] = raw_details['Due_Date'].dt.strftime('%Y-%m-%d')
            raw_details['Open_Amount'] = raw_details['Open_Amount'].map('€{:,.2f}'.format)

            st.dataframe(raw_details, use_container_width=True)

            # Download button for this vendor
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                raw_details.to_excel(writer, index=False, sheet_name='Raw_Data')
            buffer.seek(0)
            st.download_button(
                "Download This Vendor",
                data=buffer,
                file_name=f"{show_vendor.replace(' ', '_')}_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("**Click any bar** to see raw invoice lines.")

        # === Export All Raw Data ===
        st.markdown("---")
        st.subheader("Export Raw Data Only")

        def export_raw(raw_df  # Fixed indentation
            raw_df, filename):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                raw_df.to_excel(writer, sheet_name='Raw_Data', index=False,
                              startrow=1, header=False)
                workbook = writer.book
                worksheet = writer.sheets['Raw_Data']
                header_fmt = workbook.add_format({
                    'bold': True, 'bg_color': '#1f4e79',
                    'font_color': 'white', 'border': 1,
                    'font_name': 'Arial', 'font_size': 11
                })
                for col_num, value in enumerate(raw_df.columns):
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
            buf = export_raw(df, "all.xlsx")
            st.download_button("Download All Open", data=buf,
                               file_name="All_Open_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_b:
            buf = export_raw(df[df['Overdue']], "overdue.xlsx")
            st.download_button("Download All Overdue", data=buf,
                               file_name="All_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_c:
            buf = export_raw(df[~df['Overdue']], "not.xlsx")
            st.download_button("Download All Not Overdue", data=buf,
                               file_name="All_Not_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.stop()
else:
    st.info("Upload your Excel file to begin.")
