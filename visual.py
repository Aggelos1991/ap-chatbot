# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices – Priority Vendors Dashboard")
st.markdown("**Click a bar → See only that vendor’s raw data | Export → Filtered Excel**")

# Session state
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None

# Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()

            # READ ALL NEEDED COLUMNS (including R = index 17)
            keep_cols = [0, 1, 4, 6, 17, 29, 30, 31, 33, 35, 39, 55]  # A,B,E,G,R,AD,AE,AF,AH,AJ,AN,BD
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None, usecols=keep_cols)

        # Find header row
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        # Assign column names
        df.columns = [
            'Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount',
            'Alt_Document', 'Vendor_Email', 'Account_Email',
            'AF', 'AH', 'AJ', 'AN', 'BD'
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
            st.warning("No invoices match the priority filter.")
            st.stop()

        # Clean data
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No valid open invoices after cleaning.")
            st.stop()

        # Overdue logic
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # === SAFE SUMMARY ===
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
        full_summary = summary

        # === FILTERS ===
        col1, col2 = st.columns(2)
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"], key="status")
        with col2:
            vendor_list = ["Top 20"] + sorted(df['Vendor_Name'].unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendor_list, key="vendor_select")

        # === SAFE TOP 20 ===
        if status_filter == "All Open":
            top_df = full_summary.nlargest(20, 'Total').copy()
            title = "Top 20 Vendors (All Open)"
        elif status_filter == "Overdue Only":
            top_df = full_summary.nlargest(20, 'Overdue').copy()
            top_df['Not Overdue'] = 0
            title = "Top 20 Vendors (Overdue Only)"
        else:
            top_df = full_summary.nlargest(20, 'Not Overdue').copy()
            top_df['Overdue'] = 0
            title = "Top 20 Vendors (Not Overdue Only)"

        base_df = top_df if selected_vendor == "Top 20" else full_summary[full_summary['Vendor_Name'] == selected_vendor]

        # === PLOT DATA ===
        plot_df = base_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df = plot_df[plot_df['Amount'] > 0]

        # === BAR CHART ===
        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title,
            color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
            height=max(500, len(plot_df) * 45)
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
            margin=dict(l=160, r=50, t=80, b=50)
        )

        # === INTERACTIVE CHART ===
        chart = st.plotly_chart(fig, use_container_width=True, key="vendor_chart", on_select="rerun")

        # === CAPTURE CLICK ===
        if chart.selection and chart.selection['points']:
            clicked = chart.selection['points'][0]['y']
            st.session_state.clicked_vendor = clicked
        else:
            # Reset if click outside
            if st.session_state.clicked_vendor and st.session_state.clicked_vendor not in plot_df['Vendor_Name'].values:
                st.session_state.clicked_vendor = None

        # === SHOW RAW DATA (ONLY CLICKED VENDOR) ===
        show_vendor = st.session_state.clicked_vendor
        if show_vendor:
            st.markdown("---")
            st.subheader(f"Raw Data: **{show_vendor}**")

            raw_details = df[df['Vendor_Name'] == show_vendor].copy()
            raw_details = raw_details[[
                'VAT_ID', 'Due_Date', 'Open_Amount', 'Status',
                'Alt_Document', 'Vendor_Email', 'Account_Email'
            ]]

            raw_details['Due_Date'] = raw_details['Due_Date'].dt.strftime('%Y-%m-%d')
            raw_details['Open_Amount'] = raw_details['Open_Amount'].map('€{:,.2f}'.format)

            st.dataframe(raw_details, use_container_width=True)

            # Download this vendor
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                raw_details.to_excel(writer, index=False, sheet_name='Raw_Data')
            buffer.seek(0)
            st.download_button(
                "Download This Vendor",
                data=buffer,
                file_name=f"{show_vendor.replace(' ', '_')}_raw.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("**Click any bar** to view raw invoice details (including Alternative Document).")

        # === EXPORT ALL FILTERED DATA ===
        st.markdown("---")
        st.subheader("Export Filtered Raw Data")

        def export_raw(data_df):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                data_df.to_excel(writer, sheet_name='Raw_Data', index=False, startrow=1, header=False)
                workbook = writer.book
                worksheet = writer.sheets['Raw_Data']

                # Header
                header_fmt = workbook.add_format({
                    'bold': True, 'bg_color': '#1f4e79', 'font_color': 'white',
                    'border': 1, 'font_name': 'Arial', 'font_size': 11
                })
                for i, col in enumerate(data_df.columns):
                    worksheet.write(0, i, col, header_fmt)

                # Formats
                currency = workbook.add_format({'num_format': '€#,##0.00'})
                date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                worksheet.set_column('C:C', 15, currency)  # Open_Amount
                worksheet.set_column('B:B', 12, date_fmt)  # Due_Date
                worksheet.freeze_panes(1, 0)
            buffer.seek(0)
            return buffer

        export_df = df[['VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Alt_Document', 'Vendor_Email', 'Account_Email']].copy()
        export_df['Due_Date'] = export_df['Due_Date'].dt.strftime('%Y-%m-%d')
        export_df['Open_Amount'] = export_df['Open_Amount'].map('€{:,.2f}'.format)

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            buf = export_raw(export_df)
            st.download_button("All Open", data=buf, file_name="All_Open_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_b:
            buf = export_raw(export_df[df['Status'] == 'Overdue'])
            st.download_button("All Overdue", data=buf, file_name="All_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_c:
            buf = export_raw(export_df[df['Status'] == 'Not Overdue'])
            st.download_button("All Not Overdue", data=buf, file_name="All_Not_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {str(e)}")
        st.stop()
else:
    st.info("Upload your Excel file → Click bar → View raw data with **Alternative Document** → Export")
