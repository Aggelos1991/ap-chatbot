# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io
import numpy as np

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Click bar → Raw data | Export → Raw Data Only (Graph = Excel)**")

# Session state
if 'clicked_vendor' not in st.session_state:
    st.session_state.clicked_vendor = None

# Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        # Read sheet
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()
            df_raw = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', header=None)

        # Find header "VENDOR"
        header_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("VENDOR", case=False, na=False)].index
        if header_row.empty:
            st.error("Header 'VENDOR' not found in column A.")
            st.stop()

        start_row = header_row[0] + 1
        df = df_raw.iloc[start_row:].copy().reset_index(drop=True)

        if df.shape[1] < 31:
            st.error("Need columns A to AE (31 columns).")
            st.stop()

        # Map columns: A=0, B=1, E=4 (Due Date), G=6 (Open Amount), AD=29, AE=30
        df = df.iloc[:, [0, 1, 4, 6, 29, 30]].copy()
        df.columns = ['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']

        # ---- Clean & types ----
        # Force datetime then keep only the DATE part (no timezones messing the comparison)
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce').dt.date
        # Amounts
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')

        # Drop unusable rows
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No open invoices found.")
            st.stop()

        # ---- Overdue logic (DATE-ONLY, Europe/Athens) ----
        # Today as date (no time)
        today_date = pd.Timestamp.now(tz='Europe/Athens').date()

        # Overdue if due date strictly before today (i.e., already past due)
        df['Overdue'] = df['Due_Date'] < today_date
        df['Status'] = np.where(df['Overdue'], 'Overdue', 'Not Overdue')

        # ---- Summary by vendor ----
        full_summary = (
            df.groupby(['Vendor_Name', 'Status'], as_index=False)['Open_Amount']
              .sum()
              .pivot(index='Vendor_Name', columns='Status', values='Open_Amount')
              .fillna(0)
              .reset_index()
        )
        # Ensure both columns exist even if one category is empty
        if 'Overdue' not in full_summary.columns: full_summary['Overdue'] = 0
        if 'Not Overdue' not in full_summary.columns: full_summary['Not Overdue'] = 0
        full_summary['Total'] = full_summary['Overdue'] + full_summary['Not Overdue']

        # ---- Filters ----
        col1, col2, col3 = st.columns(3)
        with col1:
            status_filter = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"], key="status")
        with col2:
            vendor_list = ["Top 20", "Top 30"] + sorted(df['Vendor_Name'].unique().tolist())
            selected_vendor = st.selectbox("Select Vendor", vendor_list, key="vendor_select")
        with col3:
            st.caption(f"Today (Athens): {today_date}")

        # ---- Top vendors ----
        top_n = 30 if "30" in selected_vendor else 20
        if status_filter == "All Open":
            top_df = full_summary.nlargest(top_n, 'Total')
            title = f"Top {top_n} Vendors (All Open)"
        elif status_filter == "Overdue Only":
            top_df = full_summary.nlargest(top_n, 'Overdue').copy()
            top_df['Not Overdue'] = 0
            title = f"Top {top_n} Vendors (Overdue Only)"
        else:
            top_df = full_summary.nlargest(top_n, 'Not Overdue').copy()
            top_df['Overdue'] = 0
            title = f"Top {top_n} Vendors (Not Overdue Only)"

        # Base data (Top list or single vendor)
        base_df = top_df if selected_vendor in ["Top 20", "Top 30"] else full_summary[full_summary['Vendor_Name'] == selected_vendor]

        # ---- Chart data ----
        plot_df = base_df.melt(
            id_vars='Vendor_Name',
            value_vars=['Overdue', 'Not Overdue'],
            var_name='Type',
            value_name='Amount'
        )
        plot_df = plot_df[plot_df['Amount'] > 0]

        total_per_vendor = base_df.set_index('Vendor_Name')['Total'].to_dict()
        plot_df['Total'] = plot_df['Vendor_Name'].map(total_per_vendor)

        # ---- Chart ----
        fig = px.bar(
            plot_df,
            x='Amount',
            y='Vendor_Name',
            color='Type',
            orientation='h',
            title=title,
            color_discrete_map={'Overdue': '#8B0000', 'Not Overdue': '#4682B4'},
            height=max(500, len(plot_df) * 45),
        )

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

        st.plotly_chart(fig, use_container_width=True)

        # ======= EMAIL SECTION =======
        st.markdown("---")
        st.subheader("📧 Email Addresses (copy for Outlook)")

        if status_filter == "All Open":
            scope = df
        elif status_filter == "Overdue Only":
            scope = df[df['Overdue']]
        else:
            scope = df[~df['Overdue']]

        emails = pd.concat([scope['Vendor_Email'], scope['Account_Email']], ignore_index=True)
        emails = emails.dropna().astype(str).str.strip()
        emails = emails[emails.str.contains('@')]
        unique_emails = sorted(set(emails.tolist()))
        email_text = ", ".join(unique_emails)

        if unique_emails:
            st.text_area("All relevant emails (Ctrl+C to copy):", email_text, height=120)
            st.success(f"📋 {len(unique_emails)} unique emails found.")
        else:
            st.info("No emails found in this category.")

        # ======= RAW DATA SECTION =======
        st.markdown("---")
        st.subheader("Export Raw Data Only")

        def export_raw(raw_df):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Ensure proper types for export
                out_df = raw_df.copy()
                # Reformat for export view
                out_df = out_df[['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Vendor_Email', 'Account_Email']]
                out_df.to_excel(writer, sheet_name='Raw_Data', index=False, startrow=1, header=False)
                workbook = writer.book
                worksheet = writer.sheets['Raw_Data']

                # Header
                header_fmt = workbook.add_format({'bold': True,'bg_color': '#1f4e79','font_color': 'white','border': 1,'font_name': 'Arial','font_size': 11})
                for col_num, value in enumerate(out_df.columns):
                    worksheet.write(0, col_num, value, header_fmt)

                # Formats
                currency_fmt = workbook.add_format({'num_format': '€#,##0.00'})
                date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})
                worksheet.set_column('C:C', 12, date_fmt)       # Due_Date
                worksheet.set_column('D:D', 15, currency_fmt)   # Open_Amount
                worksheet.freeze_panes(1, 0)
            buffer.seek(0)
            return buffer

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.download_button("Download All Open", data=export_raw(df), file_name="All_Open_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_b:
            st.download_button("Download All Overdue", data=export_raw(df[df['Overdue']]), file_name="All_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_c:
            st.download_button("Download All Not Overdue", data=export_raw(df[~df['Overdue']]), file_name="All_Not_Overdue_Raw.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.stop()
else:
    st.info("Upload your Excel file → Click bar → See raw data → Export Raw")
