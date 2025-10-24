# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# Page config
st.set_page_config(page_title="Overdue Invoices App", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Upload Excel → See top vendors → Click or select → View details with emails**")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        # Read Excel (no header assumed)
        df = pd.read_excel(uploaded_file, header=None)

        # Check column count
        if df.shape[1] < 31:
            st.error("Excel must have at least 31 columns (A to AE).")
        else:
            # Map columns by position
            df.columns = [f"Col_{i}" for i in range(df.shape[1])]
            df = df.rename(columns={
                df.columns[0]: 'Vendor_Name',      # A
                df.columns[1]: 'VAT_ID',           # B
                df.columns[2]: 'Due_Date',         # C
                df.columns[6]: 'Open_Amount',      # G
                df.columns[29]: 'Vendor_Email',    # AD
                df.columns[30]: 'Account_Email'    # AE
            })

            # Keep only needed columns
            cols = ['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']
            df = df[cols].copy()

            # Data cleaning
            df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
            df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
            df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])

            # Overdue logic
            today = pd.Timestamp.today().normalize()
            df['Overdue'] = df['Due_Date'] < today
            df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

            # Filter only open items
            df = df[df['Open_Amount'] > 0]

            if df.empty:
                st.warning("No open invoices found.")
            else:
                # --- Vendor Summary ---
                vendor_summary = df.groupby('Vendor_Name').agg(
                    Total_Open=('Open_Amount', 'sum'),
                    Overdue_Amount=('Open_Amount', lambda x: x[df.loc[x.index, 'Overdue']]),
                    Not_Overdue_Amount=('Open_Amount', lambda x: x[~df.loc[x.index, 'Overdue']])
                ).reset_index()
                vendor_summary['Not_Overdue_Amount'] = vendor_summary['Total_Open'] - vendor_summary['Overdue_Amount']
                top10 = vendor_summary.nlargest(10, 'Total_Open')

                # --- Filters ---
                col1, col2 = st.columns(2)
                with col1:
                    filter_status = st.selectbox("Filter", ["All Open", "Overdue", "Not Overdue"])
                with col2:
                    vendor_options = ["Top 10"] + sorted(df['Vendor_Name'].unique().tolist())
                    selected_vendor = st.selectbox("Select Vendor", vendor_options)

                # --- Apply Status Filter ---
                plot_df = vendor_summary.copy()
                if filter_status == "Overdue":
                    plot_df['Not_Overdue_Amount'] = 0
                elif filter_status == "Not Overdue":
                    plot_df['Overdue_Amount'] = 0

                # --- Apply Vendor Selection ---
                if selected_vendor != "Top 10":
                    plot_df = plot_df[plot_df['Vendor_Name'] == selected_vendor]
                    chart_title = f"{selected_vendor} - Open Items"
                else:
                    plot_df = top10
                    chart_title = "Top 10 Vendors by Open Amount"

                # --- Prepare Data for Chart ---
                plot_df = plot_df.melt(
                    id_vars='Vendor_Name',
                    value_vars=['Overdue_Amount', 'Not_Overdue_Amount'],
                    var_name='Type',
                    value_name='Amount'
                )
                plot_df['Type'] = plot_df['Type'].map({
                    'Overdue_Amount': 'Overdue',
                    'Not_Overdue_Amount': 'Not Overdue'
                })

                # --- Interactive Bar Chart ---
                fig = px.bar(
                    plot_df,
                    x='Amount',
                    y='Vendor_Name',
                    color='Type',
                    orientation='h',
                    title=chart_title,
                    color_discrete_map={'Overdue': '#FF6B6B', 'Not Overdue': '#4ECDC4'},
                    text='Amount',
                    height=max(400, len(plot_df) * 50)
                )
                fig.update_traces(texttemplate='$%{text:,.0f}', textposition='inside')
                fig.update_layout(
                    xaxis_title="Amount ($)",
                    yaxis_title="Vendor",
                    legend_title="Status",
                    hovermode='y unified'
                )

                # Render chart
                chart = st.plotly_chart(fig, use_container_width=True, key="vendor_chart")
                click_data = st.session_state.get("vendor_chart", {}).get("clickData")

                # --- Determine Vendor for Details ---
                detail_vendor = None
                if selected_vendor != "Top 10":
                    detail_vendor = selected_vendor
                elif click_data and 'points' in click_data:
                    point = click_data['points'][0]
                    detail_vendor = point['y']

                # --- Show Details Table ---
                if detail_vendor:
                    st.subheader(f"Details: {detail_vendor}")
                    detail_df = df[df['Vendor_Name'] == detail_vendor].copy()
                    detail_df = detail_df[[
                        'VAT_ID', 'Due_Date', 'Open_Amount', 'Status',
                        'Vendor_Email', 'Account_Email'
                    ]]
                    detail_df['Due_Date'] = detail_df['Due_Date'].dt.strftime('%Y-%m-%d')
                    detail_df['Open_Amount'] = detail_df['Open_Amount'].map('$ {:,.2f}'.format)

                    st.dataframe(detail_df, use_container_width=True)

                    # --- Export Button ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        detail_df.to_excel(writer, index=False, sheet_name='Invoices')
                    output.seek(0)
                    st.download_button(
                        label="Download as Excel",
                        data=output,
                        file_name=f"{detail_vendor.replace(' ', '_')}_invoices.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("Click a bar or select a vendor to view details.")

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
else:
    st.info("Please upload your Excel file (columns A, B, C, G, AD, AE required).")
