# overdue_app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="Overdue Invoices", layout="wide")
st.title("Overdue Invoices Dashboard")
st.markdown("**Uses PivotTable13 from sheet 'Outstanding Invoices IB'**")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        # Read ONLY PivotTable13 from the correct sheet
        with pd.ExcelFile(uploaded_file) as xls:
            if 'Outstanding Invoices IB' not in xls.sheet_names:
                st.error("Sheet 'Outstanding Invoices IB' not found.")
                st.stop()
            # Read pivot table by name
            df = pd.read_excel(xls, sheet_name='Outstanding Invoices IB', 
                               skiprows=0, usecols="A:AE", 
                               engine='openpyxl')

        # Find PivotTable13 (look for header row with "Vendor Name")
        pivot_start = df[df.iloc[:, 0] == 'Vendor Name'].index
        if pivot_start.empty:
            st.error("PivotTable13 not found. Look for 'Vendor Name' in column A.")
            st.stop()
        start_row = pivot_start[0] + 1
        df = df.iloc[start_row:].copy()
        df = df.iloc[:, :31]  # A to AE

        # Map columns
        df.columns = [f"C{i}" for i in range(31)]
        df = df.rename(columns={
            'C0': 'Vendor_Name', 'C1': 'VAT_ID', 'C2': 'Due_Date',
            'C6': 'Open_Amount', 'C29': 'Vendor_Email', 'C30': 'Account_Email'
        })

        df = df[['Vendor_Name', 'VAT_ID', 'Due_Date', 'Open_Amount', 'Vendor_Email', 'Account_Email']].copy()

        # Convert
        df['Due_Date'] = pd.to_datetime(df['Due_Date'], errors='coerce')
        df['Open_Amount'] = pd.to_numeric(df['Open_Amount'], errors='coerce')
        df = df.dropna(subset=['Vendor_Name', 'Open_Amount', 'Due_Date'])
        df = df[df['Open_Amount'] > 0]

        if df.empty:
            st.warning("No open items in PivotTable13.")
            st.stop()

        # Overdue
        today = pd.Timestamp.today().normalize()
        df['Overdue'] = df['Due_Date'] < today
        df['Status'] = df['Overdue'].map({True: 'Overdue', False: 'Not Overdue'})

        # Summary
        summary = df.groupby('Vendor_Name').agg(
            Total=('Open_Amount', 'sum'),
            Overdue_Amount=('Open_Amount', lambda x: x[df.loc[x.index, 'Overdue']]),
            Not_Overdue=('Open_Amount', lambda x: x[~df.loc[x.index, 'Overdue']])
        ).reset_index()
        summary['Not_Overdue'] = summary['Total'] - summary['Overdue_Amount']
        top10 = summary.nlargest(10, 'Total')

        # Filters
        col1, col2 = st.columns(2)
        with col1:
            status = st.selectbox("Show", ["All Open", "Overdue Only", "Not Overdue Only"])
        with col2:
            vendors = ["Top 10"] + sorted(df['Vendor_Name'].unique().tolist())
            selected = st.selectbox("Select Vendor", vendors)

        # Apply
        plot = summary.copy()
        if status == "Overdue Only": plot['Not_Overdue'] = 0
        if status == "Not Overdue Only": plot['Overdue_Amount'] = 0

        if selected != "Top 10":
            plot = plot[plot['Vendor_Name'] == selected]
            title = f"{selected}"
        else:
            plot = top10
            title = "Top 10 Vendors (PivotTable13)"

        # Chart data
        plot = plot.melt(id_vars='Vendor_Name', value_vars=['Overdue_Amount', 'Not_Overdue'],
                         var_name='Type', value_name='Amount')
        plot['Type'] = plot['Type'].replace({'Overdue_Amount': 'Overdue', 'Not_Overdue': 'Not Overdue'})

        # Bar chart
        fig = px.bar(plot, x='Amount', y='Vendor_Name', color='Type', orientation='h',
                     title=title, color_discrete_map={'Overdue': '#FF6B6B', 'Not Overdue': '#4ECDC4'},
                     text='Amount', height=500)
        fig.update_traces(texttemplate='$%{text:,.0f}', textposition='inside')
        fig.update_layout(xaxis_title="Amount", yaxis_title="Vendor", legend_title="Status")

        chart = st.plotly_chart(fig, use_container_width=True, key="chart")
        click = st.session_state.get("chart", {}).get("clickData")

        show = selected if selected != "Top 10" else None
        if click: show = click['points'][0]['y']

        if show:
            st.subheader(f"Details: {show}")
            det = df[df['Vendor_Name'] == show][['VAT_ID', 'Due_Date', 'Open_Amount', 'Status', 'Vendor_Email', 'Account_Email']].copy()
            det['Due_Date'] = det['Due_Date'].dt.strftime('%Y-%m-%d')
            det['Open_Amount'] = det['Open_Amount'].map('${:,.2f}'.format)
            st.dataframe(det, use_container_width=True)

            buf = io.BytesIO()
            det.to_excel(buf, index=False, engine='xlsxwriter')
            buf.seek(0)
            st.download_button("Download Excel", data=buf,
                               file_name=f"{show}_invoices.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("Click bar or select vendor")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload your file → App reads **PivotTable13** from **Outstanding Invoices IB**")
