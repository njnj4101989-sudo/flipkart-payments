import streamlit as st
import pandas as pd
import io
import difflib
import warnings
warnings.filterwarnings('ignore')

# 🎨 Page Config
st.set_page_config(
    page_title="LAXMIPATI Sarees - Excel Processor",
    layout="wide",
    page_icon="👗",
)

# 🎨 Custom CSS
st.markdown("""
<style>
/* Header */
.header-container {
    background: linear-gradient(135deg, #e91e63, #ff9800);
    padding: 2rem;
    border-radius: 15px;
    text-align: center;
    color: white;
    box-shadow: 0px 4px 12px rgba(0,0,0,0.2);
}
.header-container h1 {
    font-size: 3rem;
    font-weight: bold;
    margin-bottom: 0.3rem;
}
.header-container p {
    font-size: 1.2rem;
    opacity: 0.9;
}

/* Style File Uploaders */
[data-testid="stFileUploadDropzone"] {
    border: 3px dashed #e91e63 !important;
    padding: 2rem;
    border-radius: 15px;
    text-align: center;
    background-color: #fff5f7;
    transition: 0.3s;
    margin-top: 20px;
}
[data-testid="stFileUploadDropzone"]:hover {
    border-color: #ff9800 !important;
    background-color: #fff0e6 !important;
}
[data-testid="stFileUploadDropzone"] label {
    font-size: 1.3rem !important;
    color: #e91e63 !important;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# 🎯 Header Section
st.markdown("""
<div class="header-container">
    <h1> LAXMIPATI Sarees </h1>
    <p>Excel Payment Processor & Data Reconciliation Tool</p>
</div>
""", unsafe_allow_html=True)

# Create tabs
tab1, tab2, tab3 = st.tabs(["📂 Upload Reports", "🔄 Reconciliation", "📊 Analytics"])

with tab1:
    # --- MAIN FILE UPLOAD ---
    uploaded_file = st.file_uploader(
        "Upload Main Excel File",
        type=["xlsx"],
        label_visibility="collapsed",
        key="main_excel_uploader"
    )

    header_row = st.sidebar.selectbox("Header Row", [0, 1, 2, 3], index=1)
    skip_rows = st.sidebar.number_input("Skip Rows (Data starts from)", 0, 10, 3)

    if uploaded_file:
        required_cols = {
            "Order ID": ["Order ID", "OrderID", "Order Id"],
            "Invoice": ["Invoice", "Invoice ID", "Invoice No"],
            "Sale Amount": ["Sale Amount", "Sale Amount (Rs.)", "Sales Amount"],
            "Payment Date": ["Payment Date", "Date", "Settlement Date"],
            "Bank Settlement Value": ["Bank Settlement Value (Rs.)", "Settlement Amount", "Net Amount"],
            "Marketplace Fee": ["Marketplace Fee", "Commission", "MP Fee"],
            "Protection Fund": ["Protection Fund", "Protection", "Fund"],
            "Refund": ["Refund", "Refund Amount", "Return Amount"],
            "Commission Rate (%)": ["Commission Rate (%)", "Commission %", "Rate of Commission"],
            "TCS (Rs.)": ["TCS (Rs.)", "TCS", "Tax Collected at Source"],
            "TDS (Rs.)": ["TDS (Rs.)", "TDS", "Tax Deducted at Source"],
            "GST on MP Fees (Rs.)": ["GST on MP Fees (Rs.)", "GST on Marketplace Fee", "GST on MP"]
        }

        # Additional columns for analysis only (not displayed in Tab 1)
        analysis_only_cols = {
            "Order Date": ["Order Date", "Order Created Date", "Created Date", "Ordered Date"],
            "Invoice Date": ["Invoice Date", "Bill Date", "Billing Date", "Invoice Created Date"],
            "Dispatch Date": ["Dispatch Date", "Shipped Date", "Shipping Date", "Dispatched Date"],
            "Return Type": ["Return Type", "Return Status", "Return Category", "Refund Type"]
        }

        try:
            temp_df = pd.read_excel(uploaded_file, sheet_name="Orders", header=None)
            column_names = temp_df.iloc[header_row].tolist()
            data_df = temp_df.iloc[skip_rows:].reset_index(drop=True)
            data_df.columns = [str(name).strip() for name in column_names[:len(data_df.columns)]]

            def match_columns(df_columns, required_cols):
                col_map = {}
                for req_col, variations in required_cols.items():
                    for variation in variations:
                        match = difflib.get_close_matches(variation, df_columns, n=1, cutoff=0.5)
                        if match:
                            col_map[req_col] = match[0]
                            break
                return col_map

            col_map = match_columns(data_df.columns, required_cols)

            # Match additional columns for analysis
            analysis_col_map = match_columns(data_df.columns, analysis_only_cols)

            if col_map:
                processed_df = data_df[list(col_map.values())].rename(columns={v: k for k, v in col_map.items()})
                for req_col in required_cols.keys():
                    if req_col not in processed_df.columns:
                        processed_df[req_col] = "NA"

                # Store in session state for analysis tab
                st.session_state.processed_df = processed_df

                # Create enhanced dataframe with analysis columns
                if analysis_col_map:
                    enhanced_df = processed_df.copy()
                    for analysis_col, original_col in analysis_col_map.items():
                        enhanced_df[analysis_col] = data_df[original_col]
                    st.session_state.analysis_df = enhanced_df
                else:
                    st.session_state.analysis_df = processed_df

                # Debug: Show what analysis columns were found
                #if analysis_col_map:
                #    st.write(f"📅 Found {len(analysis_col_map)} additional date columns for analysis: {list(analysis_col_map.keys())}")
                
                st.write("✅ Processed Main File")
                st.dataframe(processed_df)

                pivot_cols = ["Order ID", "Invoice", "Sale Amount", "Refund", "Protection Fund", 
                            "Marketplace Fee", "GST on MP Fees (Rs.)", "TCS (Rs.)", "TDS (Rs.)"]
                available_pivot_cols = [col for col in pivot_cols if col in processed_df.columns]

                if available_pivot_cols:
                    st.session_state.main_pivot = processed_df[available_pivot_cols].copy()

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    processed_df.astype(str).to_excel(writer, index=False, sheet_name="Processed_Data")

                st.download_button(
                    "⬇️ Download Processed Excel",
                    data=output.getvalue(),
                    file_name="LAXMIPATI_Processed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No matching columns found")

        except Exception as e:
            st.error(f"Error: {str(e)}")



    # 📂 PIVOT FILE UPLOAD
    pivot_file = st.file_uploader(
        "Upload Pivot Excel File",
        type=["xlsx"],
        label_visibility="collapsed",
        key="pivot_excel_uploader",
        help="Upload Excel with ZTRA, ZCN, and Claim sheets"
    )

    if pivot_file:
        try:
            ztra_df = pd.read_excel(pivot_file, sheet_name="ZTRA", dtype=str, header=1)
            zcn_df = pd.read_excel(pivot_file, sheet_name="ZCN", dtype=str, header=0)
            claim_df = pd.read_excel(pivot_file, sheet_name="Claim", dtype=str, header=0)

            ztrans_df = ztra_df.dropna(how='all')
            cn_df = zcn_df.dropna(how='all')

            ztrans_df.columns = [str(col).strip() if str(col) != 'nan' else f'Unnamed_{i}' for i, col in enumerate(ztrans_df.columns)]
            cn_df.columns = [str(col).strip() if str(col) != 'nan' else f'Unnamed_{i}' for i, col in enumerate(cn_df.columns)]

            if 'main_pivot' in st.session_state:
                main_pivot = st.session_state.main_pivot

                ztrans_grouped = ztrans_df.groupby('CUSTOMER REFERENCE').agg({
                    'Billing No': 'first',
                    'Total Amt': 'sum'
                }).reset_index()

                combined_df = main_pivot.merge(ztrans_grouped, 
                                            left_on='Invoice', 
                                            right_on='CUSTOMER REFERENCE', 
                                            how='left')

                cn_data = cn_df.groupby('Invoice Reference Number')['Total Receivable'].first().reset_index()
                final_df = combined_df.merge(cn_data, left_on='Billing No', right_on='Invoice Reference Number', how='left')

                def get_priority_claim_data(group):
                    approved_rows = group[group['STATUS-1'] == 'Approved']
                    if not approved_rows.empty:
                        return approved_rows.iloc[0]
                    else:
                        return group.iloc[0]

                claim_data = claim_df.groupby('REFERENCE NO').apply(get_priority_claim_data).reset_index(drop=True)
                claim_final = claim_data[['REFERENCE NO', 'STATUS-1', 'Approved Amount']].copy()

                final_df = final_df.merge(claim_final, left_on='Invoice', right_on='REFERENCE NO', how='left')

                final_df = final_df.rename(columns={
                    'Billing No': 'ZTRANS Invoice',
                    'Total Amt': 'ZTRANS Amount', 
                    'Total Receivable': 'ZGSTR1',
                    'STATUS-1': 'Claim Status',
                    'Approved Amount': 'Claim Approved Amt.'
                })

                final_columns = ['Order ID', 'Invoice', 'Sale Amount', 'Refund', 'Protection Fund',
                                'Marketplace Fee', 'GST on MP Fees (Rs.)', 'TCS (Rs.)', 'TDS (Rs.)',
                                'ZTRANS Invoice', 'ZTRANS Amount', 'ZGSTR1', 'Claim Status', 'Claim Approved Amt.']

                available_final_cols = [col for col in final_columns if col in final_df.columns]
                result_df = final_df[available_final_cols].fillna('NA')

                st.write("✅ Final Combined Result")
                st.dataframe(result_df)

                output_combined = io.BytesIO()
                with pd.ExcelWriter(output_combined, engine="openpyxl") as writer:
                    result_df.astype(str).to_excel(writer, index=False, sheet_name="Combined_Data")

                st.download_button(
                    "⬇️ Download Combined Data",
                    data=output_combined.getvalue(),
                    file_name="LAXMIPATI_Combined_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ Please upload and process the Main Excel file first")

        except Exception as e:
            st.error(f"Error processing Pivot file: {str(e)}")


with tab2:
    st.write("🔄 Reconciliation features coming soon...")

with tab3:
    from analysis_tab import analysis_tab
    analysis_tab()