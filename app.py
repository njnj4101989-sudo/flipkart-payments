import streamlit as st
import pandas as pd
import io
import difflib
import warnings
warnings.filterwarnings('ignore')

# Page setup
st.set_page_config(page_title="LAXMIPATI Sarees - Excel Processor", layout="wide")
st.markdown("<h1 style='text-align: center; color: darkred;'>LAXMIPATI Sarees</h1>", unsafe_allow_html=True)
st.write("### Excel Payment Processor")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Sidebar settings
header_row = st.sidebar.selectbox("Header Row", [0, 1, 2, 3], index=1)
skip_rows = st.sidebar.number_input("Skip Rows (Data starts from)", 0, 10, 3)

if uploaded_file:
    # Required columns to extract
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
    
    try:
        # Read Excel with specified header and skip rows
        temp_df = pd.read_excel(uploaded_file, sheet_name="Orders", header=None)
        
        # Get first sheet
        #sheet_name = list(temp_df.keys())[0]
        #temp_data = temp_df[sheet_name]
        
        # Extract column names from header row
        column_names = temp_df.iloc[header_row].tolist()
        
        # Extract data starting from skip_rows
        data_df = temp_df.iloc[skip_rows:].reset_index(drop=True)
        data_df.columns = [str(name).strip() for name in column_names[:len(data_df.columns)]]
        
        # Match columns using fuzzy matching
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
        
        if col_map:
            # Extract matched columns
            processed_df = data_df[list(col_map.values())].rename(columns={v: k for k, v in col_map.items()})
            
            # Add missing columns as NA
            for req_col in required_cols.keys():
                if req_col not in processed_df.columns:
                    processed_df[req_col] = "NA"
            
            # Display results
            st.success(f"Found {len(col_map)} matching columns")
            st.dataframe(processed_df)

            # Create pivot table from processed data for later use
            pivot_cols = ["Order ID", "Invoice", "Sale Amount", "Refund", "Protection Fund", 
                        "Marketplace Fee", "GST on MP Fees (Rs.)", "TCS (Rs.)", "TDS (Rs.)"]
            available_pivot_cols = [col for col in pivot_cols if col in processed_df.columns]

            if available_pivot_cols:
                main_pivot = processed_df[available_pivot_cols].copy()
                # Store in session state for second uploader
                st.session_state.main_pivot = main_pivot
    
                st.write("#### Pivot Table from Main Data")
                st.dataframe(main_pivot)
            
            # Download option
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                # Convert all columns to string to avoid PyArrow issues
                download_df = processed_df.astype(str)
                download_df.to_excel(writer, index=False, sheet_name="Processed_Data")
            
            st.download_button(
                "Download Processed Excel",
                data=output.getvalue(),
                file_name="LAXMIPATI_Processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No matching columns found")
            
    except Exception as e:
        st.error(f"Error: {str(e)}")

# Second file uploader
pivot_file = st.file_uploader("Upload Excel File for Pivot Table", type=["xlsx"], key="pivot_uploader")

if pivot_file:
    try:
        # Read specific sheets: ZTRANS and CN with dtype=str to avoid type issues
        # Read sheets with different header rows
        ztra_df = pd.read_excel(pivot_file, sheet_name="ZTRA", dtype=str, header=1)  # Row 1 for ZTRA
        zcn_df = pd.read_excel(pivot_file, sheet_name="ZCN", dtype=str, header=0)   # Row 0 for ZCN

        ztrans_df = ztra_df  # Keep the same variable name
        cn_df = zcn_df       # Keep the same variable name
        
        # Clean data thoroughly
        ztrans_df = ztrans_df.dropna(how='all')  # Remove completely empty rows
        cn_df = cn_df.dropna(how='all')

        # Clean column names - remove extra spaces and handle NaN columns
        ztrans_df.columns = [str(col).strip() if str(col) != 'nan' else f'Unnamed_{i}' 
                            for i, col in enumerate(ztrans_df.columns)]
        cn_df.columns = [str(col).strip() if str(col) != 'nan' else f'Unnamed_{i}' 
                        for i, col in enumerate(cn_df.columns)]

        # Remove rows with all NaN/empty values
        ztrans_df = ztrans_df.replace('nan', '').replace('', None).dropna(how='all')
        cn_df = cn_df.replace('nan', '').replace('', None).dropna(how='all')
        
        st.write("#### ZTRA Data Preview")
        st.write("**Available columns:**", list(ztrans_df.columns))
        st.dataframe(ztrans_df.head().fillna(''))
        
        st.write("#### ZCN Data Preview") 
        st.write("**Available columns:**", list(cn_df.columns))
        st.dataframe(cn_df.head().fillna(''))
        
        # Check if main pivot exists from first upload
        if 'main_pivot' in st.session_state:
            st.write("#### Processing Combined Data...")
            # Combining logic will go here in next step
            # Get main pivot data
            main_pivot = st.session_state.main_pivot
                        
            # Group ZTRANS data by CUSTOMER REFERENCE
            ztrans_grouped = ztrans_df.groupby('CUSTOMER REFERENCE').agg({
                'Billing No': 'first',  # Take first billing number for each customer reference
                'Total Amt': 'sum'      # Sum all amounts for each customer reference
            }).reset_index()

            # Merge main pivot with ZTRANS data on Invoice = CUSTOMER REFERENCE
            combined_df = main_pivot.merge(ztrans_grouped, 
                                        left_on='Invoice', 
                                        right_on='CUSTOMER REFERENCE', 
                                        how='left')

            # Get Total Receivable from CN sheet using Billing No
            cn_data = cn_df.groupby('Invoice Reference Number')['Total Receivable'].first().reset_index()

            # Merge with CN data
            final_df = combined_df.merge(cn_data,
                                        left_on='Billing No',
                                        right_on='Invoice Reference Number',
                                        how='left')

            # Rename and organize columns as required
            final_df = final_df.rename(columns={
                'Billing No': 'ZTRANS Invoice',
                'Total Amt': 'ZTRANS Amount', 
                'Total Receivable': 'ZGSTR1'
            })

            # Select and reorder final columns
            final_columns = ['Order ID', 'Invoice', 'Sale Amount', 'Refund', 'Protection Fund',
                            'Marketplace Fee', 'GST on MP Fees (Rs.)', 'TCS (Rs.)', 'TDS (Rs.)',
                            'ZTRANS Invoice', 'ZTRANS Amount', 'ZGSTR1']

            # Keep only available columns and fill missing with NA
            available_final_cols = [col for col in final_columns if col in final_df.columns]
            result_df = final_df[available_final_cols].fillna('NA')

            st.write("#### Final Combined Result")
            st.dataframe(result_df)
            
        else:
            st.warning("Please upload and process the main Excel file first")
            
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")        
