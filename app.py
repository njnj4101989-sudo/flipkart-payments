import streamlit as st
import pandas as pd
import io
import difflib

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
        "Refund": ["Refund", "Refund Amount", "Return Amount"]
    }
    
    try:
        # Read Excel with specified header and skip rows
        temp_df = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        
        # Get first sheet
        sheet_name = list(temp_df.keys())[0]
        temp_data = temp_df[sheet_name]
        
        # Extract column names from header row
        column_names = temp_data.iloc[header_row].tolist()
        
        # Extract data starting from skip_rows
        data_df = temp_data.iloc[skip_rows:].reset_index(drop=True)
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
            
            # Download option
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                processed_df.to_excel(writer, index=False, sheet_name="Processed_Data")
            
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