# LAXMIPATI Sarees - Excel Processor

A Streamlit app to process Flipkart payment Excel files with intelligent column matching.

## Features
- Smart column matching using fuzzy logic
- Handles multiple Excel formats
- Extracts 8 key payment columns
- Clean data output for analysis

## Usage
1. Install dependencies: `pip install -r requirements.txt`
2. Run app: `streamlit run app.py`
3. Upload Excel file and configure settings
4. Download processed data

## Required Columns
- Order ID, Invoice, Sale Amount, Payment Date
- Bank Settlement Value, Marketplace Fee, Protection Fund, Refund