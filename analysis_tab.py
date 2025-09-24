import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# Set style for better looking plots
plt.style.use('default')
sns.set_palette("husl")

def analysis_tab():
    st.header("📊 Analysis & Insights")
    
    # Check if we have processed_df from main app
    if 'analysis_df' not in st.session_state and 'processed_df' not in st.session_state:
        st.warning("⚠️ No processed data found. Please go to 'Upload Reports' tab and process your main Excel file first.")
        st.stop()
    
    # Get the enhanced dataframe with additional date columns
    df = st.session_state.analysis_df.copy() if 'analysis_df' in st.session_state else st.session_state.processed_df.copy()
    
    st.success(f"✅ Data loaded successfully! Found {len(df)} rows and {len(df.columns)} columns.")
    
    # Show which columns are available for analysis
    #if 'analysis_df' in st.session_state:
    #    date_columns_found = [col for col in ['Order Date', 'Invoice Date', 'Dispatch Date', 'Return Type'] if col in df.columns and not df[col].isna().all()]
    #    if date_columns_found:
    #        st.info(f"📅 Additional analysis columns found: {', '.join(date_columns_found)}")
    
    # Data preprocessing for analysis
    def clean_numeric_column(series, column_name):
        """Clean and convert column to numeric"""
        try:
            if series.dtype == 'object':
                cleaned = series.astype(str).str.replace(r'[₹,\s]', '', regex=True)
                cleaned = cleaned.replace(['NA', 'nan', '', 'None'], '0')
                return pd.to_numeric(cleaned, errors='coerce').fillna(0)
            else:
                return pd.to_numeric(series, errors='coerce').fillna(0)
        except:
            st.warning(f"⚠️ Could not convert {column_name} to numeric. Using zeros.")
            return pd.Series([0] * len(series))
    
    def parse_date_column(series, column_name):
        """Parse date column with multiple formats"""
        try:
            if series.isna().all():
                return pd.Series([pd.NaT] * len(series))
            
            # Try different date formats
            parsed_dates = pd.to_datetime(series, errors='coerce', dayfirst=True)
            if parsed_dates.isna().all():
                parsed_dates = pd.to_datetime(series, format='%d-%m-%Y', errors='coerce')
            if parsed_dates.isna().all():
                parsed_dates = pd.to_datetime(series, format='%Y-%m-%d', errors='coerce')
            if parsed_dates.isna().all():
                parsed_dates = pd.to_datetime(series, format='%d/%m/%Y', errors='coerce')
            
            return parsed_dates
        except:
            st.warning(f"⚠️ Could not parse {column_name} as dates.")
            return pd.Series([pd.NaT] * len(series))
    
    # Clean numeric columns
    numeric_columns = ['Sale Amount', 'Bank Settlement Value', 'Marketplace Fee', 
                      'Protection Fund', 'Refund', 'TCS (Rs.)', 'TDS (Rs.)', 'GST on MP Fees (Rs.)']
    
    for col in numeric_columns:
        if col in df.columns:
            df[f'{col}_clean'] = clean_numeric_column(df[col], col)
    
    # Clean all date columns
    date_columns = ['Payment Date', 'Order Date', 'Invoice Date', 'Dispatch Date']
    for col in date_columns:
        if col in df.columns:
            df[f'{col}_clean'] = parse_date_column(df[col], col)
    
    # Create analysis sections
    st.write("---")
    
    # 🎯 Sales & Payments Analysis
    st.subheader("💰 Sales & Payments Analysis")
    
    # Check if we have the required columns
    has_sale_amount = 'Sale Amount_clean' in df.columns
    has_settlement = 'Bank Settlement Value_clean' in df.columns
    has_payment_date = 'Payment Date_clean' in df.columns
    has_order_date = 'Order Date_clean' in df.columns and not df['Order Date_clean'].isna().all()
    has_invoice_date = 'Invoice Date_clean' in df.columns and not df['Invoice Date_clean'].isna().all()
    has_dispatch_date = 'Dispatch Date_clean' in df.columns and not df['Dispatch Date_clean'].isna().all()
    
    if has_sale_amount and has_settlement:
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Key Metrics
            st.markdown("### 📈 Key Metrics")
            
            total_sales = df['Sale Amount_clean'].sum()
            total_settlement = df['Bank Settlement Value_clean'].sum()
            total_deductions = total_sales - total_settlement
            deduction_rate = (total_deductions / total_sales * 100) if total_sales > 0 else 0
            
            st.metric("Total Sale Amount", f"₹{total_sales:,.2f}")
            st.metric("Total Bank Settlement", f"₹{total_settlement:,.2f}")
            st.metric("Total Deductions", f"₹{total_deductions:,.2f}", f"{deduction_rate:.1f}%")
            
            # Average order value
            avg_order_value = df['Sale Amount_clean'].mean()
            st.metric("Average Order Value", f"₹{avg_order_value:,.2f}")
        
        with col2:
            # Deduction Breakdown
            st.markdown("### 🔍 Deduction Breakdown")
            
            deduction_data = {}
            if 'Marketplace Fee_clean' in df.columns:
                deduction_data['Marketplace Fee'] = df['Marketplace Fee_clean'].sum()
            if 'Protection Fund_clean' in df.columns:
                deduction_data['Protection Fund'] = df['Protection Fund_clean'].sum()
            if 'TCS (Rs.)_clean' in df.columns:
                deduction_data['TCS'] = df['TCS (Rs.)_clean'].sum()
            if 'TDS (Rs.)_clean' in df.columns:
                deduction_data['TDS'] = df['TDS (Rs.)_clean'].sum()
            if 'GST on MP Fees (Rs.)_clean' in df.columns:
                deduction_data['GST on MP Fees'] = df['GST on MP Fees (Rs.)_clean'].sum()
            if 'Refund_clean' in df.columns:
                deduction_data['Refunds'] = df['Refund_clean'].sum()
            
            for fee_type, amount in deduction_data.items():
                percentage = (amount / total_sales * 100) if total_sales > 0 else 0
                st.metric(fee_type, f"₹{amount:,.2f}", f"{percentage:.1f}%")
        
        # Sales vs Settlement Comparison Charts
        st.markdown("### 📊 Sales vs Bank Settlement Analysis")
        
        # Create 2x2 subplot layout
        fig = plt.figure(figsize=(15, 12))
        gs = fig.add_gridspec(2, 2, hspace=0.3, wspace=0.25)
        ax1 = fig.add_subplot(gs[0, 0])
        ax2 = fig.add_subplot(gs[0, 1]) 
        ax3 = fig.add_subplot(gs[1, 0])
        ax4 = fig.add_subplot(gs[1, 1])
        fig.suptitle('LAXMIPATI Sarees - Sales & Payment Analysis', fontsize=16, fontweight='bold')
        
        # Chart 1: Sales vs Settlement Overview
        categories = ['Total Sales', 'Bank Settlement', 'Total Deductions']
        values = [total_sales, total_settlement, total_deductions]
        colors = ['#2E8B57', '#4169E1', '#DC143C']
        
        bars = ax1.bar(categories, values, color=colors, alpha=0.8)
        ax1.set_title('Sales vs Settlement Overview', fontweight='bold')
        ax1.set_ylabel('Amount (₹)')
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'₹{x/1000:.0f}K' if x >= 1000 else f'₹{x:.0f}'))
        
        # Add value labels on bars
        for bar, value in zip(bars, values):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                    f'₹{value:,.0f}', ha='center', va='bottom', fontweight='bold')
        
        # Chart 2: Deduction Breakdown Bar Chart
        if deduction_data and any(deduction_data.values()):
            # Get all deductions with values > 0
            deduction_items = [(k, v) for k, v in deduction_data.items() if v > 0]
            
            if deduction_items:
                deduction_names = [item[0] for item in deduction_items]
                deduction_values = [item[1] for item in deduction_items]
                
                bars = ax2.bar(range(len(deduction_names)), deduction_values, 
                            color=['#FF6B35', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD'][:len(deduction_names)])
                
                # Fix x-axis
                ax2.set_xticks(range(len(deduction_names)))
                ax2.set_xticklabels([name.replace(' ', '\n') for name in deduction_names], 
                                fontsize=8, rotation=0)
                ax2.set_xlim(-0.5, len(deduction_names) - 0.5)
                
                # Format y-axis for currency
                ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'₹{x/1000:.0f}K' if x >= 1000 else f'₹{x:.0f}'))
                
                # Labels and title
                ax2.set_title('Deduction Breakdown', fontweight='bold')
                ax2.set_ylabel('Amount (₹)')
                
                # Add value labels on bars
                for bar, value in zip(bars, deduction_values):
                    height = bar.get_height()
                    ax2.text(bar.get_x() + bar.get_width()/2., height,
                            f'₹{value:,.0f}', ha='center', va='bottom', 
                            fontweight='bold', fontsize=8)
            else:
                ax2.text(0.5, 0.5, 'No deductions data', ha='center', va='center', transform=ax2.transAxes)
                ax2.set_title('Deduction Breakdown', fontweight='bold')
        else:
            ax2.text(0.5, 0.5, 'No deductions data', ha='center', va='center', transform=ax2.transAxes)
            ax2.set_title('Deduction Breakdown', fontweight='bold')

        # Chart 3: Profit Margin Analysis - Pie Chart with Better Labels
        if 'Sale Amount_clean' in df.columns and 'Bank Settlement Value_clean' in df.columns:
            # Calculate profit margins for each order
            sample_df = df.sample(min(100, len(df)))
            sample_df_analysis = sample_df.copy()
            sample_df_analysis['Profit_Margin'] = ((sample_df_analysis['Bank Settlement Value_clean'] / 
                                                sample_df_analysis['Sale Amount_clean']) * 100).fillna(0)
            
            # Create profit margin bins with descriptive labels
            bins = [0, 50, 60, 70, 80, 90, 100]
            labels = ['Poor (0-50%)', 'Low (50-60%)', 'Fair (60-70%)', 'Good (70-80%)', 'Very Good (80-90%)', 'Excellent (90-100%)']
            
            sample_df_analysis['Margin_Category'] = pd.cut(sample_df_analysis['Profit_Margin'], 
                                                        bins=bins, labels=labels, include_lowest=True)
            
            margin_counts = sample_df_analysis['Margin_Category'].value_counts().sort_index()
            
            # Filter out zero values for cleaner pie chart
            margin_counts = margin_counts[margin_counts > 0]
            
            if len(margin_counts) > 0:
                colors = ['#DC3545', '#FF6B35', '#FFC107', '#28A745', '#17A2B8', '#6F42C1'][:len(margin_counts)]
                
                # Create custom labels with both category and percentage
                def make_autopct(values):
                    def my_autopct(pct):
                        total = sum(values)
                        val = int(round(pct*total/100.0))
                        return f'{pct:.1f}%\n({val} orders)'
                    return my_autopct
                
                wedges, texts, autotexts = ax3.pie(margin_counts.values, 
                                                labels=margin_counts.index,
                                                autopct=make_autopct(margin_counts.values),
                                                startangle=90,
                                                colors=colors,
                                                explode=[0.05] * len(margin_counts),
                                                shadow=True)
                
                ax3.set_title('Settlement Efficiency Analysis\n(How much of sale amount is actually received)', 
                            fontweight='bold', fontsize=12)
                
                # Style the text
                for autotext in autotexts:
                    autotext.set_color('white')
                    autotext.set_fontweight('bold')
                    autotext.set_fontsize(8)
                
                for text in texts:
                    text.set_fontsize(9)
                    text.set_fontweight('bold')
                
            else:
                ax3.text(0.5, 0.5, 'No data for analysis', ha='center', va='center', 
                        transform=ax3.transAxes, fontweight='bold')
                ax3.set_title('Settlement Efficiency Analysis', fontweight='bold')
        else:
            ax3.text(0.5, 0.5, 'Required data not available', ha='center', va='center', 
                    transform=ax3.transAxes, fontweight='bold')
            ax3.set_title('Settlement Efficiency Analysis', fontweight='bold')
        
        # Chart 4: Monthly trend (prefer Order Date, fallback to Payment Date)
        primary_date_col = None
        if has_order_date:
            primary_date_col = 'Order Date_clean'
            date_label = 'Order Date'
        elif has_payment_date:
            primary_date_col = 'Payment Date_clean'
            date_label = 'Payment Date'
        
        if primary_date_col:
            df_with_dates = df[df[primary_date_col].notna()].copy()
            if len(df_with_dates) > 0:
                monthly_data = df_with_dates.groupby(df_with_dates[primary_date_col].dt.to_period('M')).agg({
                    'Sale Amount_clean': 'sum',
                    'Bank Settlement Value_clean': 'sum'
                }).reset_index()
                monthly_data['Month'] = monthly_data[primary_date_col].astype(str)
                
                x_pos = np.arange(len(monthly_data))
                width = 0.35
                
                ax4.bar(x_pos - width/2, monthly_data['Sale Amount_clean'], 
                       width, label='Sales', color='#2E8B57', alpha=0.8)
                ax4.bar(x_pos + width/2, monthly_data['Bank Settlement Value_clean'], 
                       width, label='Settlement', color='#4169E1', alpha=0.8)
                
                ax4.set_xlabel('Month')
                ax4.set_ylabel('Amount (₹)')
                ax4.set_title(f'Monthly Trend (by {date_label})', fontweight='bold')
                ax4.set_xticks(x_pos)
                ax4.set_xticklabels(monthly_data['Month'], rotation=45)
                ax4.legend()
                ax4.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'₹{x/1000:.0f}K' if x >= 1000 else f'₹{x:.0f}'))
        else:
            ax4.text(0.5, 0.5, 'No date data available', ha='center', va='center', transform=ax4.transAxes)
            ax4.set_title('Monthly Trend', fontweight='bold')
        
        plt.tight_layout()
        st.pyplot(fig)
        
        # Enhanced Time Series Analysis - Simplified
        st.markdown("### 📈 Sales Trend Over Time")
        
        # Determine best date column to use
        primary_date_col = None
        date_label = None
        
        if has_order_date:
            primary_date_col = 'Order Date_clean'
            date_label = 'Order Date'
        elif has_payment_date:
            primary_date_col = 'Payment Date_clean' 
            date_label = 'Payment Date'
        elif has_invoice_date:
            primary_date_col = 'Invoice Date_clean'
            date_label = 'Invoice Date'
        
        if primary_date_col:
            df_time = df[df[primary_date_col].notna()].copy()
            
            if len(df_time) > 0:
                # Daily trend analysis
                daily_data = df_time.groupby(df_time[primary_date_col].dt.date).agg({
                    'Sale Amount_clean': 'sum',
                    'Bank Settlement Value_clean': 'sum',
                    'Order ID': 'count'
                }).reset_index()
                daily_data.columns = ['Date', 'Sales', 'Settlement', 'Orders']
                
                # Main trend chart
                fig_trend, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
                
                # Sales and settlement trend
                ax1.plot(daily_data['Date'], daily_data['Sales'], 
                        marker='o', linewidth=2, markersize=4, color='#2E8B57', label='Sales')
                ax1.plot(daily_data['Date'], daily_data['Settlement'], 
                        marker='s', linewidth=2, markersize=4, color='#4169E1', label='Settlement')
                
                ax1.set_title(f'Daily Sales & Settlement Trend (by {date_label})', fontweight='bold')
                ax1.set_ylabel('Amount (₹)')
                ax1.legend()
                ax1.grid(True, alpha=0.3)
                ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'₹{x/1000:.0f}K' if x >= 1000 else f'₹{x:.0f}'))
                
                # Order count trend
                ax2.plot(daily_data['Date'], daily_data['Orders'], 
                        marker='^', linewidth=2, markersize=4, color='#FF6B6B')
                ax2.set_title('Daily Order Count', fontweight='bold')
                ax2.set_xlabel('Date')
                ax2.set_ylabel('Number of Orders')
                ax2.grid(True, alpha=0.3)
                
                plt.xticks(rotation=45)
                plt.tight_layout()
                st.pyplot(fig_trend)
                
                # Weekly and Monthly analysis
                col1, col2 = st.columns(2)
                
                with col1:
                    # Weekly pattern
                    df_time['Weekday'] = df_time[primary_date_col].dt.day_name()
                    weekday_sales = df_time.groupby('Weekday')['Sale Amount_clean'].sum().reset_index()
                    weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                    weekday_sales['Weekday'] = pd.Categorical(weekday_sales['Weekday'], categories=weekday_order, ordered=True)
                    weekday_sales = weekday_sales.sort_values('Weekday')
                    
                    fig_week, ax = plt.subplots(figsize=(8, 6))
                    bars = ax.bar(weekday_sales['Weekday'], weekday_sales['Sale Amount_clean'], 
                                 color=sns.color_palette("viridis", len(weekday_sales)))
                    ax.set_title('Sales by Day of Week', fontweight='bold')
                    ax.set_ylabel('Sales Amount (₹)')
                    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'₹{x/1000:.0f}K' if x >= 1000 else f'₹{x:.0f}'))
                    plt.xticks(rotation=45)
                    plt.tight_layout()
                    st.pyplot(fig_week)
                
                with col2:
                    # Monthly trend
                    df_time['Month'] = df_time[primary_date_col].dt.strftime('%Y-%m')
                    monthly_sales = df_time.groupby('Month')['Sale Amount_clean'].sum().reset_index()
                    
                    fig_month, ax = plt.subplots(figsize=(8, 6))
                    ax.plot(monthly_sales['Month'], monthly_sales['Sale Amount_clean'], 
                           marker='o', linewidth=3, markersize=6, color='#FF6B6B')
                    ax.set_title('Monthly Sales Trend', fontweight='bold')
                    ax.set_xlabel('Month')
                    ax.set_ylabel('Sales Amount (₹)')
                    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'₹{x/1000:.0f}K' if x >= 1000 else f'₹{x:.0f}'))
                    ax.grid(True, alpha=0.3)
                    plt.xticks(rotation=45)
                    plt.tight_layout()
                    st.pyplot(fig_month)
            else:
                st.warning(f"No valid data found for {date_label} analysis")
        else:
            st.warning("No date columns available for time series analysis")
        
        # Return Analysis (if Return Type is available)
        if 'Return Type' in df.columns and not df['Return Type'].isna().all():
            st.markdown("### 🔄 Return Analysis")
            
            return_data = df[df['Return Type'].notna() & (df['Return Type'] != 'NA') & (df['Return Type'] != '')].copy()
            
            if len(return_data) > 0:
                col1, col2 = st.columns(2)
                
                with col1:
                    # Return type distribution
                    return_types = return_data['Return Type'].value_counts()
                    
                    fig_return, ax = plt.subplots(figsize=(8, 6))
                    ax.pie(return_types.values, labels=return_types.index, autopct='%1.1f%%', startangle=90)
                    ax.set_title('Return Type Distribution', fontweight='bold')
                    st.pyplot(fig_return)
                
                with col2:
                    # Return metrics
                    total_orders = len(df)
                    returned_orders = len(return_data)
                    return_rate = (returned_orders / total_orders) * 100
                    
                    if 'Refund_clean' in df.columns:
                        total_refund_amount = return_data['Refund_clean'].sum()
                        avg_refund = return_data['Refund_clean'].mean()
                        
                        st.metric("Return Rate", f"{return_rate:.1f}%")
                        st.metric("Total Refund Amount", f"₹{total_refund_amount:,.2f}")
                        st.metric("Average Refund", f"₹{avg_refund:,.2f}")
                    
                    st.write("**Return Types:**")
                    for return_type, count in return_types.items():
                        percentage = (count / returned_orders) * 100
                        st.write(f"• {return_type}: {count} ({percentage:.1f}%)")
        
        # Business Insights and Recommendations
        st.markdown("### 💡 Enhanced Business Insights & Recommendations")
        
        insights_col1, insights_col2 = st.columns(2)
        
        with insights_col1:
            st.markdown("**📊 Performance Insights:**")
            st.write(f"• Overall deduction rate: **{deduction_rate:.1f}%**")
            st.write(f"• Average order value: **₹{avg_order_value:,.2f}**")
            
            if deduction_data:
                max_deduction = max(deduction_data.items(), key=lambda x: x[1])
                st.write(f"• Highest deduction: **{max_deduction[0]}** (₹{max_deduction[1]:,.2f})")
            
            # Date-based insights
            if has_order_date:
                df_orders = df[df['Order Date_clean'].notna()].copy()
                if len(df_orders) > 0:
                    df_orders['Weekday'] = df_orders['Order Date_clean'].dt.day_name()
                    best_day = df_orders.groupby('Weekday')['Sale Amount_clean'].sum().idxmax()
                    st.write(f"• Best sales day: **{best_day}**")
            
            # Fulfillment insights
            if has_order_date and has_dispatch_date:
                df_fulfillment = df[(df['Order Date_clean'].notna()) & (df['Dispatch Date_clean'].notna())].copy()
                if len(df_fulfillment) > 0:
                    df_fulfillment['Fulfillment_Days'] = (df_fulfillment['Dispatch Date_clean'] - df_fulfillment['Order Date_clean']).dt.days
                    df_fulfillment = df_fulfillment[(df_fulfillment['Fulfillment_Days'] >= 0) & (df_fulfillment['Fulfillment_Days'] <= 30)]
                    if len(df_fulfillment) > 0:
                        avg_fulfillment = df_fulfillment['Fulfillment_Days'].mean()
                        st.write(f"• Average fulfillment time: **{avg_fulfillment:.1f} days**")
        
        with insights_col2:
            st.markdown("**🎯 Actionable Recommendations:**")
            
            # Deduction rate recommendations
            if deduction_rate > 20:
                st.write("• 🔴 **Critical**: Deduction rate >20% - urgent review needed")
            elif deduction_rate > 15:
                st.write("• 🟠 **High deduction rate** - review fee structures with marketplace")
            elif deduction_rate > 10:
                st.write("• 🟡 **Moderate deductions** - monitor and optimize where possible")
            else:
                st.write("• 🟢 **Good deduction management** - maintain current practices")
            
            # Order value recommendations
            if avg_order_value < 1000:
                st.write("• 📈 **Focus on increasing AOV** through bundles/upselling")
            elif avg_order_value < 2000:
                st.write("• 💡 **Good AOV** - explore premium product promotions")
            else:
                st.write("• ✅ **Excellent AOV** - maintain premium positioning")
            
            # Fulfillment recommendations
            if has_order_date and has_dispatch_date:
                df_fulfillment = df[(df['Order Date_clean'].notna()) & (df['Dispatch Date_clean'].notna())].copy()
                if len(df_fulfillment) > 0:
                    df_fulfillment['Fulfillment_Days'] = (df_fulfillment['Dispatch Date_clean'] - df_fulfillment['Order Date_clean']).dt.days
                    df_fulfillment = df_fulfillment[(df_fulfillment['Fulfillment_Days'] >= 0) & (df_fulfillment['Fulfillment_Days'] <= 30)]
                    if len(df_fulfillment) > 0:
                        avg_fulfillment = df_fulfillment['Fulfillment_Days'].mean()
                        if avg_fulfillment > 5:
                            st.write("• ⚡ **Improve fulfillment speed** - review inventory & processes")
                        elif avg_fulfillment > 3:
                            st.write("• 🚀 **Good fulfillment** - aim for <3 days for better ratings")
                        else:
                            st.write("• 🌟 **Excellent fulfillment speed** - maintain standards")
            
            # Return rate recommendations
            if 'Return Type' in df.columns:
                return_data = df[df['Return Type'].notna() & (df['Return Type'] != 'NA') & (df['Return Type'] != '')].copy()
                if len(return_data) > 0:
                    return_rate = (len(return_data) / len(df)) * 100
                    if return_rate > 10:
                        st.write("• 🔄 **High return rate** - analyze return reasons & improve quality")
                    elif return_rate > 5:
                        st.write("• 📦 **Monitor returns** - focus on accurate product descriptions")
                    else:
                        st.write("• ✅ **Low return rate** - maintain quality standards")
            
            # General recommendations
            st.write("• 📊 **Regular monitoring** of settlement vs sales gaps")
            st.write("• 📅 **Seasonal planning** based on order trends")
            st.write("• 🔍 **Weekly reviews** of key metrics for quick adjustments")
    
    else:
        st.warning("⚠️ Required columns (Sale Amount, Bank Settlement Value) not found in the data. Please check column mapping.")
        st.write("**Available columns:**")
        for col in df.columns:
            st.write(f"• {col}")
    
    # Data Quality and Completeness Report
    st.markdown("### 🔍 Data Quality & Completeness Report")
    
    quality_col1, quality_col2, quality_col3 = st.columns(3)
    
    with quality_col1:
        st.write("**📊 Data Coverage:**")
        st.write(f"• Total records: **{len(df):,}**")
        st.write(f"• Unique invoices: **{df['Invoice'].nunique() if 'Invoice' in df.columns else 'N/A'}**")
        st.write(f"• Unique orders: **{df['Order ID'].nunique() if 'Order ID' in df.columns else 'N/A'}**")
        
        # Date coverage
        date_coverage = []
        for date_col in ['Payment Date_clean', 'Order Date_clean', 'Invoice Date_clean', 'Dispatch Date_clean']:
            if date_col in df.columns:
                valid_dates = df[date_col].dropna()
                if len(valid_dates) > 0:
                    date_coverage.append(f"• {date_col.replace('_clean', '')}: **{len(valid_dates):,}** records")
        
        if date_coverage:
            st.write("**📅 Date Coverage:**")
            for coverage in date_coverage:
                st.write(coverage)
    
    with quality_col2:
        st.write("**⚠️ Missing Data:**")
        missing_data = df.isnull().sum()
        critical_missing = missing_data[missing_data > 0].sort_values(ascending=False)
        
        if len(critical_missing) > 0:
            for col, missing_count in critical_missing.head(8).items():
                percentage = (missing_count / len(df)) * 100
                if percentage > 50:
                    status = "🔴"
                elif percentage > 20:
                    status = "🟠"
                elif percentage > 5:
                    status = "🟡"
                else:
                    status = "🟢"
                st.write(f"• {status} {col}: {missing_count} ({percentage:.1f}%)")
        else:
            st.write("✅ No missing values detected!")
    
    with quality_col3:
        st.write("**📈 Data Quality Score:**")
        
        # Calculate quality score
        quality_factors = []
        
        # Check for key columns completeness
        key_columns = ['Sale Amount', 'Bank Settlement Value', 'Order ID', 'Invoice']
        key_completeness = sum(1 for col in key_columns if col in df.columns and df[col].notna().sum() > len(df) * 0.9) / len(key_columns)
        quality_factors.append(('Key Data Completeness', key_completeness * 100))
        
        # Check date data availability
        date_cols = ['Order Date_clean', 'Invoice Date_clean', 'Dispatch Date_clean', 'Payment Date_clean']
        available_dates = sum(1 for col in date_cols if col in df.columns and df[col].notna().sum() > 0)
        date_score = (available_dates / len(date_cols)) * 100
        quality_factors.append(('Date Data Availability', date_score))
        
        # Check numeric data validity (exclude date columns)
        numeric_cols = [col for col in df.columns if col.endswith('_clean') and not any(date_word in col.lower() for date_word in ['date', 'time'])]
        if numeric_cols:
            numeric_validity = sum(1 for col in numeric_cols if (df[col] >= 0).sum() > len(df) * 0.9) / len(numeric_cols) * 100
            quality_factors.append(('Numeric Data Validity', numeric_validity))
        
        # Overall quality score
        overall_score = sum(score for _, score in quality_factors) / len(quality_factors)
        
        if overall_score >= 90:
            score_status = "🟢 Excellent"
        elif overall_score >= 75:
            score_status = "🟡 Good"
        elif overall_score >= 60:
            score_status = "🟠 Fair"
        else:
            score_status = "🔴 Poor"
        
        st.metric("Overall Quality Score", f"{overall_score:.1f}%", score_status)
        
        st.write("**Quality Breakdown:**")
        for factor, score in quality_factors:
            st.write(f"• {factor}: {score:.1f}%")
    
    # Date Range Information
    st.markdown("### 📅 Data Period Analysis")
    
    period_col1, period_col2 = st.columns(2)
    
    with period_col1:
        st.write("**📊 Data Period Coverage:**")
        
        for date_col in ['Order Date_clean', 'Payment Date_clean', 'Invoice Date_clean', 'Dispatch Date_clean']:
            if date_col in df.columns:
                valid_dates = df[date_col].dropna()
                if len(valid_dates) > 0:
                    min_date = valid_dates.min()
                    max_date = valid_dates.max()
                    days_span = (max_date - min_date).days
                    
                    col_name = date_col.replace('_clean', '')
                    st.write(f"**{col_name}:**")
                    st.write(f"• From: {min_date.strftime('%d-%m-%Y')}")
                    st.write(f"• To: {max_date.strftime('%d-%m-%Y')}")
                    st.write(f"• Span: {days_span} days")
                    st.write("")
    
    with period_col2:
        st.write("**💡 Period Insights:**")
        
        # Most active period
        if has_order_date:
            df_orders = df[df['Order Date_clean'].notna()].copy()
            if len(df_orders) > 0:
                monthly_orders = df_orders.groupby(df_orders['Order Date_clean'].dt.to_period('M'))['Sale Amount_clean'].sum()
                best_month = monthly_orders.idxmax()
                best_month_sales = monthly_orders.max()
                st.write(f"• **Best Month**: {best_month} (₹{best_month_sales:,.0f})")
        
        # Recent performance
        if has_order_date:
            df_orders = df[df['Order Date_clean'].notna()].copy()
            if len(df_orders) > 0:
                latest_date = df_orders['Order Date_clean'].max()
                last_30_days = df_orders[df_orders['Order Date_clean'] >= (latest_date - timedelta(days=30))]
                if len(last_30_days) > 0:
                    recent_sales = last_30_days['Sale Amount_clean'].sum()
                    recent_orders = len(last_30_days)
                    st.write(f"• **Last 30 Days**: {recent_orders} orders, ₹{recent_sales:,.0f}")
        
        # Growth trend (if sufficient data)
        if has_order_date:
            df_orders = df[df['Order Date_clean'].notna()].copy()
            if len(df_orders) > 60:  # Need at least 2 months of data
                df_orders['Month'] = df_orders['Order Date_clean'].dt.to_period('M')
                monthly_sales = df_orders.groupby('Month')['Sale Amount_clean'].sum()
                if len(monthly_sales) >= 2:
                    recent_avg = monthly_sales.tail(2).mean()
                    earlier_avg = monthly_sales.head(max(2, len(monthly_sales)//2)).mean()
                    growth = ((recent_avg - earlier_avg) / earlier_avg) * 100
                    trend_arrow = "📈" if growth > 0 else "📉" if growth < -5 else "➡️"
                    st.write(f"• **Trend**: {trend_arrow} {growth:+.1f}% vs earlier period")
    
    # Export Enhanced Data Option
    st.markdown("### 📤 Export Enhanced Analysis Data")
    
    export_col1, export_col2 = st.columns(2)
    
    with export_col1:
        if st.button("📊 Generate Analysis Summary Report"):
            # Create summary report
            summary_data = {
                'Metric': [
                    'Total Sales Amount',
                    'Total Settlement Amount', 
                    'Total Deductions',
                    'Deduction Rate (%)',
                    'Average Order Value',
                    'Total Orders',
                    'Data Quality Score (%)'
                ],
                'Value': [
                    f"₹{total_sales:,.2f}",
                    f"₹{total_settlement:,.2f}",
                    f"₹{total_deductions:,.2f}",
                    f"{deduction_rate:.1f}%",
                    f"₹{avg_order_value:,.2f}",
                    f"{len(df):,}",
                    f"{overall_score:.1f}%"
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                
                # Add detailed data if user wants
                df_export = df.copy()
                # Remove helper columns
                cols_to_remove = [col for col in df_export.columns if col.endswith('_clean')]
                df_export = df_export.drop(columns=cols_to_remove)
                df_export.to_excel(writer, sheet_name="Detailed_Data", index=False)
            
            st.download_button(
                "⬇️ Download Analysis Report",
                data=output.getvalue(),
                file_name=f"LAXMIPATI_Analysis_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    with export_col2:
        st.write("**📋 Report Contents:**")
        st.write("• Executive summary with key metrics")
        st.write("• Data quality assessment")
        st.write("• Complete processed dataset")
        st.write("• Ready for stakeholder review")
        st.write("• Timestamped for record keeping")