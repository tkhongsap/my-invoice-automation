#!/usr/bin/env python3
"""
Simple CSV to Excel Converter using pandas

This script converts the consolidated invoice CSV file to an Excel file.
"""

import pandas as pd
from pathlib import Path

def main():
    """Main function to convert CSV to Excel."""
    print("CSV to Excel Converter")
    print("=" * 60)
    
    # Setup paths
    base_dir = Path(__file__).parent.parent
    csv_path = base_dir / "output" / "consolidated_invoices.csv"
    excel_path = base_dir / "output" / "consolidated_invoices.xlsx"
    
    if not csv_path.exists():
        print(f"Error: CSV file not found at {csv_path}")
        return
    
    print(f"Reading CSV from: {csv_path}")
    
    # Read CSV file
    df = pd.read_csv(csv_path)
    
    # Convert Date column to datetime
    df['Date'] = pd.to_datetime(df['Date'])
    
    # Create vendor column by extracting from description
    df['Vendor'] = df['Description'].str.extract(r'^"?([^,\d]+)')
    df['Vendor'] = df['Vendor'].str.strip()
    
    # Create vendor summary
    vendor_summary = df.groupby('Vendor').agg({
        'Amount (THB)': ['sum', 'count', 'mean']
    }).round(2)
    
    vendor_summary.columns = ['Total Amount (THB)', 'Transaction Count', 'Average Amount (THB)']
    vendor_summary = vendor_summary.sort_values('Total Amount (THB)', ascending=False)
    vendor_summary.reset_index(inplace=True)
    
    # Create monthly summary
    df['Month'] = df['Date'].dt.to_period('M')
    monthly_summary = df.groupby('Month').agg({
        'Amount (THB)': ['sum', 'count', 'mean']
    }).round(2)
    
    monthly_summary.columns = ['Total Amount (THB)', 'Transaction Count', 'Average Amount (THB)']
    monthly_summary.reset_index(inplace=True)
    monthly_summary['Month'] = monthly_summary['Month'].astype(str)
    
    # Save to Excel with multiple sheets
    print(f"Converting to Excel: {excel_path}")
    
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        # Write transactions sheet
        df[['Date', 'Description', 'Amount (THB)', 'Source File']].to_excel(
            writer, sheet_name='Transactions', index=False
        )
        
        # Write vendor summary sheet
        vendor_summary.to_excel(
            writer, sheet_name='Vendor Summary', index=False
        )
        
        # Write monthly summary sheet
        monthly_summary.to_excel(
            writer, sheet_name='Monthly Summary', index=False
        )
        
        # Get workbook and worksheets
        workbook = writer.book
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#366092',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        money_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1
        })
        
        border_format = workbook.add_format({
            'border': 1
        })
        
        date_format = workbook.add_format({
            'num_format': 'yyyy-mm-dd',
            'border': 1
        })
        
        # Format Transactions sheet
        worksheet = writer.sheets['Transactions']
        worksheet.set_column('A:A', 12, date_format)  # Date column
        worksheet.set_column('B:B', 45, border_format)  # Description column
        worksheet.set_column('C:C', 15, money_format)  # Amount column
        worksheet.set_column('D:D', 40, border_format)  # Source File column
        
        # Apply header format
        for col_num, value in enumerate(df[['Date', 'Description', 'Amount (THB)', 'Source File']].columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Format Vendor Summary sheet
        worksheet = writer.sheets['Vendor Summary']
        worksheet.set_column('A:A', 30, border_format)  # Vendor column
        worksheet.set_column('B:B', 20, money_format)  # Total Amount column
        worksheet.set_column('C:C', 18, border_format)  # Transaction Count column
        worksheet.set_column('D:D', 20, money_format)  # Average Amount column
        
        # Apply header format
        for col_num, value in enumerate(vendor_summary.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Format Monthly Summary sheet
        worksheet = writer.sheets['Monthly Summary']
        worksheet.set_column('A:A', 15, border_format)  # Month column
        worksheet.set_column('B:B', 20, money_format)  # Total Amount column
        worksheet.set_column('C:C', 18, border_format)  # Transaction Count column
        worksheet.set_column('D:D', 20, money_format)  # Average Amount column
        
        # Apply header format
        for col_num, value in enumerate(monthly_summary.columns.values):
            worksheet.write(0, col_num, value, header_format)
    
    print("\n" + "=" * 60)
    print(f"Excel file created successfully!")
    print(f"Output saved to: {excel_path}")
    print(f"\nSummary:")
    print(f"  Total transactions: {len(df)}")
    print(f"  Total amount: ฿{df['Amount (THB)'].sum():,.2f}")
    print(f"  Date range: {df['Date'].min().strftime('%Y-%m-%d')} to {df['Date'].max().strftime('%Y-%m-%d')}")
    print(f"  Number of vendors: {df['Vendor'].nunique()}")
    print("\nSheets created:")
    print("  1. Transactions - All transaction details")
    print("  2. Vendor Summary - Totals by vendor")
    print("  3. Monthly Summary - Totals by month")
    print("=" * 60)

if __name__ == "__main__":
    main()