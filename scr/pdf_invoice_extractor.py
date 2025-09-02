#!/usr/bin/env python3
"""
PDF Invoice Data Extractor

This script extracts transaction data (date, description, amount) from American Express
PDF statements and consolidates them into a single CSV file.
"""

import os
import csv
import re
from pathlib import Path
from datetime import datetime
import fitz  # PyMuPDF

def setup_directories():
    """Ensure required directories exist."""
    base_dir = Path(__file__).parent.parent
    invoices_dir = base_dir / "invoices"
    output_dir = base_dir / "output"
    
    if not invoices_dir.exists():
        print(f"Error: Invoices directory not found at {invoices_dir}")
        return None, None
    
    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    
    return invoices_dir, output_dir

def get_pdf_files(invoices_dir):
    """Get all PDF files from the invoices directory."""
    pdf_files = list(invoices_dir.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in {invoices_dir}")
        return []
    
    print(f"Found {len(pdf_files)} PDF files to process")
    return sorted(pdf_files)

def extract_transaction_data(pdf_path):
    """Extract transaction data from a single PDF file."""
    transactions = []
    
    try:
        # Open the PDF
        doc = fitz.open(pdf_path)
        
        for page_num, page in enumerate(doc):
            # Extract text from the page
            text = page.get_text()
            
            # Split into lines
            lines = text.split('\n')
            
            # Find transaction data
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                
                # Look for date pattern - handle both single line and split dates
                month_pattern = r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$'
                date_pattern = r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}$'
                
                date_str = None
                
                # Check for full date on single line
                if re.match(date_pattern, line):
                    date_str = line
                    i += 1
                # Check for month alone, with day on next line
                elif re.match(month_pattern, line):
                    # Check if next line is a day number
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if re.match(r'^\d{1,2}$', next_line):
                            date_str = f"{line} {next_line}"
                            i += 2
                        else:
                            i += 1
                    else:
                        i += 1
                else:
                    i += 1
                    continue
                
                if date_str:
                    # Next lines contain the description
                    description_parts = []
                    
                    # Collect description lines until we find the amount
                    while i < len(lines):
                        next_line = lines[i].strip()
                        
                        # Check if this line contains the amount (Thai Baht symbol)
                        if '฿' in next_line:
                            amount_str = next_line
                            # Extract numeric value - handle both formats with and without spaces
                            # Also handle multiline amounts where ฿ is on one line and amount on next
                            amount_match = re.search(r'฿\s*([\d,]+\.?\d*)', amount_str)
                            if not amount_match:
                                # Try without space after symbol
                                amount_match = re.search(r'฿([\d,]+\.?\d*)', amount_str)
                            
                            # If still no match, check if ฿ is alone and amount is on next line
                            if not amount_match and amount_str.strip() == '฿':
                                if i + 1 < len(lines):
                                    next_amount_line = lines[i + 1].strip()
                                    amount_match = re.match(r'^([\d,]+\.?\d*)$', next_amount_line)
                                    if amount_match:
                                        i += 1  # Skip the amount line
                            
                            if amount_match:
                                amount = amount_match.group(1).replace(',', '')
                                
                                # Build description
                                description = ' '.join(description_parts)
                                
                                # Clean up description - remove extra details
                                # Keep main merchant name
                                if description_parts:
                                    # Usually the first line is the merchant name
                                    main_description = description_parts[0]
                                    # Add location info if available
                                    for part in description_parts[1:]:
                                        if part and not part.isdigit() and len(part) > 2:
                                            if not any(x in part.lower() for x in ['will appear', 'statement', 'foreign', 'card', 'account', '00001']):
                                                main_description += f" {part}"
                                                break
                                    description = main_description
                                
                                # Add to transactions
                                if date_str and description and amount:
                                    transactions.append({
                                        'date': date_str,
                                        'description': description,
                                        'amount_thb': float(amount),
                                        'source_file': pdf_path.name
                                    })
                            break
                        elif next_line and not next_line.startswith('Will appear') and not next_line.startswith('FOREIGN'):
                            # Add to description if it's not a metadata line
                            if next_line not in ['CARD', 'ACCOUNT_ENDING', 'CARD_MEMBER'] and not next_line.startswith('TOTRAKOOL'):
                                description_parts.append(next_line)
                        
                        i += 1
        
        doc.close()
        
    except Exception as e:
        print(f"Error processing {pdf_path.name}: {str(e)}")
    
    return transactions

def parse_date_with_year(date_str, filename):
    """Parse date string and infer year from filename or use current year."""
    # Try to extract year from statement date in filename if available
    current_year = datetime.now().year
    
    # For now, use 2025 as the year since these are June 2025 statements
    # You can enhance this to extract from the PDF content if needed
    year = 2025
    
    # Parse the month and day
    try:
        # Add year to the date string
        full_date_str = f"{date_str} {year}"
        parsed_date = datetime.strptime(full_date_str, "%b %d %Y")
        return parsed_date.strftime("%Y-%m-%d")
    except:
        return date_str  # Return original if parsing fails

def save_to_csv(all_transactions, output_path):
    """Save all transactions to a CSV file."""
    if not all_transactions:
        print("No transactions to save")
        return
    
    # Sort transactions by date
    all_transactions.sort(key=lambda x: x['date'])
    
    # Write to CSV
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Date', 'Description', 'Amount (THB)', 'Source File']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for transaction in all_transactions:
            writer.writerow({
                'Date': transaction['date'],
                'Description': transaction['description'],
                'Amount (THB)': transaction['amount_thb'],
                'Source File': transaction['source_file']
            })
    
    print(f"Saved {len(all_transactions)} transactions to {output_path}")

def main():
    """Main function to process all PDF files."""
    print("PDF Invoice Data Extractor")
    print("=" * 60)
    
    # Setup directories
    invoices_dir, output_dir = setup_directories()
    if not invoices_dir or not output_dir:
        return
    
    # Get PDF files
    pdf_files = get_pdf_files(invoices_dir)
    if not pdf_files:
        return
    
    # Process each PDF file
    all_transactions = []
    successful = 0
    failed = 0
    
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"[{i}/{len(pdf_files)}] Processing {pdf_file.name}...")
        transactions = extract_transaction_data(pdf_file)
        
        if transactions:
            all_transactions.extend(transactions)
            successful += 1
            print(f"  Found {len(transactions)} transaction(s)")
        else:
            failed += 1
            print(f"  No transactions found")
    
    # Parse dates with year
    for transaction in all_transactions:
        transaction['date'] = parse_date_with_year(transaction['date'], transaction['source_file'])
    
    # Save to CSV
    output_path = output_dir / "consolidated_invoices.csv"
    save_to_csv(all_transactions, output_path)
    
    # Summary
    print("\n" + "=" * 60)
    print(f"Processing complete!")
    print(f"Successfully processed: {successful} files")
    print(f"Failed/No data: {failed} files")
    print(f"Total transactions extracted: {len(all_transactions)}")
    print(f"Output saved to: {output_path}")
    print("=" * 60)

if __name__ == "__main__":
    main()