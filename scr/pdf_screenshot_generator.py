#!/usr/bin/env python3
"""
PDF Screenshot Generator

This script reads PDF files from the invoices folder, converts each PDF to an image,
and saves the screenshots in the output/screenshot directory.
"""

import os
import sys
from pathlib import Path
from pdf2image import convert_from_path
from PIL import Image

def setup_directories():
    """Ensure required directories exist."""
    base_dir = Path(__file__).parent.parent
    invoices_dir = base_dir / "invoices"
    output_dir = base_dir / "output" / "screenshot"
    
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

def convert_pdf_to_screenshot(pdf_path, output_dir):
    """Convert a single PDF to a screenshot image."""
    try:
        print(f"Processing: {pdf_path.name}")
        
        # Convert PDF to images (first page only for screenshot)
        images = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=200)
        
        if not images:
            print(f"Warning: No images generated from {pdf_path.name}")
            return False
        
        # Get the first (and only) page
        image = images[0]
        
        # Create output filename (replace .pdf with .png)
        output_filename = pdf_path.stem + ".png"
        output_path = output_dir / output_filename
        
        # Save the image
        image.save(output_path, "PNG", quality=95, optimize=True)
        print(f"✓ Screenshot saved: {output_filename}")
        
        return True
        
    except Exception as e:
        print(f"✗ Error processing {pdf_path.name}: {str(e)}")
        return False

def main():
    """Main function to process all PDF files."""
    print("PDF Screenshot Generator")
    print("=" * 50)
    
    # Setup directories
    invoices_dir, output_dir = setup_directories()
    if not invoices_dir or not output_dir:
        sys.exit(1)
    
    # Get PDF files
    pdf_files = get_pdf_files(invoices_dir)
    if not pdf_files:
        sys.exit(1)
    
    # Process each PDF file
    successful = 0
    failed = 0
    
    for pdf_file in pdf_files:
        if convert_pdf_to_screenshot(pdf_file, output_dir):
            successful += 1
        else:
            failed += 1
    
    # Summary
    print("\n" + "=" * 50)
    print(f"Processing complete!")
    print(f"✓ Successful: {successful}")
    print(f"✗ Failed: {failed}")
    print(f"Screenshots saved to: {output_dir}")

if __name__ == "__main__":
    main() 