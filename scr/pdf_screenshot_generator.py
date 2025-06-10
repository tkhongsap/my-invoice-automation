#!/usr/bin/env python3
"""
PDF Screenshot Generator - Full Page Version

This script reads PDF files from the invoices folder, converts each PDF to a full-page
high-quality image, and saves the screenshots in the output/screenshot_zoomed directory.

Updated to use full PDF pages without cropping for better compatibility.
"""

import os
import sys
from pathlib import Path
from pdf2image import convert_from_path
from PIL import Image

# Configuration constants
DPI = 150                          # Higher DPI for better quality full-page screenshots

def setup_directories():
    """Ensure required directories exist."""
    base_dir = Path(__file__).parent.parent
    invoices_dir = base_dir / "invoices"
    output_dir = base_dir / "output" / "screenshot_zoomed"
    
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

def convert_pdf_to_full_screenshot(pdf_path, output_dir):
    """Convert a single PDF to a full-page high-quality screenshot image."""
    try:
        # Create output filename (replace .pdf with .png)
        output_filename = pdf_path.stem + ".png"
        output_path = output_dir / output_filename
        
        # Check if file already exists (idempotency)
        if output_path.exists():
            print(f"[SKIP] {output_filename} already exists")
            return True
        
        print(f"[PROCESSING] {pdf_path.name}")
        
        # Convert PDF to images (first page only) at specified DPI
        images = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=DPI)
        
        if not images:
            print(f"[ERROR] No images generated from {pdf_path.name}")
            return False
        
        # Get the first (and only) page - use full page, no cropping
        page = images[0]
        
        # Save the full page image directly
        page.save(output_path, "PNG", quality=95, optimize=True)
        print(f"[OK] {output_filename} (Full page: {page.width}x{page.height})")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Processing {pdf_path.name}: {str(e)}")
        return False

def main():
    """Main function to process all PDF files."""
    print("PDF Screenshot Generator - Full Page Version")
    print("=" * 60)
    print(f"Configuration:")
    print(f"  DPI: {DPI}")
    print(f"  Mode: Full page screenshots (no cropping)")
    print("=" * 60)
    
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
    skipped = 0
    
    for pdf_file in pdf_files:
        result = convert_pdf_to_full_screenshot(pdf_file, output_dir)
        if result:
            if (output_dir / (pdf_file.stem + ".png")).exists():
                successful += 1
            else:
                skipped += 1
        else:
            failed += 1
    
    # Summary
    print("\n" + "=" * 60)
    print(f"Processing complete!")
    print(f"[OK] Successful: {successful}")
    print(f"[SKIP] Skipped (already exists): {skipped}")
    print(f"[ERROR] Failed: {failed}")
    print(f"Screenshots saved to: {output_dir}")
    print("=" * 60)

if __name__ == "__main__":
    main() 