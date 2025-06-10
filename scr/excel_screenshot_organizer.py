#!/usr/bin/env python3
"""
Excel Screenshot Organizer - One Image Per Sheet

This script creates an Excel file with one large screenshot per sheet,
sorted from oldest to newest based on the filename numbering.
Each image is sized large enough to see details clearly.
"""

import os
import re
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image

def setup_directories():
    """Setup and validate directories."""
    base_dir = Path(__file__).parent.parent
    screenshots_dir = base_dir / "output" / "screenshot_zoomed"
    output_dir = base_dir / "output"
    
    if not screenshots_dir.exists():
        print(f"Error: Screenshots directory not found at {screenshots_dir}")
        return None, None
    
    return screenshots_dir, output_dir

def extract_number_from_filename(filename):
    """Extract the number from filename for sorting."""
    # Remove extension
    name = filename.stem
    
    # Handle the base file without number (treat as 0)
    if name == "American Express - Account Activity":
        return 0
    
    # Extract number from filename like "American Express - Account Activity-1"
    match = re.search(r'-(\d+(?:\.\d+)?)$', name)
    if match:
        return float(match.group(1))
    
    return 0

def get_sorted_screenshots(screenshots_dir):
    """Get all screenshot files sorted from oldest to newest."""
    png_files = list(screenshots_dir.glob("*.png"))
    
    if not png_files:
        print(f"No PNG files found in {screenshots_dir}")
        return []
    
    # Sort by extracted number (oldest to newest)
    sorted_files = sorted(png_files, key=extract_number_from_filename)
    
    print(f"Found {len(sorted_files)} screenshot files")
    return sorted_files

def create_excel_with_large_screenshots(screenshots, output_path):
    """Create Excel file with one large screenshot per sheet."""
    try:
        # Create workbook
        wb = Workbook()
        
        # Remove the default sheet - we'll create our own
        wb.remove(wb.active)
        
        # Process each screenshot
        for i, screenshot_path in enumerate(screenshots):
            print(f"Processing sheet {i+1}/{len(screenshots)}: {screenshot_path.name}")
            
            # Create new worksheet for this screenshot
            sheet_name = f"Invoice_{i+1}"
            ws = wb.create_sheet(title=sheet_name)
            
            # Add title with filename
            ws['A1'] = screenshot_path.stem
            ws['A1'].font = ws['A1'].font.copy(bold=True)
            ws.row_dimensions[1].height = 25
            
            # Set column widths for large display
            ws.column_dimensions['A'].width = 120
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 20
            
            try:
                # Load and resize image for optimal display
                with Image.open(screenshot_path) as img:
                    img_width, img_height = img.size
                    
                    # Calculate size for good visibility (aim for ~1000px wide max)
                    max_display_width = 1000
                    if img_width > max_display_width:
                        scale_factor = max_display_width / img_width
                        display_width = max_display_width
                        display_height = int(img_height * scale_factor)
                    else:
                        display_width = img_width
                        display_height = img_height
                    
                    print(f"  Image size: {img_width}x{img_height} -> Display: {display_width}x{display_height}")
                    
                    # Create Excel image object
                    excel_img = ExcelImage(screenshot_path)
                    
                    # Set image size for good visibility
                    excel_img.width = display_width
                    excel_img.height = display_height
                    
                    # Position image starting from row 3 (leave space for title)
                    excel_img.anchor = 'A3'
                    ws.add_image(excel_img)
                    
                    # Set row height to accommodate the image
                    row_height_points = display_height * 0.75  # Convert pixels to points
                    for row in range(3, 3 + int(display_height / 20)):  # Approximate rows needed
                        ws.row_dimensions[row].height = min(400, max(15, row_height_points / 10))
                
            except Exception as img_error:
                print(f"  Error processing image {screenshot_path.name}: {str(img_error)}")
                # Add error message to sheet
                ws['A3'] = f"Error loading image: {str(img_error)}"
                continue
        
        # Save the workbook
        wb.save(output_path)
        print(f"\n✓ Excel file created successfully: {output_path}")
        return True
        
    except Exception as e:
        print(f"\n✗ Error creating Excel file: {str(e)}")
        return False

def main():
    """Main function to create Excel file with large screenshots."""
    print("Excel Screenshot Organizer - One Image Per Sheet")
    print("=" * 60)
    
    # Setup directories
    screenshots_dir, output_dir = setup_directories()
    if not screenshots_dir or not output_dir:
        sys.exit(1)
    
    # Get sorted screenshots
    screenshots = get_sorted_screenshots(screenshots_dir)
    if not screenshots:
        sys.exit(1)
    
    # Create output Excel file path
    excel_path = output_dir / "invoice_screenshots_organized.xlsx"
    
    # Create Excel file with large screenshots
    print(f"\nCreating Excel file with {len(screenshots)} screenshots...")
    print("Format: One large screenshot per sheet for maximum readability")
    
    if create_excel_with_large_screenshots(screenshots, excel_path):
        print(f"\n" + "=" * 60)
        print("✅ Excel file creation complete!")
        print(f"📁 File saved: {excel_path}")
        print(f"📊 Total screenshots: {len(screenshots)}")
        print(f"📋 Layout: One large image per sheet ({len(screenshots)} sheets)")
        print("✨ Enhanced readability: Large images for easy detail viewing")
        print("\nYou can now open the Excel file to view your organized screenshots!")
        print("Use the sheet tabs at the bottom to navigate between invoices.")
    else:
        print("\n❌ Failed to create Excel file")
        sys.exit(1)

if __name__ == "__main__":
    main() 