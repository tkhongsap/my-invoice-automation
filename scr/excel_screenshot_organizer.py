#!/usr/bin/env python3
"""
Excel Screenshot Organizer

This script creates an Excel file with screenshots organized in 2 columns,
sorted from oldest to newest based on the filename numbering.
"""

import os
import re
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from PIL import Image
import numpy as np

def setup_directories():
    """Setup and validate directories."""
    base_dir = Path(__file__).parent.parent
    screenshots_dir = base_dir / "output" / "screenshot"
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

def crop_white_borders(img, border_color=255, tolerance=10):
    """Crop white borders from an image."""
    # Convert to numpy array
    img_array = np.array(img)
    
    # If image has alpha channel, use RGB only
    if len(img_array.shape) == 3 and img_array.shape[2] == 4:
        img_array = img_array[:, :, :3]
    
    # Create a mask for non-white pixels
    if len(img_array.shape) == 3:
        # For color images
        mask = np.any(img_array < (border_color - tolerance), axis=2)
    else:
        # For grayscale images
        mask = img_array < (border_color - tolerance)
    
    # Find the bounding box of non-white pixels
    rows = np.any(mask, axis=1)
    cols = np.any(mask, axis=0)
    
    if not np.any(rows) or not np.any(cols):
        # If the entire image is white, return original
        return img
    
    # Get the bounding box
    top, bottom = np.where(rows)[0][[0, -1]]
    left, right = np.where(cols)[0][[0, -1]]
    
    # Add a small margin (10 pixels) to avoid cutting too close
    margin = 10
    top = max(0, top - margin)
    bottom = min(img.height - 1, bottom + margin)
    left = max(0, left - margin)
    right = min(img.width - 1, right + margin)
    
    # Crop the image
    return img.crop((left, top, right + 1, bottom + 1))

def process_image_for_excel(image_path, max_width=400, max_height=500):
    """Crop white borders and resize image to fit nicely in Excel cell while maintaining aspect ratio."""
    try:
        with Image.open(image_path) as img:
            # First, crop white borders to show content more clearly
            print(f"  Cropping white borders from {image_path.name}...")
            cropped_img = crop_white_borders(img)
            
            # Calculate new size maintaining aspect ratio
            img_width, img_height = cropped_img.size
            
            # Calculate scaling factor
            width_ratio = max_width / img_width
            height_ratio = max_height / img_height
            scale_factor = min(width_ratio, height_ratio)
            
            new_width = int(img_width * scale_factor)
            new_height = int(img_height * scale_factor)
            
            # Resize image if needed
            if scale_factor < 1.0:
                resized_img = cropped_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            else:
                resized_img = cropped_img
                new_width, new_height = img_width, img_height
            
            # Create safe filename for temp file
            safe_name = image_path.name.replace(" ", "_").replace("-", "_")
            temp_path = image_path.parent / f"temp_processed_{safe_name}"
            resized_img.save(temp_path, "PNG", optimize=True)
            
            return temp_path, new_width, new_height
            
    except Exception as e:
        print(f"Error processing image {image_path}: {str(e)}")
        return image_path, 300, 400

def create_excel_with_screenshots(screenshots, output_path):
    """Create Excel file with screenshots organized in 2 columns."""
    try:
        # Create workbook and worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Invoice Screenshots"
        
        # Set column headers
        ws['A1'] = "Column 1"
        ws['B1'] = "Column 2"
        
        # Set column widths (in Excel units)
        ws.column_dimensions['A'].width = 50
        ws.column_dimensions['B'].width = 50
        
        # Track current row
        current_row = 3
        
        # Process screenshots in pairs for 2 columns
        for i in range(0, len(screenshots), 2):
            print(f"Processing row {(i//2) + 1}: {screenshots[i].name}", end="")
            
            # Add filename above image for Column A
            ws[f"A{current_row}"] = screenshots[i].stem
            ws[f"A{current_row}"].font = ws[f"A{current_row}"].font.copy(bold=True)
            
            # Set row height for image (in points)
            ws.row_dimensions[current_row + 1].height = 250
            
            # Process and crop first image (Column A)
            processed_path1, width1, height1 = process_image_for_excel(screenshots[i])
            img1 = ExcelImage(processed_path1)
            # Scale image to fit in cell
            img1.width = 300
            img1.height = 400
            
            # Position image in cell A (one row below filename)
            cell_a = f"A{current_row + 1}"
            img1.anchor = cell_a
            ws.add_image(img1)
            
            # Process second image (Column B) if it exists
            if i + 1 < len(screenshots):
                print(f" and {screenshots[i + 1].name}")
                
                # Add filename above image for Column B
                ws[f"B{current_row}"] = screenshots[i + 1].stem
                ws[f"B{current_row}"].font = ws[f"B{current_row}"].font.copy(bold=True)
                
                # Process and crop second image
                processed_path2, width2, height2 = process_image_for_excel(screenshots[i + 1])
                img2 = ExcelImage(processed_path2)
                # Scale image to fit in cell
                img2.width = 300
                img2.height = 400
                
                # Position image in cell B (one row below filename)
                cell_b = f"B{current_row + 1}"
                img2.anchor = cell_b
                ws.add_image(img2)
            else:
                print()
            
            # Move to next row (leave space for filename and image)
            current_row += 22
        
        # Save the workbook
        wb.save(output_path)
        
        # Clean up temporary processed files
        print("\nCleaning up temporary files...")
        temp_files = list(screenshots[0].parent.glob("temp_processed_*"))
        for temp_file in temp_files:
            try:
                temp_file.unlink()
            except Exception as e:
                print(f"Warning: Could not delete temp file {temp_file}: {e}")
        
        print(f"\n✓ Excel file created successfully: {output_path}")
        return True
        
    except Exception as e:
        print(f"\n✗ Error creating Excel file: {str(e)}")
        return False

def main():
    """Main function to create Excel file with organized screenshots."""
    print("Excel Screenshot Organizer")
    print("=" * 50)
    
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
    
    # Create Excel file with screenshots
    print(f"\nCreating Excel file with {len(screenshots)} screenshots...")
    print("Organizing in 2 columns, sorted from oldest to newest...")
    print("Features: Cropping white borders, filenames above images")
    
    if create_excel_with_screenshots(screenshots, excel_path):
        print(f"\n" + "=" * 50)
        print("✅ Excel file creation complete!")
        print(f"📁 File saved: {excel_path}")
        print(f"📊 Total screenshots: {len(screenshots)}")
        print(f"📋 Layout: 2 columns, {(len(screenshots) + 1) // 2} rows")
        print("✨ Improvements: Cropped borders for clearer content, filenames on top")
        print("\nYou can now open the Excel file to view your organized screenshots!")
    else:
        print("\n❌ Failed to create Excel file")
        sys.exit(1)

if __name__ == "__main__":
    main() 