# Invoice PDF Screenshot Generator

This project automatically generates screenshots from PDF invoice files.

## Setup

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

   **Note for Windows users:** You may also need to install `poppler-utils`:
   - Download poppler for Windows from: https://github.com/oschwartz10612/poppler-windows/releases/
   - Extract and add the `bin` folder to your system PATH
   - Or install via conda: `conda install -c conda-forge poppler`

2. **Directory Structure:**
   ```
   my-invoices-automation/
   ├── invoices/           # Place your PDF files here
   ├── output/
   │   └── screenshot/     # Screenshots will be saved here
   └── scr/
       └── pdf_screenshot_generator.py
   ```

## Usage

### Generate Full-Page Invoice Screenshots from PDFs (Recommended)
1. **Place PDF files** in the `invoices/` folder
2. **Run the screenshot generator:**
   ```bash
   python scr/pdf_screenshot_generator.py
   ```
   Or double-click `run_screenshot_generator.bat` on Windows

   This enhanced version:
   - Renders PDFs at 150 DPI for optimal quality
   - Generates full-page screenshots (no cropping for better compatibility)
   - Creates high-resolution images (1275x1650 pixels) for clear detail viewing
   - Saves to `output/screenshot_zoomed/` folder
   - Provides idempotent operation (skips existing files)

### Organize Screenshots in Excel
1. **After generating screenshots**, create an organized Excel file:
   ```bash
   python scr/excel_screenshot_organizer.py
   ```
   Or double-click `run_excel_organizer.bat` on Windows

   The organizer creates **one large screenshot per Excel sheet** for maximum readability and easy navigation.

## Features

### PDF Screenshot Generator (Full Page Version)
- Processes all PDF files in the `invoices/` folder
- Generates high-quality PNG screenshots (150 DPI for optimal resolution)
- Creates full-page screenshots (no cropping) for better compatibility
- High-resolution output (1275x1650 pixels) for clear detail viewing
- Captures the first page of each PDF
- Saves screenshots with corresponding filenames to `output/screenshot_zoomed/`
- Provides idempotent operation (skips existing files)
- Clear progress feedback with [OK], [SKIP], and [ERROR] status indicators
- Creates output directory automatically if it doesn't exist

### Excel Screenshot Organizer (One Image Per Sheet)
- Creates an Excel file with one large screenshot per sheet (40 sheets for 40 invoices)
- Sorts screenshots from oldest to newest based on filename
- Large image display (1000px wide) for easy detail reading
- Includes filename labels above each screenshot
- Sheet tabs for easy navigation between invoices
- No image compression - preserves full detail for review

## Output

### Screenshots
Full-page screenshots are saved as PNG files in `output/screenshot_zoomed/` with names corresponding to the original PDF files:
- `American Express - Account Activity.pdf` → `American Express - Account Activity.png`
- `American Express - Account Activity-1.pdf` → `American Express - Account Activity-1.png`

These enhanced screenshots feature:
- High resolution (1275x1650 pixels) for clear detail viewing
- Full-page content (no cropping) for complete invoice visibility
- Optimized file sizes (typically 70-85KB per PNG)

### Excel File
The organized Excel file is saved as `output/invoice_screenshots_organized.xlsx` with:
- One large screenshot per sheet for maximum readability
- Screenshots sorted from oldest to newest (40 sheets total)
- Large image display (1000px wide) for easy detail reading
- Sheet tabs for quick navigation between invoices
- Filename labels above each image for easy identification

## Requirements

- Python 3.7+
- pdf2image library
- Pillow (PIL) library
- openpyxl library (for Excel file creation)
- poppler-utils (system dependency) 