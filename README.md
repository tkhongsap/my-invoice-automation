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

1. **Place PDF files** in the `invoices/` folder
2. **Run the script:**
   ```bash
   python scr/pdf_screenshot_generator.py
   ```

## Features

- Processes all PDF files in the `invoices/` folder
- Generates high-quality PNG screenshots (200 DPI)
- Captures the first page of each PDF
- Saves screenshots with corresponding filenames
- Provides progress feedback and error handling
- Creates output directory automatically if it doesn't exist

## Output

Screenshots are saved as PNG files in `output/screenshot/` with names corresponding to the original PDF files:
- `American Express - Account Activity.pdf` → `American Express - Account Activity.png`
- `American Express - Account Activity-1.pdf` → `American Express - Account Activity-1.png`

## Requirements

- Python 3.7+
- pdf2image library
- Pillow (PIL) library
- poppler-utils (system dependency) 