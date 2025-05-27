@echo off
echo PDF Screenshot Generator
echo ========================
echo.
echo This script will generate screenshots from all PDF files in the invoices folder.
echo.
pause
echo.
echo Running PDF screenshot generator...
python scr/pdf_screenshot_generator.py
echo.
echo Press any key to exit...
pause > nul 