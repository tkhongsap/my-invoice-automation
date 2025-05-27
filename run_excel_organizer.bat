@echo off
echo Excel Screenshot Organizer
echo ==========================
echo.
echo This script will create an Excel file with screenshots organized in 2 columns,
echo sorted from oldest to newest.
echo.
pause
echo.
echo Running Excel screenshot organizer...
python scr/excel_screenshot_organizer.py
echo.
echo Opening the Excel file...
start "" "output\invoice_screenshots_organized.xlsx"
echo.
echo Press any key to exit...
pause > nul 