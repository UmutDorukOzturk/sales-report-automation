@echo off
REM Double-click: builds the report from every CSV / Excel file in the "exports" folder next to this file.
cd /d "%~dp0"
if not exist exports mkdir exports
python build_report.py exports
echo.
echo The report is in the "output" folder.
pause
