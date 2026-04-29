@echo off
echo ================================================
echo  Reconciliation Dashboard - Starting...
echo ================================================
echo.
echo Opening in browser at http://localhost:8501
echo Press Ctrl+C to stop the server
echo.
cd /d "%~dp0"
echo Current folder: %CD%
echo Looking for app.py...
if not exist app.py (
    echo ERROR: app.py not found in %CD%
    echo Make sure start.bat is in the same folder as app.py
    pause
    exit
)
streamlit run app.py --server.headless false --browser.gatherUsageStats false
pause
