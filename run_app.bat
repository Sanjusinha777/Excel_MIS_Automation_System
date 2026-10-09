@echo off
echo ========================================================
echo   Launching Excel MIS Report Automation System
echo ========================================================
echo.
cd /d "%~dp0"
echo Starting Streamlit Web App on port 8585...
streamlit run app.py --server.port 8585
pause
