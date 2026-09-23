@echo off
title Fathima Traders - Payroll & Settlement Dashboard
echo ========================================================
echo   Fathima Traders - Payroll & Settlement Dashboard
echo ========================================================
echo.
echo Starting application...
echo.

cd /d "%~dp0"
if exist ".venv\Scripts\streamlit.exe" (
    ".venv\Scripts\streamlit.exe" run app.py
) else (
    streamlit run app.py
)

pause
