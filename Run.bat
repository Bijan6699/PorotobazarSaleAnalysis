@echo off
title KianDashboard
echo Starting KianDashboard... (close this window to stop the app)
py -m streamlit run app.py
if %errorlevel% neq 0 (
    echo.
    echo An error occurred while starting the app.
    echo If libraries are missing, run Setup.bat first.
    pause
)
