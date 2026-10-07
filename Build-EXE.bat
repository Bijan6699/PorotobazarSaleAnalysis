@echo off
title KianDashboard - Build Standalone EXE
echo ==========================================
echo  Building standalone EXE with PyInstaller
echo ==========================================
echo.
py -m PyInstaller --noconfirm --onedir --name KianDashboard run_helper.py --collect-all streamlit --collect-all plotly
if %errorlevel% neq 0 (
    echo Build failed. Run Setup.bat first to install PyInstaller.
    pause
    exit /b 1
)
copy /y app.py dist\KianDashboard\app.py >nul
echo.
echo ==========================================
echo  Build finished: dist\KianDashboard\KianDashboard.exe
echo  Note: the EXE works only on the same Windows
echo  type (64-bit) it was built on.
echo ==========================================
pause
