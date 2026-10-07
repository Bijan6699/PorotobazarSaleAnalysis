@echo off
title KianDashboard - One Time Setup
chcp 65001 >nul
echo ==========================================
echo  KianDashboard - One Time Setup
echo ==========================================
echo.

echo Checking for Python...
py --version >nul 2>&1
if errorlevel 1 (
    echo Python was not found. Downloading Python 3.13.1 installer...
    curl -L -o python-3.13.1-amd64.exe https://www.python.org/ftp/python/3.13.1/python-3.13.1-amd64.exe
    if not exist python-3.13.1-amd64.exe (
        echo Download failed. Please check your internet connection.
        pause
        exit /b 1
    )
    echo Installing Python silently, please wait...
    python-3.13.1-amd64.exe /quiet InstallAllUsers=0 PrependPath=1
    echo Re-checking Python...
    py --version >nul 2>&1
    if errorlevel 1 (
        echo Python still not found. Please restart your PC and run Setup.bat again.
        pause
        exit /b 1
    )
)
py --version
echo.

echo Upgrading pip (optional)...
py -m pip install --upgrade pip

echo Installing required libraries: streamlit plotly pandas openpyxl pyinstaller
echo Attempt 1 of 4: standard PyPI connection...
py -m pip install streamlit plotly pandas openpyxl pyinstaller
if not errorlevel 1 goto install_success

echo.
echo Standard installation failed. Retrying with a longer timeout and trusted hosts...
py -m pip install --default-timeout=100 --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org streamlit plotly pandas openpyxl pyinstaller
if not errorlevel 1 goto install_success

echo.
echo PyPI retry failed. Trying the Aliyun PyPI mirror...
py -m pip install --default-timeout=100 -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com streamlit plotly pandas openpyxl pyinstaller
if not errorlevel 1 goto install_success

echo.
echo Aliyun failed. Trying the Tsinghua PyPI mirror...
py -m pip install --default-timeout=100 -i https://pypi.tuna.tsinghua.edu.cn/simple/ --trusted-host pypi.tuna.tsinghua.edu.cn streamlit plotly pandas openpyxl pyinstaller
if not errorlevel 1 goto install_success

echo.
echo ==========================================
echo نصب کتابخانه‌ها ناموفق بود.
echo اتصال اینترنت را بررسی کنید و در صورت نیاز فیلترشکن را روشن کنید.
echo سپس Setup.bat را دوباره اجرا کنید.
echo اگر مشکل ادامه داشت، نصب پایتون و دسترسی به مخازن pip را بررسی کنید.
echo ==========================================
pause
exit /b 1

:install_success
echo.
echo ==========================================
echo  Setup finished successfully!
echo  You can now run Run.bat to start the app.
echo ==========================================
pause
