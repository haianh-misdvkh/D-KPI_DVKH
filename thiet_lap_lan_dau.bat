@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo   THIET LAP LAN DAU - PUSH CODE LEN GITHUB
echo ============================================
echo.

echo Dang cau hinh danh tinh Git...
git config --global user.email haianhpth53@gmail.com
git config --global user.name haianh-misdvkh

echo.
echo Dang them file va tao commit...
git add .
git commit -m "Dashboard KPI ban dau"

echo.
echo Dang day len GitHub...
git push -u origin main

if not %errorlevel%==0 (
    echo.
    echo [LOI] Push that bai. Xem thong bao loi o tren.
    echo Chup lai man hinh nay gui de kiem tra.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   THANH CONG! Code va du lieu da len GitHub.
echo   Tiep tuc sang buoc deploy tren share.streamlit.io
echo ============================================
echo.
pause