@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo   CAP NHAT DU LIEU KPI LEN STREAMLIT CLOUD
echo ============================================
echo.

if not exist ".git" (
    echo [LOI] Khong tim thay thu muc .git tai day.
    echo Hay dat file Update_KPI.bat nay vao dung thu muc chua
    echo kpi_streamlit_app.py va DATA_KPI.xlsb
    echo Vi du: C:\DUNGPV\DVKH_KPI\giphub
    echo.
    pause
    exit /b 1
)

echo Dang dong bo voi GitHub truoc khi cap nhat...
git pull origin main --no-edit

if not %errorlevel%==0 (
    echo.
    echo [LOI] Buoc dong bo du lieu voi GitHub that bai.
    echo Co the do xung dot du lieu hoac mat ket noi mang.
    echo Xem thong bao loi o tren de biet chi tiet.
    echo Neu la xung dot file, can nho nguoi biet Git ho tro xu ly.
    echo.
    pause
    exit /b 1
)

echo.
echo Dang kiem tra thay doi...
git add .

git diff --cached --quiet
if %errorlevel%==0 (
    echo.
    echo Khong co gi thay doi so voi lan truoc.
    echo Khong can lam gi them.
    echo.
    pause
    exit /b 0
)

echo Phat hien co thay doi. Dang tao commit...
git commit -m "Cap nhat du lieu KPI - %date% %time%"

if not %errorlevel%==0 (
    echo.
    echo [LOI] Buoc tao commit that bai. Xem thong bao loi o tren.
    pause
    exit /b 1
)

echo.
echo Dang day len GitHub...
git push

if not %errorlevel%==0 (
    echo.
    echo [LOI] Buoc day du lieu len GitHub that bai.
    echo Kiem tra ket noi mang hoac dang nhap GitHub.
    echo Co the thu chay lai file nay lan nua.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   THANH CONG!
echo   Du lieu da duoc day len GitHub.
echo   Streamlit Cloud se tu dong cap nhat trong 1-2 phut.
echo   Neu sau vai phut van chua thay doi, vao share.streamlit.io,
echo   mo app, bam nut Manage app, chon Reboot app.
echo ============================================
echo.
pause