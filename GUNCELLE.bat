@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo === 1/3 Dashboard uretiliyor... ===
python build_dashboard.py servis_rapor.xlsx index.html
if errorlevel 1 ( echo HATA: servis_rapor.xlsx bu klasorde mi? & pause & exit /b 1 )
echo.
echo === 2/3 Paketleniyor... ===
git add index.html
git commit -m "dashboard guncelleme %date% %time%"
echo.
echo === 3/3 GitHub'a gonderiliyor... ===
git push
echo.
echo ============================================================
echo  TAMAM! Link 1-2 dakika icinde guncellenecek.
echo  https://inciroglu.github.io/servis-dashboard/
echo ============================================================
echo.
pause
