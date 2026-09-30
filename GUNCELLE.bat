@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo === 1/3 Panel uretiliyor (panel_sablon.html + servis_rapor.xlsx) ===
if not exist panel_sablon.html ( echo HATA: panel_sablon.html bu klasorde yok. & pause & exit /b 1 )
python build_dashboard.py servis_rapor.xlsx index.html
if errorlevel 1 ( echo HATA: servis_rapor.xlsx bu klasorde mi, Excel kapali mi? & pause & exit /b 1 )
echo.
echo === 2/3 Paketleniyor (panel + sablon + scriptler) ===
git add index.html panel_sablon.html build_dashboard.py sifre_degistir.py robots.txt GUNCELLE.bat
git commit -m "panel guncelleme %date% %time%"
echo.
echo === 3/3 GitHub'a gonderiliyor... ===
git push
if errorlevel 1 (
  echo.
  echo UYARI: Gonderme tamamlanamadi. GitHub Desktop'i acip "Push origin" butonuna basin.
  echo.
  pause
  exit /b 1
)
echo.
echo ============================================================
echo  TAMAM! Link 1-2 dakika icinde guncellenecek.
echo  https://inciroglu.github.io/servis-dashboard/
echo ============================================================
echo.
pause
