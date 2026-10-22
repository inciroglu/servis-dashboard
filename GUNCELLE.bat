@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo === 1/3 Panel uretiliyor (servis_rapor.xlsx + ek_satis.xlsx) ===
if not exist panel_sablon.html ( echo HATA: panel_sablon.html bu klasorde yok. & pause & exit /b 1 )
if not exist ek_veri.py ( echo HATA: ek_veri.py bu klasorde yok. & pause & exit /b 1 )
if not exist ek_satis.xlsx echo UYARI: ek_satis.xlsx yok - ek satis verisi onceki haliyle kalacak.
if not exist garanti_arac_girisi.xlsx echo UYARI: garanti_arac_girisi.xlsx yok - garanti arac girisi onceki haliyle kalacak.
if exist panel_ayar.json findstr /c:"script.google.com" panel_ayar.json >nul && (python -c "import cryptography" 2>nul || (echo Sifreleme kutuphanesi kuruluyor - tek seferlik... & python -m pip install --quiet cryptography))
python build_dashboard.py servis_rapor.xlsx index.html ek_satis.xlsx
if errorlevel 1 ( echo HATA: servis_rapor.xlsx bu klasorde mi, Excel dosyalari kapali mi? & pause & exit /b 1 )
echo.
echo === 2/3 Paketleniyor (Excel dosyalari GitHub'a GONDERILMEZ) ===
git add index.html panel_sablon.html build_dashboard.py ek_veri.py sifre_degistir.py robots.txt GUNCELLE.bat .gitignore manifest.webmanifest uygulama sw.js .github/workflows/pages.yml
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
