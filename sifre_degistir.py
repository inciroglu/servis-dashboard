"""
Dashboard giriş şifresini değiştirir.
Kullanım:  python sifre_degistir.py YeniSifre123
index.html içindeki SIFRE_HASH değerini yeni şifrenin SHA-256 özetiyle değiştirir.
"""
import hashlib, re, sys

if len(sys.argv) != 2:
    print("Kullanım: python sifre_degistir.py YeniSifre")
    sys.exit(1)

yeni = sys.argv[1]
h = hashlib.sha256(yeni.encode("utf-8")).hexdigest()
yol = "index.html"
html = open(yol, encoding="utf-8").read()
html, n = re.subn(r"const SIFRE_HASH='[0-9a-f]{64}';", f"const SIFRE_HASH='{h}';", html)
if n != 1:
    print("HATA: index.html içinde şifre alanı bulunamadı.")
    sys.exit(1)
open(yol, "w", encoding="utf-8").write(html)
print("✓ Şifre güncellendi. Şimdi GitHub Desktop'ta Commit + Push yapın.")
