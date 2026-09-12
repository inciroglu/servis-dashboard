# İnciroğlu Servis Dashboard — Giriş Korumalı Canlı Panel

Müdürlere tek link: https://inciroglu.github.io/servis-dashboard/
Giriş: kullanıcı adı **inciroglu** · şifre → **AŞAĞIDAN DEĞİŞTİR** (şu an 123456)

## ⚠️ İLK İŞ: ŞİFREYİ DEĞİŞTİR
1. `build_dashboard.py`'yi Not Defteri ile aç.
2. En üstteki (28-29. satır) şu iki satırı bul:
       GIRIS_KULLANICI = "inciroglu"
       GIRIS_SIFRE = "123456"
3. Şifreyi güçlü bir şeyle değiştir (örn. büyük/küçük harf + rakam,
   en az 10 karakter). Kullanıcı adını da istersen değiştir.
4. Kaydet. (Şifre HTML'e düz yazılmaz, SHA-256 hash olarak gömülür.)
5. GUNCELLE.bat'a çift tıkla → yeni şifre yayınlanır.

## GÜVENLİK NOTU (public repo)
Repo public. Şifre koruması "caydırıcı" seviyededir — sıradan kişiyi durdurur,
ama teknik biri index.html'i indirip veriyi görebilir. Bu riski kabul ettin.
Ek önlemler bu pakette hazır:
- noindex: Google vb. arama motorları sayfayı listelemez (sadece link bilenler girer).
- robots.txt: arama motorlarına "tarama" engeli.
- Ham Excel GitHub'a GİTMEZ (.gitignore). Sadece index.html gider.
İleride gerçek güvenlik istersen Cloudflare Access (ücretsiz) eklenebilir.

## BİR KEZLİK KURULUM (inciroglu hesabında)
1. github.com (inciroglu ile giriş) → New repository → ad: servis-dashboard,
   Public, README/gitignore/license kapalı → Create.
2. GitHub Desktop'ı inciroglu hesabına bağla → repoyu Clone et.
3. Bu paketteki dosyaları repo klasörüne kopyala:
   build_dashboard.py, index.html, GUNCELLE.bat, .gitignore, robots.txt,
   .github/ klasörü — ve güncel Excel'ini `servis_rapor.xlsx` adıyla.
4. (Yukarıdaki gibi önce ŞİFREYİ DEĞİŞTİR, sonra GUNCELLE.bat'a bas ya da
   GitHub Desktop'tan Commit + Push yap.)
5. github.com → repo → Settings → Pages → Source: **GitHub Actions**.
6. 1-2 dk sonra link aktif. Müdürlere link + kullanıcı adı/şifre ilet.

## HAFTALIK GÜNCELLEME (tek tık)
1. Excel'i yenile + `servis_rapor.xlsx` olarak repo klasörüne kaydet.
2. GUNCELLE.bat'a çift tıkla. Dashboard üretir + GitHub'a gönderir.
3. Müdürler linke girince güncel veriyi görür.

## NOTLAR
- Giriş yapan, sekme açık kaldıkça tekrar şifre sormaz. Sağ üstte "Çıkış" var.
- Python + Git (GitHub Desktop ile gelir) kurulu olmalı.
- Dashboard: marka/ay/YTD seçici, araç başı (Mekanik), günlük tempo, iş günü
  (Cmt dahil, resmi tatil hariç), güncel ay otomatik tespiti.
