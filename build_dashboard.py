#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
İNCİROĞLU OTOMOTİV — Servis Gidişat Dashboard üreteci
------------------------------------------------------
Excel'in GÖRÜNÜR sayfalarını okur, tek dosyalık kurumsal HTML dashboard üretir.
Gizli sayfalara (Fatura Detay, Fatura Listesi, Bütçe vb.) DOKUNMAZ.

Kullanım:
    python build_dashboard.py                       # varsayılan: servis_rapor.xlsx -> dashboard.html
    python build_dashboard.py "dosyam.xlsx"          # farklı girdi
    python build_dashboard.py "dosyam.xlsx" cikti.html

Veri yenileme akışı:
    1. Excel'de ham veriyi yenile ve KAYDET.
    2. Bu scripti tekrar çalıştır.
    3. dashboard.html'i tarayıcıda aç (çift tıkla) — güncel veri gelir.
"""

import sys
import os
import hashlib

# ================================================================== #
#  GİRİŞ BİLGİSİ — müdürlere vereceğiniz ortak kullanıcı adı + şifre
#  Değiştirmek için: aşağıdaki iki değeri yenileyip script'i tekrar
#  çalıştırın. Şifre HTML'e düz yazılmaz, SHA-256 hash olarak gömülür.
# ================================================================== #
GIRIS_KULLANICI = "inciroglu"
GIRIS_SIFRE = "123456"
# ================================================================== #

import re
import json
import zipfile
import datetime as _dt
import xml.etree.ElementTree as ET

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


class HizliExcel:
    """
    Sadece gereken sayfaların XML'ini okur; 400bin satırlık gizli sayfaları
    hiç açmaz. openpyxl'in tüm workbook'u parse etme maliyetinden kaçınır.
    Değer hücrelerini (formül sonuçlarının cache'lenmiş hali) okur.
    """
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self._shared = None
        self._sheet_map = {}   # sayfa adı -> worksheets/sheetN.xml
        self._state = {}       # sayfa adı -> visible/hidden
        self._parse_workbook()

    def _parse_workbook(self):
        wbxml = self.z.read("xl/workbook.xml").decode("utf-8")
        rels = self.z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        for m in re.finditer(r'<sheet\b[^>]*/>', wbxml):
            tag = m.group(0)
            name = re.search(r'name="([^"]+)"', tag)
            rid = re.search(r'r:id="(rId\d+)"', tag)
            state = re.search(r'state="([^"]+)"', tag)
            if not (name and rid):
                continue
            name = self._unescape(name.group(1))
            tgt = relmap.get(rid.group(1), "")
            tgt = tgt.lstrip("/")               # "/xl/..." -> "xl/..."
            if not tgt.startswith("xl/"):
                tgt = "xl/" + tgt
            self._sheet_map[name] = tgt
            self._state[name] = state.group(1) if state else "visible"

    @staticmethod
    def _unescape(s):
        return (s.replace("&amp;", "&").replace("&lt;", "<")
                 .replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'"))

    @property
    def sheetnames(self):
        return list(self._sheet_map.keys())

    def sheet_state(self, name):
        return self._state.get(name, "visible")

    def _shared_strings(self):
        if self._shared is None:
            self._shared = []
            try:
                data = self.z.read("xl/sharedStrings.xml")
            except KeyError:
                return self._shared
            root = ET.fromstring(data)
            for si in root.findall(f"{_NS}si"):
                self._shared.append("".join(t.text or "" for t in si.iter(f"{_NS}t")))
        return self._shared

    def read_sheet(self, name):
        """{ 'A1': deger, ... } döner (sadece dolu hücreler, değer olarak)."""
        tgt = self._sheet_map.get(name)
        if not tgt:
            return {}
        shared = self._shared_strings()
        cells = {}
        root = ET.fromstring(self.z.read(tgt))
        for c in root.iter(f"{_NS}c"):
            ref = c.get("r")
            if not ref:
                continue
            t = c.get("t")
            v_el = c.find(f"{_NS}v")
            if t == "s":
                idx = int(v_el.text) if v_el is not None and v_el.text else -1
                cells[ref] = shared[idx] if 0 <= idx < len(shared) else None
            elif t == "inlineStr":
                is_el = c.find(f"{_NS}is")
                cells[ref] = "".join(x.text or "" for x in is_el.iter(f"{_NS}t")) if is_el is not None else None
            else:
                if v_el is None or v_el.text is None:
                    cells[ref] = None
                else:
                    txt = v_el.text
                    try:
                        cells[ref] = float(txt)
                    except ValueError:
                        cells[ref] = txt
        return cells

    def close(self):
        self.z.close()

# ------------------------------------------------------------------ #
#  Sabitler — Excel yapısına göre eşleme
# ------------------------------------------------------------------ #

# Konsolide + 10 marka sayfası aynı şablonda. Her biri:
#   YTD bloğu  : satır 4..11  (Toplam + 7 kategori)
#   Aylık blok : her ay için 9 satırlık grup (başlık + Toplam + 7 kategori)
# 12 ay bloklarının BAŞLIK satırları (I sütununda ay adı) şu satırlarda:
AY_BASLIK_SATIRLARI = [13, 22, 31, 40, 49, 58, 67, 76, 85, 94, 103, 112]
AYLAR = ["OCAK", "ŞUBAT", "MART", "NİSAN", "MAYIS", "HAZİRAN",
         "TEMMUZ", "AĞUSTOS", "EYLÜL", "EKİM", "KASIM", "ARALIK"]
KATEGORILER = ["Mekanik", "Garanti", "Kaporta Boya", "Banko", "Car Care", "Aksesuar", "Lastik"]

# ------------------------------------------------------------------ #
#  RESMİ TATİLLER (2026) — iş günü hesabı için
#  İş günü kuralı: Cumartesi DAHİL, Pazar HARİÇ, aşağıdaki günler HARİÇ.
#  Gün ekleyip çıkarmak için bu listeyi düzenleyin ('YYYY-MM-DD').
# ------------------------------------------------------------------ #
RESMI_TATILLER = {
    "2026-01-01",                                            # Yılbaşı
    "2026-03-20", "2026-03-21", "2026-03-22",              # Ramazan Bayramı
    "2026-04-23",                                            # 23 Nisan
    "2026-05-01",                                            # 1 Mayıs
    "2026-05-19",                                            # 19 Mayıs
    "2026-05-27", "2026-05-28", "2026-05-29", "2026-05-30", # Kurban Bayramı
    "2026-07-15",                                            # 15 Temmuz Demokrasi
    "2026-08-30",                                            # 30 Ağustos
    "2026-10-29",                                            # 29 Ekim
}


def _ay_is_gunleri(yil, ay):
    """O ayın iş günü olan gün numaralarını döner (Cmt dahil, Paz + resmi tatil hariç)."""
    import calendar
    n = calendar.monthrange(yil, ay)[1]
    gunler = []
    for d in range(1, n + 1):
        t = _dt.date(yil, ay, d)
        if t.weekday() == 6:  # Pazar
            continue
        if t.strftime("%Y-%m-%d") in RESMI_TATILLER:
            continue
        gunler.append(d)
    return gunler


def is_gunu_hesapla(yil, ay, kesim_gun):
    """Dönem / geçen (kesim günü hariç) / kalan iş günü sayılarını döner."""
    gunler = _ay_is_gunleri(yil, ay)
    donem = len(gunler)
    gecen = sum(1 for d in gunler if d < kesim_gun)   # bugün (kesim) hariç
    return donem, gecen, donem - gecen

# Marka sayfaları (Konsolide + 10 marka). Sıra dashboard'daki seçici sırası.
MARKA_SAYFALARI = ["Servis Konsolide Rapor", "Peugeot", "Opel", "Citroen", "Fiat",
                   "Arj", "Honda", "Bmw", "Motorrad", "Tesla", "Jaecoo"]

# Fatura Detay'daki "Marka" değerini marka sayfa anahtarına çevir
FATURA_MARKA_ESLEME = {
    "peugeot": "Peugeot", "opel": "Opel", "citroen": "Citroen", "citroën": "Citroen",
    "fiat": "Fiat", "arj": "Arj", "alfa": "Arj", "jeep": "Arj",
    "honda": "Honda", "bmw": "Bmw", "bmw motorrad": "Motorrad", "motorrad": "Motorrad",
    "tesla": "Tesla", "jaecoo": "Jaecoo", "omoda": "Jaecoo",
}

# Ek satış sayılan Stok Özel Kod (G sütunu) kuralları
def _ek_satis_mi(g):
    if not g:
        return False
    g = str(g).strip().upper()
    return g.startswith("AKSAT") or g in ("MOTOR YAĞI", "MOTOR YAĞI2", "CAR CARE", "AKSAT- CAR CARE")


def ek_satis_oku(dosya):
    """
    Ham 'Fatura Detay Liste' sayfasından kişi (servis danışmanı) bazında ek satış
    özetini çıkarır. Streaming okur (dev sayfa), yalnızca ÖZET döner — ham fatura
    (müşteri, fatura no, şase) DIŞARI ÇIKMAZ.
    Dönüş: { marka_key: { danisman: { ay_no: [ciro, adet] } } }
    """
    import zipfile as _zip
    import xml.etree.ElementTree as _ET
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    SHEET = "Fatura Detay Liste"

    z = _zip.ZipFile(dosya)
    # sayfa hedefini bul
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    tgt = None
    for m in re.finditer(r'<sheet\b[^>]*/>', wbxml):
        tag = m.group(0)
        nm = re.search(r'name="([^"]+)"', tag)
        rid = re.search(r'r:id="(rId\d+)"', tag)
        if nm and nm.group(1) == SHEET and rid:
            t = relmap.get(rid.group(1), "").lstrip("/")
            tgt = t if t.startswith("xl/") else "xl/" + t
            break
    if not tgt:
        print("  ⚠ 'Fatura Detay Liste' bulunamadı — ek satış atlanıyor.")
        z.close()
        return {}

    # shared strings (streaming)
    shared = []
    for ev, el in _ET.iterparse(z.open("xl/sharedStrings.xml")):
        if el.tag == NS + "si":
            shared.append("".join(t.text or "" for t in el.iter(NS + "t")))
            el.clear()

    def val(c):
        t = c.get("t"); v = c.find(NS + "v")
        if v is None or v.text is None:
            return None
        return shared[int(v.text)] if t == "s" else v.text

    from collections import defaultdict
    data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: [0.0, 0])))
    satir = 0
    for ev, el in _ET.iterparse(z.open(tgt)):
        if el.tag == NS + "row":
            rw = int(el.get("r", "0"))
            if rw >= 4:  # başlıklar 3. satırda
                cells = {}
                for c in el.iter(NS + "c"):
                    ref = c.get("r"); mm = re.match(r"([A-Z]+)", ref)
                    if mm:
                        cells[mm.group(1)] = val(c)
                g = cells.get("G")
                if _ek_satis_mi(g):
                    dan = (cells.get("H") or "DANIŞMANSIZ").strip()
                    mk_raw = (cells.get("K") or "").strip().lower()
                    mk = FATURA_MARKA_ESLEME.get(mk_raw, cells.get("K") or "Bilinmiyor")
                    ay = cells.get("P")
                    try:
                        ay = int(float(ay))
                    except Exception:
                        ay = 0
                    try:
                        f = float(cells.get("F") or 0)
                    except Exception:
                        f = 0.0
                    try:
                        q = int(float(cells.get("I") or 0))
                    except Exception:
                        q = 0
                    data[mk][dan][ay][0] += f
                    data[mk][dan][ay][1] += q
                satir += 1
            el.clear()
    z.close()
    print(f"  ✓ Ek satış: {satir:,} fatura satırı tarandı, {sum(len(v) for v in data.values())} danışman kaydı")
    # defaultdict -> normal dict
    return {mk: {dan: {str(ay): v for ay, v in aylar.items()} for dan, aylar in dans.items()}
            for mk, dans in data.items()}


MARKA_ETIKET = {
    "Servis Konsolide Rapor": "TÜM MARKALAR (Konsolide)",
    "Peugeot": "Peugeot", "Opel": "Opel", "Citroen": "Citroën", "Fiat": "Fiat",
    "Arj": "Alfa / Jeep", "Honda": "Honda", "Bmw": "BMW", "Motorrad": "BMW Motorrad",
    "Tesla": "Tesla", "Jaecoo": "Jaecoo / Omoda",
}

# Metrik blokları -> sütun harfleri (K/L/M/N/O ... ). Her metrik: Bütçe, Fiili, %, 2025 Fiili, Δ%
# Sütunlar: adet K-O | işçilik P-T | YP U-Y | toplam Z-AD
METRIK_SUTUN = {
    "adet":     dict(butce="K", fiili="L", yuzde="M", g2025="N", delta="O"),
    "iscilik":  dict(butce="P", fiili="Q", yuzde="R", g2025="S", delta="T"),
    "yp":       dict(butce="U", fiili="V", yuzde="W", g2025="X", delta="Y"),
    "toplam":   dict(butce="Z", fiili="AA", yuzde="AB", g2025="AC", delta="AD"),
}


def _num(v):
    """Hücre değerini güvenli float'a çevir (None/boş -> 0)."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(".", "").replace(",", ".").replace("₺", "").strip())
    except Exception:
        return 0.0


def _oku_metrik_bloku(cells, satir):
    """Bir satırdaki 4 metriğin (adet/işçilik/yp/toplam) tüm alanlarını sözlük olarak döner."""
    out = {}
    for m, cols in METRIK_SUTUN.items():
        out[m] = {alan: _num(cells.get(f"{harf}{satir}")) for alan, harf in cols.items()}
    return out


def marka_sayfasini_oku(cells):
    """Bir marka sayfasının hücrelerinden YTD + 12 aylık + kategori kırılımını çıkarır."""
    veri = {"ytd": {}, "aylar": {}}

    # --- YTD bloğu (satır 4 = Toplam, 5..11 = kategoriler) ---
    veri["ytd"]["toplam"] = _oku_metrik_bloku(cells, 4)
    veri["ytd"]["kategoriler"] = {}
    for i, kat in enumerate(KATEGORILER):
        veri["ytd"]["kategoriler"][kat] = _oku_metrik_bloku(cells, 5 + i)

    # --- YTD araç başı (Excel'in hazır bloğu, Mekanik = satır 5) ---
    # İşçilik: AH=bütçe, AI=2026 fiili, AJ=2025 fiili
    # YP     : AM=bütçe, AN=2026 fiili, AO=2025 fiili
    veri["ytd"]["arac_basi"] = {
        "iscilik": {"butce": _num(cells.get("AH5")), "fiili": _num(cells.get("AI5")), "g2025": _num(cells.get("AJ5"))},
        "yp":      {"butce": _num(cells.get("AM5")), "fiili": _num(cells.get("AN5")), "g2025": _num(cells.get("AO5"))},
    }

    # --- 12 aylık bloklar ---
    for idx, bas in enumerate(AY_BASLIK_SATIRLARI):
        ay = AYLAR[idx]
        blok = {"toplam": _oku_metrik_bloku(cells, bas), "kategoriler": {}}
        for i, kat in enumerate(KATEGORILER):
            blok["kategoriler"][kat] = _oku_metrik_bloku(cells, bas + 1 + i)
        veri["aylar"][ay] = blok
    return veri


def _dt_days(yil, ay):
    import calendar as _cal
    return _cal.monthrange(yil, ay)[1]


def dashboard_ust_bilgi(cells):
    """Dashboard sayfasından dönem/iş günü bilgisi (seçili ay, kesim tarihi, iş günleri)."""
    def g(cell):
        return cells.get(cell)
    kesim_raw = g("O4")
    kesim_dt = None
    if isinstance(kesim_raw, (int, float)):
        try:
            kesim_dt = _dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(kesim_raw))
        except Exception:
            kesim_dt = None
    elif isinstance(kesim_raw, str) and kesim_raw.strip():
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                kesim_dt = _dt.datetime.strptime(kesim_raw.strip(), fmt); break
            except Exception:
                pass

    secili_ay = str(g("G4") or "").strip()

    # İş günü: kesim tarihinden yıl+ay al; yoksa seçili ay adından ay numarası bul
    if kesim_dt is not None:
        yil, ay, kgun = kesim_dt.year, kesim_dt.month, kesim_dt.day
    else:
        yil = 2026
        ay = (AYLAR.index(secili_ay.upper()) + 1) if secili_ay.upper() in AYLAR else 7
        import calendar
        kgun = calendar.monthrange(yil, ay)[1] + 1  # kesim yoksa ay sonu kabul

    donem, gecen, kalan = is_gunu_hesapla(yil, ay, kgun)
    kesim_str = kesim_dt.strftime("%d.%m.%Y") if kesim_dt else str(kesim_raw or "").strip()

    # YTD iş günü: ocak..(ay-1) tam aylar + kesim ayının geçen günü
    ytd_gecen = sum(len(_ay_is_gunleri(yil, a)) for a in range(1, ay)) + gecen
    # YTD dönem iş günü: ocak..ay tam (kesim ayı dahil tamamı)
    ytd_donem = sum(len(_ay_is_gunleri(yil, a)) for a in range(1, ay + 1))

    return {
        "secili_ay": secili_ay,
        "donem_tarihleri": str(g("M4") or "").strip(),
        "veri_kesim": kesim_str,
        "kesim_yil": yil,
        "kesim_ay": ay,
        "donem_is_gunu": donem,
        "gecen_is_gunu": gecen,
        "kalan_is_gunu": kalan,
        "ytd_donem_is_gunu": ytd_donem,
        "ytd_gecen_is_gunu": ytd_gecen,
        "ytd_kalan_is_gunu": ytd_donem - ytd_gecen,
    }


def veriyi_topla(dosya):
    print(f"→ Okunuyor: {dosya}  (yalnızca görünür sayfalar, gizli sayfalar atlanır)")
    wb = HizliExcel(dosya)
    mevcut = set(wb.sheetnames)

    ust = dashboard_ust_bilgi(wb.read_sheet("Dashboard")) if "Dashboard" in mevcut else {}
    if not ust.get("veri_kesim"):
        sft = son_fatura_tarihi(dosya)
        if sft:
            ust["veri_kesim"] = sft.strftime("%d.%m.%Y")
            ust.setdefault("kesim_yil", sft.year)
            print(f"  ✓ Kesim tarihi son faturadan alındı: {ust['veri_kesim']}")

    markalar = {}
    for sheet in MARKA_SAYFALARI:
        if sheet not in mevcut:
            print(f"  ⚠ '{sheet}' bulunamadı, atlanıyor.")
            continue
        if wb.sheet_state(sheet) != "visible":
            print(f"  ⚠ '{sheet}' gizli, atlanıyor.")
            continue
        cells = wb.read_sheet(sheet)
        markalar[sheet] = {
            "etiket": MARKA_ETIKET.get(sheet, sheet),
            "veri": marka_sayfasini_oku(cells),
        }
        print(f"  ✓ {MARKA_ETIKET.get(sheet, sheet)}")

    wb.close()

    # --- Güncel ayı Dashboard!G4 yerine VERİDEN tespit et ---
    # Konsolide (yoksa ilk marka) sayfasında adet fiili > 0 olan SON ay = güncel ay.
    ref_kaynak = markalar.get("Servis Konsolide Rapor") or (next(iter(markalar.values())) if markalar else None)
    guncel_ay = None
    if ref_kaynak:
        for ay in reversed(AYLAR):
            blk = ref_kaynak["veri"]["aylar"].get(ay)
            if blk and blk["toplam"]["adet"]["fiili"] > 0:
                guncel_ay = ay
                break
    if guncel_ay:
        ust["secili_ay"] = guncel_ay
        # İş günü / dönem bilgisini güncel aya göre yeniden hesapla
        ay_no = AYLAR.index(guncel_ay) + 1
        yil = ust.get("kesim_yil", 2026)
        # Kesim günü: kesim tarihi güncel ay ile aynı aydaysa onun günü; değilse bugüne göre
        kdt = None
        if ust.get("veri_kesim"):
            for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                try:
                    kdt = _dt.datetime.strptime(ust["veri_kesim"], fmt); break
                except Exception:
                    pass
        if kdt is not None and kdt.month == ay_no and kdt.year == yil:
            kgun = kdt.day
            kesim_str = ust["veri_kesim"]
        else:
            # kesim tarihi bu ayı göstermiyor -> bugünün tarihini kullan
            bugun = _dt.date.today()
            if bugun.month == ay_no and bugun.year == yil:
                kgun = bugun.day
            else:
                import calendar as _cal
                kgun = _cal.monthrange(yil, ay_no)[1] + 1  # ay bitmişse tüm ay geçti say
            kesim_str = f"{min(kgun, _dt_days(yil, ay_no)):02d}.{ay_no:02d}.{yil}"

        donem, gecen, kalan = is_gunu_hesapla(yil, ay_no, kgun)
        ytd_gecen = sum(len(_ay_is_gunleri(yil, a)) for a in range(1, ay_no)) + gecen
        ytd_donem = sum(len(_ay_is_gunleri(yil, a)) for a in range(1, ay_no + 1))
        ust.update({
            "kesim_ay": ay_no,
            "veri_kesim": kesim_str,
            "donem_tarihleri": f"01.{ay_no:02d}.{yil} – {_dt_days(yil, ay_no):02d}.{ay_no:02d}.{yil}",
            "donem_is_gunu": donem, "gecen_is_gunu": gecen, "kalan_is_gunu": kalan,
            "ytd_donem_is_gunu": ytd_donem, "ytd_gecen_is_gunu": ytd_gecen,
            "ytd_kalan_is_gunu": ytd_donem - ytd_gecen,
        })

        # Her ay için ayrı iş günü tablosu (ay seçicide seçilen aya göre gösterilir)
        # - Geçmiş aylar: dönem = geçen (ay bitmiş)
        # - Güncel ay: kesim gününe göre geçen/kalan
        # - Gelecek aylar: geçen 0
        ay_is_gunu = {}
        for i, ad in enumerate(AYLAR):
            no = i + 1
            toplam = len(_ay_is_gunleri(yil, no))
            if no < ay_no:
                g_, k_ = toplam, 0
            elif no == ay_no:
                g_, k_ = gecen, kalan
            else:
                g_, k_ = 0, toplam
            ay_is_gunu[ad] = {"donem": toplam, "gecen": g_, "kalan": k_}
        ust["ay_is_gunu"] = ay_is_gunu

    # Konsolide sayfasında araç başı bloğu yok — tüm markaların Mekanik YTD
    # değerlerinden hesapla (işçilik toplamı ÷ Mekanik adet toplamı vb.)
    KONSOL = "Servis Konsolide Rapor"
    if KONSOL in markalar:
        mek_isc_f = mek_isc_b = mek_isc_25 = 0.0
        mek_yp_f = mek_yp_b = mek_yp_25 = 0.0
        mek_adet_f = mek_adet_25 = mek_adet_b = 0.0
        for s, obj in markalar.items():
            if s == KONSOL:
                continue
            mek = obj["veri"]["ytd"]["kategoriler"]["Mekanik"]
            mek_isc_f += mek["iscilik"]["fiili"]; mek_isc_b += mek["iscilik"]["butce"]; mek_isc_25 += mek["iscilik"]["g2025"]
            mek_yp_f += mek["yp"]["fiili"];       mek_yp_b += mek["yp"]["butce"];       mek_yp_25 += mek["yp"]["g2025"]
            mek_adet_f += mek["adet"]["fiili"];   mek_adet_b += mek["adet"]["butce"];   mek_adet_25 += mek["adet"]["g2025"]
        def _bol(x, y): return (x / y) if y else 0.0
        markalar[KONSOL]["veri"]["ytd"]["arac_basi"] = {
            "iscilik": {"butce": _bol(mek_isc_b, mek_adet_b), "fiili": _bol(mek_isc_f, mek_adet_f), "g2025": _bol(mek_isc_25, mek_adet_25)},
            "yp":      {"butce": _bol(mek_yp_b, mek_adet_b),  "fiili": _bol(mek_yp_f, mek_adet_f),  "g2025": _bol(mek_yp_25, mek_adet_25)},
        }

    # Ek satış özelliği kaldırıldı (ham fatura okunmuyor — hızlı build)
    ek_satis = {}

    return {
        "ust": ust,
        "aylar": AYLAR,
        "kategoriler": KATEGORILER,
        "marka_sirasi": [s for s in MARKA_SAYFALARI if s in markalar],
        "markalar": markalar,
        "ek_satis": ek_satis,
        "uretim": _dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
    }


def html_uret(data):
    """Tek dosyalık, bağımsız kurumsal HTML dashboard (giriş korumalı)."""
    payload = json.dumps(data, ensure_ascii=False)
    kul_hash = hashlib.sha256(GIRIS_KULLANICI.strip().lower().encode()).hexdigest()
    sif_hash = hashlib.sha256(GIRIS_SIFRE.encode()).hexdigest()
    return (HTML_SABLON
            .replace("__DATA__", payload)
            .replace("__KUL_HASH__", kul_hash)
            .replace("__SIF_HASH__", sif_hash))


# ------------------------------------------------------------------ #
#  HTML / CSS / JS şablonu  (Chart.js CDN'den, geri kalan gömülü)
# ------------------------------------------------------------------ #
HTML_SABLON = r"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow, noarchive">
<meta name="googlebot" content="noindex, nofollow">
<meta name="referrer" content="no-referrer">
<title>İnciroğlu Otomotiv · Servis Gidişat Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{
    --bg:#f2ede3; --panel:#ffffff; --panel2:#f6f1e7; --line:#ddd3c0;
    --ink:#1a2c4e; --muted:#6b7280; --brand:#1a2c4e; --brand2:#14213d;
    --ok:#2f7d54; --warn:#b8860b; --bad:#b23b3b; --acc:#3a5a8c;
    --gold:#b8945a;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:Arial,'Segoe UI',sans-serif;background:var(--bg);color:var(--ink);
       -webkit-font-smoothing:antialiased;padding:0 0 60px}
  .wrap{max-width:1340px;margin:0 auto;padding:0 24px}

  header{background:linear-gradient(135deg,#1a2c4e,#14213d);border-bottom:3px solid var(--gold);
         padding:22px 0 18px;position:sticky;top:0;z-index:20;box-shadow:0 4px 18px rgba(26,44,78,.25)}
  .hd{display:flex;align-items:center;gap:18px;flex-wrap:wrap}
  .logo{display:flex;align-items:center;gap:14px}
  .mark{width:46px;height:46px;border-radius:10px;background:var(--gold);
        display:flex;align-items:center;justify-content:center;font-weight:800;font-size:22px;
        color:#1a2c4e;letter-spacing:-1px;box-shadow:0 4px 14px rgba(184,148,90,.4)}
  .logo h1{font-size:19px;letter-spacing:.3px;line-height:1.1;color:#fff}
  .logo p{font-size:11px;color:#c9d4e8;letter-spacing:2px;margin-top:2px}
  .meta{margin-left:auto;text-align:right;font-size:12px;color:#c9d4e8;line-height:1.7}
  .meta b{color:#fff;font-weight:700}

  .controls{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:22px 0 6px}
  .controls label{font-size:12px;color:var(--muted);margin-right:2px}
  select{background:var(--panel);color:var(--ink);border:1px solid var(--line);
         border-radius:9px;padding:10px 14px;font-size:14px;font-weight:600;cursor:pointer;min-width:210px}
  select:focus{outline:2px solid var(--brand)}
  .seg{display:inline-flex;background:var(--panel);border:1px solid var(--line);border-radius:9px;overflow:hidden}
  .seg button{background:transparent;color:var(--muted);border:0;padding:10px 16px;font-size:13px;
             font-weight:700;cursor:pointer;font-family:inherit}
  .seg button.on{background:var(--brand);color:#fff}
  .spacer{flex:1}
  .tag{font-size:12px;color:var(--ink);background:var(--panel2);border:1px solid var(--line);
       padding:8px 12px;border-radius:9px}

  .period-bar{display:flex;gap:20px;flex-wrap:wrap;align-items:center;background:var(--panel);
              border:1px solid var(--line);border-radius:12px;padding:14px 20px;margin:14px 0 4px;box-shadow:0 1px 3px rgba(26,44,78,.06)}
  .period-bar .pi{font-size:12px;color:var(--muted)}
  .period-bar .pi b{display:block;font-size:16px;color:var(--ink);margin-top:2px}
  .prog{flex:1;min-width:180px}
  .prog .bar{height:8px;background:var(--panel2);border-radius:99px;overflow:hidden;margin-top:6px}
  .prog .fill{height:100%;background:linear-gradient(90deg,var(--brand),var(--gold))}

  h2.sec{font-size:13px;letter-spacing:2px;color:var(--muted);margin:30px 0 12px;font-weight:700;
         display:flex;align-items:center;gap:12px}
  h2.sec .metric-seg{margin-left:auto}
  h2.sec .metric-seg button{padding:7px 13px;font-size:12px}

  .cards{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
  @media(max-width:1050px){.cards{grid-template-columns:repeat(2,1fr)}}
  @media(max-width:560px){.cards{grid-template-columns:1fr}}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px 20px;
        position:relative;overflow:hidden;box-shadow:0 1px 3px rgba(26,44,78,.06)}
  .card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--brand)}
  .card.alt::before{background:var(--gold)}
  .card .t{font-size:12px;color:var(--muted);letter-spacing:.5px;font-weight:700}
  .card .v{font-size:26px;font-weight:800;margin:8px 0 2px;letter-spacing:-.5px}
  .card .sub{font-size:12px;color:var(--muted);line-height:1.7}
  .pill{display:inline-block;font-size:11px;font-weight:800;padding:3px 9px;border-radius:99px;margin-top:8px}
  .pill.up{background:rgba(47,125,84,.13);color:var(--ok)}
  .pill.dn{background:rgba(178,59,59,.13);color:var(--bad)}
  .pill.nt{background:rgba(107,114,128,.13);color:var(--muted)}
  .kv{display:flex;justify-content:space-between;font-size:12px;color:var(--muted);margin-top:5px}
  .kv b{color:var(--ink);font-weight:700}

  .grid2{display:grid;grid-template-columns:1.15fr .85fr;gap:16px}
  @media(max-width:980px){.grid2{grid-template-columns:1fr}}
  .box{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px 20px;box-shadow:0 1px 3px rgba(26,44,78,.06)}
  .box h3{font-size:14px;margin-bottom:14px;font-weight:700}
  .box h3 span{color:var(--muted);font-weight:400;font-size:12px}

  table{width:100%;border-collapse:collapse;font-size:12.5px}
  th,td{padding:9px 8px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap}
  th:first-child,td:first-child{text-align:left}
  thead th{color:var(--muted);font-size:11px;letter-spacing:.4px;font-weight:700;border-bottom:2px solid var(--line)}
  tbody tr:hover{background:var(--panel2)}
  tr.tot{font-weight:800}
  tr.tot td{border-top:2px solid var(--brand);border-bottom:0}
  .mini{font-size:11px;font-weight:800;padding:2px 7px;border-radius:99px}
  .mini.up{background:rgba(47,125,84,.15);color:var(--ok)}
  .mini.dn{background:rgba(178,59,59,.15);color:var(--bad)}
  .mini.nt{color:var(--muted)}
  .barcell{position:relative}
  .barcell .b{position:absolute;left:0;top:0;bottom:0;background:rgba(58,90,140,.14);border-radius:4px}
  .barcell span{position:relative}

  canvas{max-height:300px}
  .foot{margin-top:34px;font-size:11px;color:var(--muted);text-align:center;line-height:1.8}
  .foot b{color:var(--muted)}

  /* ---- SOL MENÜLÜ LAYOUT ---- */
  .layout{display:flex;min-height:100vh}
  .sidebar{width:225px;flex-shrink:0;background:linear-gradient(180deg,#1a2c4e,#14213d);color:#fff;
           position:sticky;top:0;height:100vh;overflow-y:auto}
  .sb-logo{display:flex;align-items:center;gap:11px;padding:20px 18px;border-bottom:1px solid rgba(255,255,255,.1)}
  .sb-mark{width:38px;height:38px;border-radius:9px;background:var(--gold);display:flex;align-items:center;
           justify-content:center;font-weight:800;font-size:19px;color:#1a2c4e}
  .sb-logo h1{font-size:15px;line-height:1.1;color:#fff}
  .sb-logo p{font-size:9px;color:#c9d4e8;letter-spacing:1.5px;margin-top:2px}
  .sb-sec{font-size:10px;letter-spacing:1.5px;color:#8a9bc0;padding:16px 18px 7px;font-weight:700}
  .sb-item{display:flex;align-items:center;gap:10px;padding:10px 18px;font-size:13px;color:#dbe3f0;
           cursor:pointer;border-left:3px solid transparent}
  .sb-item:hover{background:rgba(255,255,255,.06)}
  .sb-item.on{background:rgba(184,148,90,.16);border-left-color:var(--gold);color:#fff;font-weight:700}
  .sb-item .ic{width:16px;text-align:center;opacity:.9;font-size:12px}
  .main{flex:1;min-width:0}
  .topbar{background:#fff;border-bottom:1px solid var(--line);padding:13px 26px;display:flex;
          align-items:center;gap:14px;position:sticky;top:0;z-index:10;flex-wrap:wrap}
  .topbar .meta{font-size:12px;color:var(--muted);text-align:right;line-height:1.6}
  .topbar .meta b{color:var(--ink)}
  .content{padding:20px 26px 50px;max-width:1400px}
  .selbar{background:#fff;border:1px solid var(--line);border-radius:12px;padding:13px 18px;margin-bottom:18px;
          box-shadow:0 1px 3px rgba(26,44,78,.06)}
  .selrow{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .sellbl{font-size:12px;color:var(--muted);font-weight:700;min-width:42px}
  .chips{display:flex;gap:6px;flex-wrap:wrap}
  .mchip{border:1px solid var(--line);background:#fff;color:var(--ink);border-radius:20px;padding:6px 12px;
         font-size:12px;cursor:pointer;font-weight:600;user-select:none}
  .mchip.on{background:var(--gold);color:#1a2c4e;border-color:var(--gold);font-weight:800}
  .selhint{font-size:11px;color:var(--muted);margin-top:9px}

  /* ---- Giriş ekranı ---- */
  #app{display:none}
  #login{position:fixed;inset:0;z-index:100;background:radial-gradient(1200px 600px at 50% -10%,#22375f,#14213d);
         display:flex;align-items:center;justify-content:center;padding:20px}
  .login-card{width:100%;max-width:380px;background:var(--panel);border:1px solid var(--line);
              border-top:3px solid var(--gold);border-radius:16px;padding:34px 30px;
              box-shadow:0 20px 60px rgba(0,0,0,.35)}
  .login-card .brandrow{display:flex;align-items:center;gap:13px;margin-bottom:6px}
  .login-card .mark{width:44px;height:44px;border-radius:10px;background:var(--brand);
        display:flex;align-items:center;justify-content:center;font-weight:800;font-size:22px;color:var(--gold)}
  .login-card h1{font-size:17px;letter-spacing:.3px;color:var(--ink)}
  .login-card p.sub{font-size:11px;color:var(--muted);letter-spacing:2px;margin-top:2px}
  .login-card h2{font-size:14px;color:var(--muted);font-weight:600;margin:22px 0 16px;text-align:center}
  .login-card label{display:block;font-size:12px;color:var(--muted);margin:12px 0 6px}
  .login-card input{width:100%;background:var(--panel2);border:1px solid var(--line);border-radius:9px;
        padding:12px 14px;font-size:14px;color:var(--ink);font-family:inherit}
  .login-card input:focus{outline:2px solid var(--brand)}
  .login-card button{width:100%;margin-top:22px;background:var(--brand);color:#fff;border:0;border-radius:9px;
        padding:13px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit}
  .login-card button:hover{background:var(--brand2)}
  .login-err{display:none;margin-top:14px;font-size:12.5px;color:var(--bad);text-align:center}
  .login-foot{margin-top:20px;font-size:10px;color:var(--muted);text-align:center;letter-spacing:1px}
  .logout{cursor:pointer;font-size:11px;color:var(--muted);border:1px solid var(--line);
          background:var(--panel2);padding:7px 12px;border-radius:9px}
  .logout:hover{color:var(--ink)}
</style>
</head>
<body>

<!-- ================= GİRİŞ EKRANI ================= -->
<div id="login">
  <div class="login-card">
    <div class="brandrow">
      <div class="mark">İ</div>
      <div>
        <h1>İNCİROĞLU OTOMOTİV</h1>
        <p class="sub">SERVİS GİDİŞAT DASHBOARD</p>
      </div>
    </div>
    <h2>Panele erişim için giriş yapın</h2>
    <label>Kullanıcı Adı</label>
    <input id="l_user" type="text" autocomplete="username" autocapitalize="none">
    <label>Şifre</label>
    <input id="l_pass" type="password" autocomplete="current-password">
    <button id="l_btn">Giriş Yap</button>
    <div class="login-err" id="l_err">Kullanıcı adı veya şifre hatalı.</div>
    <div class="login-foot">YETKİLİ ERİŞİM · İNCİROĞLU OTOMOTİV</div>
  </div>
</div>

<!-- ================= DASHBOARD (giriş sonrası) ================= -->
<div id="app">
<div class="layout">
  <!-- SOL MENÜ -->
  <aside class="sidebar">
    <div class="sb-logo">
      <div class="sb-mark">İ</div>
      <div><h1>İNCİROĞLU</h1><p>SERVİS PANELİ</p></div>
    </div>
    <div class="sb-sec">RAPORLAR</div>
    <div id="sbBolum">
      <div class="sb-item on" data-b="servis"><span class="ic">▮</span> Servis Gidişat</div>
      <div class="sb-item" data-b="yedekparca"><span class="ic">◫</span> Yedek Parça</div>
      <div class="sb-item" data-b="tahmin"><span class="ic">◈</span> Kapanış Tahmini</div>
      <div class="sb-item" data-b="homer"><span class="ic">✚</span> HOMER / Hasar</div>
      <div class="sb-item" data-b="ozet"><span class="ic">▤</span> Yönetici Özet</div>
    </div>
    <div class="sb-sec">MARKA</div>
    <div id="sbMarka"></div>
    <div style="padding:16px 18px;margin-top:10px">
      <div class="logout" id="logoutBtn" title="Çıkış yap">Çıkış Yap</div>
    </div>
  </aside>

  <!-- İÇERİK -->
  <div class="main">
    <div class="topbar">
      <div class="tag" id="donemTag"></div>
      <div class="spacer"></div>
      <div class="meta" id="meta"></div>
    </div>

    <div class="content">
      <!-- ORTAK: AY / ÇEYREK SEÇİCİ -->
      <div class="selbar">
        <div class="selrow" style="margin-bottom:10px">
          <span class="sellbl">Hızlı:</span>
          <div class="seg" id="hizliSeg">
            <button data-k="buay">Bu Ay</button>
            <button data-k="ytd">YTD</button>
            <button data-k="Q1">Q1</button>
            <button data-k="Q2">Q2</button>
            <button data-k="Q3">Q3</button>
            <button data-k="Q4">Q4</button>
          </div>
        </div>
        <div class="selrow">
          <span class="sellbl">Aylar:</span>
          <div id="ayChips" class="chips"></div>
        </div>
        <div class="selhint">Birden çok ay seçebilirsin (örn. Şubat + Mart) — seçilenler toplanır. YTD: Ocak'tan son seçili aya kadar.</div>
      </div>

      <!-- ===== BÖLÜM: SERVİS GİDİŞAT ===== -->
      <div id="bolum-servis" class="bolum">
        <div class="period-bar" id="periodBar"></div>

        <h2 class="sec">DÖNEM ÖZETİ · CİRO & İŞ EMRİ</h2>
        <div class="cards" id="kpiCards"></div>

        <h2 class="sec">ARAÇ BAŞI CİRO <span style="font-weight:400;letter-spacing:0;text-transform:none;font-size:12px;color:var(--muted)">· Mekanik işçilik + Mekanik YP ÷ Mekanik araç adedi</span></h2>
        <div class="cards" id="aracCards"></div>

        <h2 class="sec">GÜNLÜK TEMPO & BÜTÇE HEDEFİ <span id="tempoSub" style="font-weight:400;letter-spacing:0;text-transform:none;font-size:12px;color:var(--muted)"></span></h2>
        <div class="cards" id="tempoCards"></div>

        <h2 class="sec">DETAY
          <div class="seg metric-seg" id="metricSeg">
            <button data-m="toplam" class="on">Toplam Ciro</button>
            <button data-m="iscilik">İşçilik</button>
            <button data-m="yp">Yedek Parça</button>
            <button data-m="adet">İş Emri Adet</button>
          </div>
        </h2>

        <div class="grid2">
          <div class="box">
            <h3>Aylık Trend <span id="trendSub">· bütçe vs fiili</span></h3>
            <canvas id="trendChart"></canvas>
          </div>
          <div class="box">
            <h3>Kategori Payı <span id="katSub">· seçili dönem</span></h3>
            <canvas id="katChart"></canvas>
          </div>
        </div>

        <div class="box" style="margin-top:16px">
          <h3>Kategori Kırılımı <span id="katTblSub">· toplam ciro</span></h3>
          <div style="overflow-x:auto"><table id="katTable"></table></div>
        </div>

        <div class="box" style="margin-top:16px">
          <h3>Marka Karşılaştırma <span id="mkSub">· seçili dönem</span></h3>
          <div style="overflow-x:auto"><table id="markaTable"></table></div>
        </div>
      </div>

      <!-- ===== DİĞER BÖLÜMLER (yakında) ===== -->
      <div id="bolum-yedekparca" class="bolum" style="display:none">
        <div class="box"><h3>Yedek Parça</h3><p style="color:var(--muted);font-size:13px;margin-top:8px">Bu bölüm hazırlanıyor.</p></div>
      </div>
      <div id="bolum-tahmin" class="bolum" style="display:none">
        <div class="box"><h3>Kapanış Tahmini</h3><p style="color:var(--muted);font-size:13px;margin-top:8px">Bu bölüm hazırlanıyor.</p></div>
      </div>
      <div id="bolum-homer" class="bolum" style="display:none">
        <div class="box"><h3>HOMER / Hasar</h3><p style="color:var(--muted);font-size:13px;margin-top:8px">Bu bölüm hazırlanıyor.</p></div>
      </div>
      <div id="bolum-ozet" class="bolum" style="display:none">
        <div class="box"><h3>Yönetici Özet</h3><p style="color:var(--muted);font-size:13px;margin-top:8px">Bu bölüm hazırlanıyor.</p></div>
      </div>

      <div class="foot">
        <b>İnciroğlu Otomotiv</b> · Servis Paneli · Üretim: <span id="uretim"></span>
      </div>
    </div>
  </div>
</div>

<script>
// ================= GİRİŞ KONTROLÜ =================
const _KUL_HASH = "__KUL_HASH__";
const _SIF_HASH = "__SIF_HASH__";
async function _sha256(s){
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}
function _girisAc(){
  document.getElementById('login').style.display='none';
  document.getElementById('app').style.display='block';
  if(window.__dashboardInit) window.__dashboardInit();
}
async function _girisDene(){
  const u = document.getElementById('l_user').value.trim().toLowerCase();
  const p = document.getElementById('l_pass').value;
  const uh = await _sha256(u), ph = await _sha256(p);
  if(uh===_KUL_HASH && ph===_SIF_HASH){
    sessionStorage.setItem('inc_auth','1');
    _girisAc();
  } else {
    document.getElementById('l_err').style.display='block';
    document.getElementById('l_pass').value='';
  }
}
document.addEventListener('DOMContentLoaded',()=>{
  // Sekme açıkken tekrar sorma
  if(sessionStorage.getItem('inc_auth')==='1'){ _girisAc(); }
  document.getElementById('l_btn').addEventListener('click',_girisDene);
  document.getElementById('l_pass').addEventListener('keydown',e=>{if(e.key==='Enter')_girisDene();});
  document.getElementById('l_user').addEventListener('keydown',e=>{if(e.key==='Enter')document.getElementById('l_pass').focus();});
  const lo=document.getElementById('logoutBtn');
  if(lo) lo.addEventListener('click',()=>{ sessionStorage.removeItem('inc_auth'); location.reload(); });
});
// ==================================================

const DATA = __DATA__;

const TL = n => new Intl.NumberFormat('tr-TR',{maximumFractionDigits:0}).format(Math.round(n||0));
const ADET = n => new Intl.NumberFormat('tr-TR',{maximumFractionDigits:0}).format(Math.round(n||0));
const PCT = n => (n==null?'–':(n*100).toFixed(1).replace('.',',')+'%');
const DPCT = n => (n==null?'':(n>=0?'+':'')+(n*100).toFixed(1).replace('.',',')+'%');
const cls = n => n>0.0001?'up':(n<-0.0001?'dn':'nt');
const METRIK_AD = {toplam:'Toplam Ciro', iscilik:'İşçilik', yp:'Yedek Parça', adet:'İş Emri Adet'};
const isAdet = m => m==='adet';
const fmt = (m,v) => isAdet(m)? ADET(v) : TL(v)+' ₺';

// Çoklu ay: state.aylar = seçili ay adları dizisi. mode: 'aylar' (seçili aylar toplanır) | 'ytd' (Ocak→son seçili ay)
let state = {
  bolum: 'servis',
  marka: DATA.marka_sirasi[0],
  metric: 'toplam',
  mode: 'aylar',
  aylar: [ (DATA.ust.secili_ay||'TEMMUZ').toUpperCase() ]
};
let charts = {};

const AYLAR_JS = ['OCAK','ŞUBAT','MART','NİSAN','MAYIS','HAZİRAN','TEMMUZ','AĞUSTOS','EYLÜL','EKİM','KASIM','ARALIK'];
const AYK = {OCAK:'Oca',ŞUBAT:'Şub',MART:'Mar','NİSAN':'Nis',MAYIS:'May',HAZİRAN:'Haz',TEMMUZ:'Tem','AĞUSTOS':'Ağu','EYLÜL':'Eyl',EKİM:'Eki',KASIM:'Kas',ARALIK:'Ara'};

// Aktif ay listesini döndürür (mode'a göre)
function aktifAylar(){
  if(state.mode==='ytd'){
    // Ocak'tan, seçili ayların EN SONuncusuna kadar
    let maxIdx = 0;
    state.aylar.forEach(a=>{ maxIdx = Math.max(maxIdx, AYLAR_JS.indexOf(a)); });
    return AYLAR_JS.slice(0, maxIdx+1);
  }
  // seçili aylar (kronolojik sırada)
  return AYLAR_JS.filter(a=>state.aylar.includes(a));
}

// Verilen ay listesini toplayan blok üretir
function aylariTopla(markaKey, ayListesi){
  const v = DATA.markalar[markaKey].veri;
  const metrikler = ['adet','iscilik','yp','toplam'];
  const alanlar = ['butce','fiili','g2025'];
  const acc = {toplam:{}, kategoriler:{}};
  metrikler.forEach(m=>{acc.toplam[m]={butce:0,fiili:0,g2025:0};});
  DATA.kategoriler.forEach(k=>{acc.kategoriler[k]={}; metrikler.forEach(m=>{acc.kategoriler[k][m]={butce:0,fiili:0,g2025:0};});});
  ayListesi.forEach(ay=>{
    const b = v.aylar[ay]; if(!b) return;
    metrikler.forEach(m=>alanlar.forEach(a=>{ acc.toplam[m][a]+=b.toplam[m][a]||0; }));
    DATA.kategoriler.forEach(k=>metrikler.forEach(m=>alanlar.forEach(a=>{
      acc.kategoriler[k][m][a]+=(b.kategoriler[k]?b.kategoriler[k][m][a]:0)||0;
    })));
  });
  metrikler.forEach(m=>{
    const t=acc.toplam[m]; t.delta = t.g2025>0? t.fiili/t.g2025-1 : 0;
    DATA.kategoriler.forEach(k=>{const c=acc.kategoriler[k][m]; c.delta=c.g2025>0?c.fiili/c.g2025-1:0;});
  });
  return acc;
}

function blok(markaKey){
  return aylariTopla(markaKey, aktifAylar());
}

// Aktif dönemin iş günü toplamı {donem,gecen,kalan}
function aktifIsGunu(){
  const ag = DATA.ust.ay_is_gunu || {};
  let d=0,g=0,k=0;
  aktifAylar().forEach(ay=>{ const x=ag[ay]; if(x){d+=x.donem;g+=x.gecen;k+=x.kalan;} });
  return {donem:d,gecen:g,kalan:k};
}

// Aktif dönem etiketi (örn "Şubat+Mart" veya "Ocak–Ağustos (YTD)")
function donemEtiket(){
  const aylar = aktifAylar();
  if(state.mode==='ytd') return 'Ocak – '+(AYK[aylar[aylar.length-1]]||'');
  if(state.aylar.length===1) return state.aylar[0].charAt(0)+state.aylar[0].slice(1).toLowerCase();
  return state.aylar.map(a=>AYK[a]).join(' + ');
}

function renderMeta(){
  const u = DATA.ust;
  document.getElementById('meta').innerHTML =
    `Dönem: <b>${donemEtiket()}</b><br>Veri kesim: <b>${u.veri_kesim||'-'}</b>`;
  document.getElementById('uretim').textContent = DATA.uretim;
  const ig = aktifIsGunu();
  const donem=ig.donem, gecen=ig.gecen, kalan=ig.kalan;
  const y = donem>0 ? (gecen/donem):0;
  document.getElementById('periodBar').innerHTML = `
    <div class="pi">Seçili Dönem<b>${donemEtiket()}</b></div>
    <div class="pi">${state.mode==='ytd'?'Toplam İş Günü':'Dönem İş Günü'}<b>${ADET(donem)}</b></div>
    <div class="pi">Geçen<b>${ADET(gecen)}</b></div>
    <div class="pi">Kalan<b>${ADET(kalan)}</b></div>
    <div class="prog"><div style="font-size:12px;color:var(--muted)">Dönem ilerlemesi · ${PCT(y)}</div>
      <div class="bar"><div class="fill" style="width:${(y*100).toFixed(0)}%"></div></div></div>`;
  document.getElementById('donemTag').textContent =
    DATA.markalar[state.marka].etiket + ' · ' + donemEtiket();
}

function kartCiro(title, m, unit, alt){
  const f=m.fiili, b=m.butce, oran=b>0?f/b:null, d=m.delta;
  const val = unit==='adet'? ADET(f) : TL(f)+' ₺';
  const budget = unit==='adet'? ADET(b) : TL(b)+' ₺';
  const pill = oran==null?'nt':(oran>=1?'up':'dn');
  return `<div class="card${alt?' alt':''}">
    <div class="t">${title}</div>
    <div class="v">${val}</div>
    <span class="pill ${pill}">Bütçe %${b>0?(oran*100).toFixed(0):'–'}</span>
    <div class="kv"><span>Bütçe</span><b>${budget}</b></div>
    <div class="kv"><span>2025 → Δ</span><b style="color:var(--${d>=0?'ok':'bad'})">${DPCT(d)}</b></div>
  </div>`;
}
function renderCards(){
  const bl = blok(state.marka).toplam;
  document.getElementById('kpiCards').innerHTML =
    kartCiro('TOPLAM CİRO', bl.toplam,'tl') +
    kartCiro('İŞÇİLİK CİROSU', bl.iscilik,'tl') +
    kartCiro('YEDEK PARÇA CİROSU', bl.yp,'tl') +
    kartCiro('İŞ EMRİ ADET', bl.adet,'adet');
}

// araç başı kartı: fiili & 2025 değerleri doğrudan verilir
function kartAracV(title, perF, per25){
  const d = per25>0? perF/per25 - 1 : null;
  return `<div class="card alt">
    <div class="t">${title}</div>
    <div class="v">${TL(perF)} ₺</div>
    <span class="pill nt">araç başı</span>
    <div class="kv"><span>2025 araç başı</span><b>${TL(per25)} ₺</b></div>
    <div class="kv"><span>Δ</span><b style="color:var(--${d==null?'muted':(d>=0?'ok':'bad')})">${d==null?'–':DPCT(d)}</b></div>
  </div>`;
}
function renderArac(){
  const b = blok(state.marka);
  const mek = b.kategoriler['Mekanik'];
  const mAdet = mek.adet.fiili, mAdet25 = mek.adet.g2025;
  const toplamAdet = b.toplam.adet.fiili;

  // Hem aylık hem YTD: Mekanik işçilik / YP ÷ Mekanik araç adedi (YTD'de toplanan bloktan)
  const iscF = mAdet>0? mek.iscilik.fiili/mAdet : 0;
  const isc25 = mAdet25>0? mek.iscilik.g2025/mAdet25 : 0;
  const ypF = mAdet>0? mek.yp.fiili/mAdet : 0;
  const yp25 = mAdet25>0? mek.yp.g2025/mAdet25 : 0;
  const topF = iscF + ypF, top25 = isc25 + yp25;

  const dortuncu = (state.mode==='ytd')
    ? `<div class="card alt"><div class="t">MEKANİK ARAÇ (YTD)</div>
         <div class="v">${ADET(mAdet)}</div><span class="pill nt">mekanik adet</span>
         <div class="kv"><span>2025 mekanik</span><b>${ADET(mAdet25)}</b></div>
         <div class="kv"><span>Δ</span><b style="color:var(--${mek.adet.delta>=0?'ok':'bad'})">${DPCT(mek.adet.delta)}</b></div>
       </div>`
    : `<div class="card alt"><div class="t">MEKANİK ARAÇ GİRİŞİ</div>
         <div class="v">${ADET(mAdet)}</div><span class="pill nt">mekanik adet · toplam ${ADET(toplamAdet)}</span>
         <div class="kv"><span>2025 mekanik</span><b>${ADET(mAdet25)}</b></div>
         <div class="kv"><span>Δ</span><b style="color:var(--${mek.adet.delta>=0?'ok':'bad'})">${DPCT(mek.adet.delta)}</b></div>
       </div>`;

  document.getElementById('aracCards').innerHTML =
    kartAracV('ARAÇ BAŞI TOPLAM', topF, top25) +
    kartAracV('ARAÇ BAŞI İŞÇİLİK', iscF, isc25) +
    kartAracV('ARAÇ BAŞI YEDEK PARÇA', ypF, yp25) +
    dortuncu;
}

// Günlük tempo & bütçe hedefi — HER ZAMAN seçili ayın verisini ve o ayın iş gününü kullanır
function renderTempo(){
  const u = DATA.ust;
  // Seçili ayın KENDİ iş günü (her ay farklı)
  const ig = aktifIsGunu();
  const gecen = ig.gecen, kalan = ig.kalan, donem = ig.donem;
  // aktif dönem toplamı
  const t = blok(state.marka).toplam;

  function tempoKart(baslik, metrikBlok, birim){
    const butce = metrikBlok.butce, fiili = metrikBlok.fiili;
    const f = birim==='adet' ? (x=>ADET(x)) : (x=>TL(x)+' ₺');

    const gunlukFiili   = gecen>0 ? fiili/gecen : 0;       // fiili günlük ortalama
    const gunlukHedef   = donem>0 ? butce/donem : 0;       // bütçe günlük hedef (sabit referans)
    const olmasiGereken = gunlukHedef * gecen;             // bugüne kadar olması gereken (kümülatif)
    const fark          = fiili - olmasiGereken;           // + önde / − geride
    const kalanHedef    = Math.max(butce - fiili, 0);
    const gerekenGunluk = kalan>0 ? kalanHedef/kalan : 0;  // kalan günlerde gereken/gün

    const onde = fark >= 0;
    const durumYazi = kalanHedef<=0 ? 'Bütçe tamam ✓' : (onde ? 'Bütçenin önünde ✓' : 'Bütçenin gerisinde');
    const durumPill = kalanHedef<=0 ? 'up' : (onde ? 'up' : 'dn');
    return `<div class="card">
      <div class="t">${baslik}</div>
      <div class="v">${f(gunlukFiili)}<span style="font-size:13px;color:var(--muted);font-weight:600"> /gün fiili</span></div>
      <span class="pill ${durumPill}">${durumYazi}</span>
      <div class="kv"><span>Bütçe günlük hedef</span><b>${f(gunlukHedef)}</b></div>
      <div class="kv"><span>Bugüne dek olması gereken</span><b>${f(olmasiGereken)}</b></div>
      <div class="kv"><span>Fiili / Fark</span><b style="color:var(--${onde?'ok':'bad'})">${f(fiili)} · ${onde?'+':''}${f(fark)}</b></div>
      <div class="kv"><span>Kalanda gereken/gün (${ADET(kalan)}g)</span><b>${f(gerekenGunluk)}</b></div>
    </div>`;
  }

  document.getElementById('tempoCards').innerHTML =
    tempoKart('GÜNLÜK ARAÇ GİRİŞİ', t.adet, 'adet') +
    tempoKart('GÜNLÜK TOPLAM CİRO', t.toplam, 'tl') +
    tempoKart('GÜNLÜK İŞÇİLİK', t.iscilik, 'tl') +
    tempoKart('GÜNLÜK YEDEK PARÇA', t.yp, 'tl');

  document.getElementById('tempoSub').textContent =
    `· ${donemEtiket()} · ${ADET(donem)} iş günü (Cmt dahil, resmi tatil hariç) · ${ADET(gecen)} geçti, ${ADET(kalan)} kaldı`;
}

function renderKatTable(){
  const m = state.metric, bl = blok(state.marka);
  let max=0;
  DATA.kategoriler.forEach(k=>{ max=Math.max(max, bl.kategoriler[k][m].fiili); });
  let rows = DATA.kategoriler.map(k=>{
    const t = bl.kategoriler[k][m];
    const oran = t.butce>0? t.fiili/t.butce:null;
    const w = max>0?(t.fiili/max*100):0;
    return `<tr>
      <td>${k}</td><td>${fmt(m,t.butce)}</td>
      <td class="barcell"><div class="b" style="width:${w}%"></div><span>${fmt(m,t.fiili)}</span></td>
      <td><span class="mini ${oran==null?'nt':(oran>=1?'up':'dn')}">${oran==null?'–':(oran*100).toFixed(0)+'%'}</span></td>
      <td style="color:var(--${t.delta>=0?'ok':'bad'})">${DPCT(t.delta)}</td>
    </tr>`;
  }).join('');
  const tt = bl.toplam[m], oranT = tt.butce>0?tt.fiili/tt.butce:null;
  rows += `<tr class="tot"><td>TOPLAM</td><td>${fmt(m,tt.butce)}</td><td style="text-align:right">${fmt(m,tt.fiili)}</td>
    <td>${oranT==null?'–':(oranT*100).toFixed(0)+'%'}</td>
    <td style="color:var(--${tt.delta>=0?'ok':'bad'})">${DPCT(tt.delta)}</td></tr>`;
  document.getElementById('katTable').innerHTML =
    `<thead><tr><th>Kategori</th><th>Bütçe</th><th>Fiili</th><th>%</th><th>2025 Δ</th></tr></thead><tbody>${rows}</tbody>`;
  document.getElementById('katTblSub').textContent = '· '+METRIK_AD[m].toLowerCase();
}

function renderMarkaTable(){
  const m = state.metric;
  const others = DATA.marka_sirasi.filter(s=>s!=='Servis Konsolide Rapor');
  let toplamFiili=0;
  const rows0 = others.map(s=>{
    const t = blok(s).toplam[m];
    toplamFiili += t.fiili;
    return {s, etiket:DATA.markalar[s].etiket, t};
  }).sort((a,b)=>b.t.fiili-a.t.fiili);
  const rows = rows0.map(r=>{
    const oran=r.t.butce>0?r.t.fiili/r.t.butce:null;
    const pay=toplamFiili>0?r.t.fiili/toplamFiili:0;
    return `<tr${r.s===state.marka?' style="background:var(--panel2)"':''}>
      <td>${r.etiket}</td><td>${fmt(m,r.t.fiili)}</td>
      <td><span class="mini ${oran==null?'nt':(oran>=1?'up':'dn')}">${oran==null?'–':(oran*100).toFixed(0)+'%'}</span></td>
      <td style="color:var(--${r.t.delta>=0?'ok':'bad'})">${DPCT(r.t.delta)}</td>
      <td>${(pay*100).toFixed(1).replace('.',',')}%</td>
    </tr>`;
  }).join('');
  document.getElementById('markaTable').innerHTML =
    `<thead><tr><th>Marka</th><th>Fiili</th><th>Bütçe%</th><th>2025 Δ</th><th>Pay</th></tr></thead><tbody>${rows}</tbody>
     <tfoot><tr class="tot"><td>TOPLAM</td><td style="text-align:right">${fmt(m,toplamFiili)}</td><td></td><td></td><td>100%</td></tr></tfoot>`;
  document.getElementById('mkSub').textContent = '· '+METRIK_AD[m].toLowerCase();
}

function renderTrend(){
  const m = state.metric, v = DATA.markalar[state.marka].veri;
  const labels = DATA.aylar.map(a=>a.substring(0,3));
  const butce = DATA.aylar.map(a=>v.aylar[a].toplam[m].butce);
  const fiili = DATA.aylar.map(a=>v.aylar[a].toplam[m].fiili);
  const selSet = new Set(aktifAylar());
  if(charts.trend) charts.trend.destroy();
  charts.trend = new Chart(document.getElementById('trendChart'),{
    data:{labels,datasets:[
      {type:'bar',label:'Fiili',data:fiili,backgroundColor:DATA.aylar.map(a=>selSet.has(a)?'#b8945a':'#c9d0dc'),borderRadius:4,order:2},
      {type:'line',label:'Bütçe',data:butce,borderColor:'#1a2c4e',backgroundColor:'#1a2c4e',borderWidth:2,pointRadius:2,tension:.3,order:1}
    ]},
    options:{responsive:true,plugins:{legend:{labels:{color:'#6b7280',boxWidth:12}},
      tooltip:{callbacks:{label:c=>c.dataset.label+': '+fmt(m,c.raw)}}},
      scales:{x:{ticks:{color:'#6b7280'},grid:{display:false}},
              y:{ticks:{color:'#6b7280',callback:v=>isAdet(m)?ADET(v):TL(v/1e6)+'M'},grid:{color:'#ddd3c0'}}}}
  });
  document.getElementById('trendSub').textContent='· bütçe vs fiili · '+METRIK_AD[m].toLowerCase();
}

function renderKatChart(){
  const m = state.metric, bl = blok(state.marka);
  const data = DATA.kategoriler.map(k=>bl.kategoriler[k][m].fiili);
  const pal = ['#1a2c4e','#b8945a','#3a5a8c','#8a9bb5','#c2a878','#5a6d8c','#d9cbb0'];
  if(charts.kat) charts.kat.destroy();
  charts.kat = new Chart(document.getElementById('katChart'),{
    type:'doughnut',
    data:{labels:DATA.kategoriler,datasets:[{data,backgroundColor:pal,borderColor:'#ffffff',borderWidth:2}]},
    options:{responsive:true,cutout:'58%',plugins:{legend:{position:'right',labels:{color:'#6b7280',boxWidth:12,padding:8,font:{size:11}}},
      tooltip:{callbacks:{label:c=>c.label+': '+fmt(m,c.raw)}}}}
  });
  document.getElementById('katSub').textContent='· seçili dönem · '+METRIK_AD[m].toLowerCase();
}


function renderAll(){
  // Aktif bölümü göster, diğerlerini gizle
  ['servis','yedekparca','tahmin','homer','ozet'].forEach(b=>{
    const el=document.getElementById('bolum-'+b);
    if(el) el.style.display = (b===state.bolum)?'block':'none';
  });
  // Metrik seçici sadece servis bölümünde görünür
  document.getElementById('uretim').textContent = DATA.uretim;
  document.getElementById('donemTag').textContent = DATA.markalar[state.marka].etiket + ' · ' + donemEtiket();

  if(state.bolum==='servis'){
    const fns=[renderMeta,renderCards,renderArac,renderTempo,renderKatTable,renderMarkaTable,renderTrend,renderKatChart];
    fns.forEach(fn=>{ try{ fn(); }catch(e){ console.error(fn.name,e); } });
  }
  // diğer bölümler yakında
}

const QCEYREK = { Q1:['OCAK','ŞUBAT','MART'], Q2:['NİSAN','MAYIS','HAZİRAN'], Q3:['TEMMUZ','AĞUSTOS','EYLÜL'], Q4:['EKİM','KASIM','ARALIK'] };

function ayChipleriCiz(){
  const box = document.getElementById('ayChips');
  box.innerHTML = AYLAR_JS.map(a=>{
    const on = state.aylar.includes(a) ? ' on' : '';
    return `<span class="mchip${on}" data-ay="${a}">${AYK[a]}</span>`;
  }).join('');
  box.querySelectorAll('.mchip').forEach(c=>c.addEventListener('click',()=>{
    const ay=c.dataset.ay;
    if(state.aylar.includes(ay)){
      if(state.aylar.length>1) state.aylar=state.aylar.filter(x=>x!==ay); // en az 1 kalsın
    } else {
      state.aylar=[...state.aylar, ay];
    }
    state.mode='aylar';
    guncelleSeciciler(); renderAll();
  }));
}

function guncelleSeciciler(){
  ayChipleriCiz();
  // Q ve mod butonlarının aktifliği
  document.querySelectorAll('#hizliSeg button').forEach(b=>{
    const k=b.dataset.k; let aktif=false;
    if(k==='ytd') aktif = (state.mode==='ytd');
    else if(QCEYREK[k]) aktif = (state.mode==='aylar' && state.aylar.length===3 && QCEYREK[k].every(a=>state.aylar.includes(a)));
    b.classList.toggle('on', aktif);
  });
}

function initControls(){
  // Sol menü — bölümler
  document.querySelectorAll('#sbBolum .sb-item').forEach(it=>it.addEventListener('click',()=>{
    document.querySelectorAll('#sbBolum .sb-item').forEach(x=>x.classList.remove('on'));
    it.classList.add('on'); state.bolum=it.dataset.b; renderAll();
  }));
  // Sol menü — markalar
  const mbox = document.getElementById('sbMarka');
  mbox.innerHTML = DATA.marka_sirasi.map((s,i)=>{
    const on = s===state.marka ? ' on' : '';
    return `<div class="sb-item${on}" data-mk="${s}"><span class="ic">${s==='Servis Konsolide Rapor'?'●':'○'}</span> ${DATA.markalar[s].etiket}</div>`;
  }).join('');
  mbox.querySelectorAll('.sb-item').forEach(it=>it.addEventListener('click',()=>{
    mbox.querySelectorAll('.sb-item').forEach(x=>x.classList.remove('on'));
    it.classList.add('on'); state.marka=it.dataset.mk; renderAll();
  }));

  // Hızlı seçim: Bu Ay / YTD / Q1-Q4
  document.querySelectorAll('#hizliSeg button').forEach(b=>b.addEventListener('click',()=>{
    const k=b.dataset.k;
    if(k==='buay'){ state.mode='aylar'; state.aylar=[ (DATA.ust.secili_ay||'AĞUSTOS').toUpperCase() ]; }
    else if(k==='ytd'){ state.mode='ytd'; if(state.aylar.length===0) state.aylar=[(DATA.ust.secili_ay||'AĞUSTOS').toUpperCase()]; }
    else if(QCEYREK[k]){ state.mode='aylar'; state.aylar=[...QCEYREK[k]]; }
    guncelleSeciciler(); renderAll();
  }));

  // Metrik seçici
  document.querySelectorAll('#metricSeg button').forEach(b=>b.addEventListener('click',()=>{
    document.querySelectorAll('#metricSeg button').forEach(x=>x.classList.remove('on'));
    b.classList.add('on'); state.metric=b.dataset.m; renderAll();
  }));

  guncelleSeciciler();
}
// Dashboard'u giriş başarılı olunca başlat (giriş öncesi çizim yapma)
window.__dashboardInit = function(){
  if(window.__inited) return;
  window.__inited = true;
  initControls();
  renderAll();
};
// Eğer zaten giriş yapılmışsa (sessionStorage), DOMContentLoaded içinde _girisAc çağrılır
// ve o da __dashboardInit'i tetikler.
</script>
</body>
</html>"""


# ==================================================================== #
#  Akışlı sayfa okuyucu (büyük sayfalar için, belleği şişirmez)
# ==================================================================== #
def _sayfa_akis(dosya, sayfa_adi, sutunlar=None):
    """Bir sayfayı satır satır okur -> (satir_no, {sütun: değer}). Sayfa yoksa hiçbir şey döndürmez."""
    import zipfile as _zip
    import xml.etree.ElementTree as _ET
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    z = _zip.ZipFile(dosya)
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', z.read("xl/_rels/workbook.xml.rels").decode("utf-8")))
    hedef = None
    for t in re.findall(r"<sheet\b[^>]*/>", wbx):
        nm = re.search(r'name="([^"]+)"', t); rid = re.search(r'r:id="(rId\d+)"', t)
        if nm and rid and nm.group(1) == sayfa_adi:
            yol = rels.get(rid.group(1), "").lstrip("/")
            hedef = yol if yol.startswith("xl/") else "xl/" + yol
    if not hedef:
        z.close(); return
    ortak = []
    if "xl/sharedStrings.xml" in z.namelist():
        for ev, el in _ET.iterparse(z.open("xl/sharedStrings.xml")):
            if el.tag == NS + "si":
                ortak.append("".join(t.text or "" for t in el.iter(NS + "t"))); el.clear()
    for ev, el in _ET.iterparse(z.open(hedef)):
        if el.tag == NS + "row":
            satir = {}
            for c in el.iter(NS + "c"):
                sut = re.match(r"[A-Z]+", c.get("r", "")).group(0)
                if sutunlar and sut not in sutunlar:
                    continue
                v = c.find(NS + "v")
                if v is None or v.text is None:
                    ist = c.find(NS + "is")
                    deger = "".join(t.text or "" for t in ist.iter(NS + "t")) if ist is not None else None
                else:
                    deger = ortak[int(v.text)] if c.get("t") == "s" else v.text
                satir[sut] = deger
            yield int(el.get("r", "0")), satir
            el.clear()
    z.close()


def son_fatura_tarihi(dosya):
    """Fatura Listesi'ndeki en son fatura tarihi (F sütunu, 4. satırdan itibaren). Yoksa None."""
    en_buyuk = 0
    for no, r in _sayfa_akis(dosya, "Fatura Listesi", {"F"}):
        if no < 4:
            continue
        try:
            en_buyuk = max(en_buyuk, float(r.get("F") or 0))
        except ValueError:
            pass
    if en_buyuk <= 0:
        return None
    return _dt.date(1899, 12, 30) + _dt.timedelta(days=int(en_buyuk))


# ==================================================================== #
#  Ek satış (danışman bazında) — "Ek Satış Ham" sayfasından
# ==================================================================== #
EK_HAVUZ = {"ARJ SERVIS": "ARJ Servis Ortak Havuz",
            "FIAT SERVIS": "Fiat Servis Ortak Havuz",
            "DANISMANI BELIRLE": "Danışmanı Belirlenmemiş"}
EK_MARKA = {"bmw motorrad": "Motorrad", "motorrad": "Motorrad"}
EK_BUTCE_TURLERI = ["Aksesuar", "Lastik"]   # servis kategori bütçelerinden gelir
AY_ADLARI = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
             "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]

def _tr_ust(s):
    return s.replace("i", "İ").replace("ı", "I").upper()

def _isim_anahtar(s):
    return _tr_ust(" ".join(s.split())).translate(str.maketrans("ÇĞİÖŞÜ", "CGIOSU"))

def _tr_baslik(s):
    kelimeler = []
    for w in " ".join(s.split()).split(" "):
        k = "".join({"I": "ı", "İ": "i"}.get(c, c.lower()) for c in w)
        kelimeler.append(({"i": "İ", "ı": "I"}.get(k[0], k[0].upper()) + k[1:]) if k else k)
    return " ".join(kelimeler)


def ek_satis_uret(dosya, data):
    """Paneldeki EXTRA_DATA yapısını üretir. 'Ek Satış Ham' yoksa None döner."""
    from collections import Counter, defaultdict
    yazimlar = defaultdict(Counter)
    ham = []
    for no, r in _sayfa_akis(dosya, "Ek Satış Ham", {"A", "B", "C", "D", "E", "F"}):
        if no == 1:
            continue
        try:
            y, m, v = int(float(r["C"])), int(float(r["D"])), float(r.get("F") or 0)
        except (KeyError, TypeError, ValueError):
            continue
        ad = " ".join((r.get("A") or "").split()) or "DANIŞMANI BELİRLE"
        marka = (r.get("B") or "").strip()
        marka = EK_MARKA.get(marka.lower(), marka)
        tur = (r.get("E") or "").strip()
        anahtar = _isim_anahtar(ad)
        yazimlar[anahtar][ad] += 1
        ham.append((y, m, marka, anahtar, tur, v))
    if not ham:
        return None

    def gorunen(anahtar):
        if anahtar in EK_HAVUZ:
            return EK_HAVUZ[anahtar]
        en_cok = yazimlar[anahtar].most_common(1)[0][0]
        # Türkçe karakterli yazımı tercih et (AKÇA > AKCA)
        for yazim, _ in yazimlar[anahtar].most_common():
            if any(c in yazim for c in "ÇĞİÖŞÜçğıöşü"):
                en_cok = yazim; break
        return _tr_baslik(en_cok)

    ad_map = {a: gorunen(a) for a in yazimlar}
    havuzlar = set(EK_HAVUZ.values())
    top = defaultdict(float)
    for y, m, marka, a, tur, v in ham:
        top[(y, m, marka, ad_map[a], tur)] += v
    actual = [{"y": y, "m": m, "brand": b, "consultant": c, "type": t, "value": round(v, 2)}
              for (y, m, b, c, t), v in sorted(top.items()) if abs(v) > 0.004]

    yil = data["ust"].get("kesim_yil", _dt.date.today().year)
    aktif = defaultdict(set)
    for r in actual:
        if r["y"] == yil:
            aktif[(r["m"], r["brand"])].add(r["consultant"])
    marka_havuz = {"Arj": "ARJ Servis Ortak Havuz", "Fiat": "Fiat Servis Ortak Havuz"}

    def kadro(m, marka):
        if (m, marka) in aktif:                       # fiilisi olan ay: o ay çalışanlar
            k = aktif[(m, marka)] - havuzlar
        else:                                          # gelecek ay: yıl içinde çalışan herkes
            k = set().union(*[s for (mm, b), s in aktif.items() if b == marka] or [set()]) - havuzlar
        return k or {marka_havuz.get(marka, "Danışmanı Belirlenmemiş")}

    budgets, targets = [], []
    for marka, bilgi in data["markalar"].items():
        if marka == "Servis Konsolide Rapor":
            continue
        for mi, ay in enumerate(data["aylar"], start=1):
            kat = bilgi["veri"]["aylar"][ay]["kategoriler"]
            for tur in EK_BUTCE_TURLERI:
                b = kat.get(tur, {}).get("toplam", {}).get("butce", 0) or 0
                if b <= 0:
                    continue
                budgets.append({"y": yil, "m": mi, "brand": marka, "type": tur, "value": round(b, 2)})
                kisiler = sorted(kadro(mi, marka))
                for k in kisiler:
                    targets.append({"y": yil, "m": mi, "brand": marka, "consultant": k,
                                    "type": tur, "value": round(b / len(kisiler), 2)})

    return {
        "meta": {"sourceRows": len(ham), "sourceYear": yil, "sourceMonth": data["ust"].get("kesim_ay"),
                 "sourceHasDate": False,
                 "allocationMethod": "Eşit dağıtım · aynı marka ve ayda aktif danışmanlar",
                 "note": "Ham kaynakta yalnızca Yıl ve Ay bulunuyor; gerçek günlük işlem trendi için Fatura Tarihi eklenmeli."},
        "months": AY_ADLARI,
        "brands": sorted({r["brand"] for r in actual}),
        "types": sorted({r["type"] for r in actual}),
        "poolNames": sorted(havuzlar & {r["consultant"] for r in actual} | havuzlar & {t["consultant"] for t in targets}),
        "actual": actual, "budgets": budgets, "targets": targets,
    }


def panel_uret(data, sablon_yolu, cikti_yolu):
    """Yeni paneli (panel_sablon.html) Excel verisiyle doldurur.
    Sadece servis DATA bloğu yenilenir; EXTRA_DATA ve STOCKS şablondaki haliyle kalır.
    Mevcut index.html'deki şifre (SIFRE_HASH) korunur."""
    html = open(sablon_yolu, encoding="utf-8").read()
    js = json.dumps({k: v for k, v in data.items() if not k.startswith("_")}, ensure_ascii=False, separators=(",", ":"))
    html, n = re.subn(r"const DATA=\{.*?\};\n", lambda m: "const DATA=" + js + ";\n", html, count=1, flags=re.S)
    if n != 1:
        raise SystemExit("HATA: panel_sablon.html içinde veri alanı bulunamadı.")
    ek = data.get("_ek_satis")
    if ek:
        ej = json.dumps(ek, ensure_ascii=False, separators=(",", ":"))
        html, n2 = re.subn(r"const EXTRA_DATA=\{.*?\};\n", lambda m: "const EXTRA_DATA=" + ej + ";\n", html, count=1, flags=re.S)
        print(f"  ✓ Ek satış güncellendi ({ek['meta']['sourceRows']:,} satır, {len(ek['actual'])} kayıt)" if n2 else "  ⚠ Ek satış alanı şablonda bulunamadı")
    else:
        print("  ⚠ 'Ek Satış Ham' sayfası bulunamadı — ek satış verisi şablondaki haliyle kaldı")
    if os.path.exists(cikti_yolu):
        eski = open(cikti_yolu, encoding="utf-8").read()
        m = re.search(r"const SIFRE_HASH='([0-9a-f]{64})';", eski)
        if m:
            html = re.sub(r"const SIFRE_HASH='[0-9a-f]{64}';", "const SIFRE_HASH='" + m.group(1) + "';", html, count=1)
    print("  ✓ Yeni panel şablonu kullanıldı (panel_sablon.html)")
    return html


def main():
    girdi = sys.argv[1] if len(sys.argv) > 1 else "servis_rapor.xlsx"
    cikti = sys.argv[2] if len(sys.argv) > 2 else "dashboard.html"
    data = veriyi_topla(girdi)
    sablon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "panel_sablon.html")
    if os.path.exists(sablon):
        data["_ek_satis"] = ek_satis_uret(girdi, data)
        html = panel_uret(data, sablon, cikti)
    else:
        html = html_uret(data)
    with open(cikti, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n✓ Dashboard üretildi: {cikti}")
    print("  Tarayıcıda açmak için dosyaya çift tıklayın.")


if __name__ == "__main__":
    main()
