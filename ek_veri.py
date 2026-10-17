"""Haftalık gidişat ve ek satış / prim verisi üretimi (build_dashboard.py'ye eklenecek)."""
import datetime as _dt
from collections import defaultdict
import build_dashboard as b

_BAZ = _dt.date(1899, 12, 30)
FATURA_KAT = {"MEKANİK": "Mekanik", "KAPORTA BOYA": "Kaporta Boya", "GARANTİ": "Garanti", "BANKO": "Banko"}
FATURA_MARKA = {"bmw motorrad": "Motorrad", "mini": "Bmw", "bmw": "Bmw", "peugeot": "Peugeot", "opel": "Opel",
                "citroen": "Citroen", "fiat": "Fiat", "arj": "Arj", "honda": "Honda", "tesla": "Tesla", "jaecoo": "Jaecoo"}
PRIM_TUR = {"AKSAT-AKS": ("Aksesuar", "k", .11), "AKSAT-LAS": ("Lastik", "k", .11), "AKSAT-AKÜ": ("Akü", "k", .11),
            "AKSAT-ANTF": ("Antifiriz", "c", .03), "AKSAT-CAMSUYU": ("Cam Suyu", "c", .03), "AKSAT-MTKRM": ("Motor Koruma", "c", .03),
            "AKSAT-FAR": ("Far Ayarı", "c", .11), "AKSAT-ROTB": ("Rot Balans", "c", .11)}


def _n(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def _tarih(x):
    try:
        return _BAZ + _dt.timedelta(days=int(float(x)))
    except (TypeError, ValueError):
        return None


def _blok(gun):
    return min((gun - 1) // 7, 4)          # 0:1–7 1:8–14 2:15–21 3:22–28 4:29–ay sonu


def haftalik_uret(dosya, kesim_yil=None, kesim_ay=None, kesim_gun=None):
    """Fatura Listesi → [yıl, ay, blok, marka, kategori, adet, işçilik, yp] (marka dışı satırlar 'Diğer').
    Ayrıca bir önceki yılın kesim ayı için 'kesim gününe kadar' satırları blok=9 ile döner (adil kıyas)."""
    t = defaultdict(lambda: [0, 0.0, 0.0])
    for no, r in b._sayfa_akis(dosya, "Fatura Listesi", {"A", "E", "F", "L", "M"}):
        if no < 4:
            continue
        d = _tarih(r.get("F"))
        kat = FATURA_KAT.get((r.get("E") or "").strip())
        if not d or not kat:
            continue
        marka = FATURA_MARKA.get((r.get("A") or "").strip().lower(), "Diğer")
        a = t[(d.year, d.month, _blok(d.day), marka, kat)]
        a[0] += 1; a[1] += _n(r.get("L")); a[2] += _n(r.get("M"))
        if kesim_ay and d.year == kesim_yil - 1 and d.month == kesim_ay and d.day <= kesim_gun:
            a = t[(d.year, d.month, 9, marka, kat)]
            a[0] += 1; a[1] += _n(r.get("L")); a[2] += _n(r.get("M"))
    return [[y, m, bl, mk, k, v[0], round(v[1], 2), round(v[2], 2)] for (y, m, bl, mk, k), v in sorted(t.items())]


def _danisman(ad):
    ad = " ".join((ad or "").split())
    if not ad or b._isim_anahtar(ad) in ("DANISMANI BELIRLE", "DANISMANI BELIRLENMEMIS"):
        return "Danışmanı Belirlenmemiş"
    return b._tr_baslik(ad)


def prim_uret(dosya):
    """ek_satis.xlsx (Prim Yp + Prim İşçilik) → [yıl, ay, marka, danışman, tür, ciro, karlılık, prim]."""
    t = defaultdict(lambda: [0.0, 0.0, 0.0])
    for sayfa in ("Prim Yp", "Prim İşçilik"):
        for no, r in b._sayfa_akis(dosya, sayfa, {"C", "G", "H", "I", "L", "V"}):
            if no < 4:
                continue
            tur = PRIM_TUR.get((r.get("H") or "").strip())
            d = _tarih(r.get("C"))
            if not tur or not d:
                continue
            ciro, kar = _n(r.get("G")), _n(r.get("V"))
            prim = (kar if tur[1] == "k" else ciro) * tur[2]
            marka = FATURA_MARKA.get((r.get("L") or "").strip().lower(), (r.get("L") or "").strip())
            a = t[(d.year, d.month, marka, _danisman(r.get("I")), tur[0])]
            a[0] += ciro; a[1] += kar; a[2] += prim
    return [[y, m, mk, dn, tr, round(v[0], 2), round(v[1], 2), round(v[2], 2)] for (y, m, mk, dn, tr), v in sorted(t.items())]


def ek_ham_uret(dosya, data):
    """servis_rapor 'Ek Satış Ham' → [yıl, ay, marka, danışman, tür, tutar] (isimler normalize)."""
    ek = b.ek_satis_uret(dosya, data)
    if not ek:
        return []
    return [[r["y"], r["m"], r["brand"], r["consultant"], r["type"], r["value"]] for r in ek["actual"]]


def ek_butce_uret(dosya):
    """ek_satis.xlsx 'Ek Satış Bütçe' → [yıl, ay, marka, tür, bütçe ciro]."""
    t = defaultdict(float)
    for no, r in b._sayfa_akis(dosya, "Ek Satış Bütçe", {"C", "E", "F", "I", "P"}):
        if no < 2:
            continue
        try:
            y, m = int(float(r["E"])), int(float(r["F"]))
        except (KeyError, TypeError, ValueError):
            continue
        marka = FATURA_MARKA.get((r.get("C") or "").strip().lower(), (r.get("C") or "").strip())
        t[(y, m, marka, (r.get("I") or "").strip())] += _n(r.get("P"))
    return [[y, m, mk, tr, round(v, 2)] for (y, m, mk, tr), v in sorted(t.items()) if v]


# ---------------- Ek Satış Q4 hedef dosyası (Ek_Satış_Son_Rapor_Q4.xlsx) ----------------
Q4_BLOK = {"PEUGEOT": ["Peugeot"], "OPEL": ["Opel"], "FİAT ARJ": ["Fiat", "Arj"], "BMW": ["Bmw"],
           "CİTROEN": ["Citroen"], "HONDA": ["Honda"], "JAECOO": ["Jaecoo"]}
Q4_BASLIK = {"PEUGEOT SSH": "PEUGEOT", "OPEL SSH": "OPEL", "FİAT ARJ SSH": "FİAT ARJ", "BMW SSH": "BMW",
             "CİTROENSSH": "CİTROEN", "CİTROEN SSH": "CİTROEN", "HONDA SSH": "HONDA", "JAECOO SSH": "JAECOO"}
Q4_SAYFA = {"PEUGEOT": "PEUGEOT", "CİTROEN": "CİTROEN", "OPEL": "OPEL", "HONDA": "HONDA", "FİAT ARJ": "FİAT+ARJ",
            "JAECOO": "JAECOO", "BMW": "BMW"}
Q4_KOD = {"Lastik": "AKSAT-LAS", "Cam Suyu": "AKSAT-CAMSUYU", "Aksesuar": "AKSAT-AKS",
          "Far Ayarı": "AKSAT-FAR", "Rot Balans": "AKSAT-ROTB"}
Q4_KRITER = ["Car Care", "Lastik", "Cam Suyu", "Akü", "Aksesuar", "Banko Satış", "Far Ayarı", "Rot Balans"]
HARIC_CARI = "120.101.4780513615"          # İnciroğlu Otomotiv iç cari — hariç
_MARKA_BLOK = {m: bl for bl, ms in Q4_BLOK.items() for m in ms}


def _q4_kriter(bolum, kod):
    """Excel 'Konsolide Hedef' formülleriyle aynı kural."""
    if kod == "CAR CARE":
        return "Car Care"                                   # bölümden bağımsız
    if bolum == "BANKO" and kod == "":
        return "Banko Satış"                                # özel kodu boş + BANKO
    if bolum in ("MEKANİK", "BANKO"):
        for k, c in Q4_KOD.items():
            if kod == c:
                return k
    return None


def q4_uret(dosya):
    # 1) hedefler: blok × kriter × ay (adet, plan ciro, bütçe)
    hedef, blok = [], None
    for no, r in b._sayfa_akis(dosya, "Konsolide Hedef", {"A", "B", "C", "D", "F", "G", "H", "I", "J", "K"}):
        a = (r.get("A") or "").strip()
        if a.upper() in Q4_BASLIK:
            blok = Q4_BASLIK[a.upper()]; continue
        if blok and a in Q4_KRITER:
            for m, (ca, cp, cb) in zip((10, 11, 12), (("B", "C", "D"), ("F", "G", "H"), ("I", "J", "K"))):
                hedef.append([blok, a, m, _n(r.get(ca)), _n(r.get(cp)), _n(r.get(cb))])
        if a == "Toplam":
            blok = None
    # 2) fiili: Fatura Detay Liste (Excel kuralları), 2025 ve 2026
    aylik, hafta, dan = defaultdict(lambda: [0.0, 0.0]), defaultdict(lambda: [0.0, 0.0]), defaultdict(lambda: [0.0, 0.0])
    son = None
    for no, r in b._sayfa_akis(dosya, "Fatura Detay Liste", {"C", "E", "F", "G", "H", "I", "K", "L"}):
        if no < 4:
            continue
        d = _tarih(r.get("C"))
        if not d or d.year < 2025:
            continue
        son = d if (son is None or d > son) else son        # kesim: sayfadaki son fatura tarihi (tüm satırlar)
        if (r.get("E") or "").startswith(HARIC_CARI):
            continue
        bl = _MARKA_BLOK.get((r.get("K") or "").strip())
        k = _q4_kriter((r.get("L") or "").strip(), (r.get("G") or "").strip())
        if not bl or not k:
            continue
        v, q = _n(r.get("F")), _n(r.get("I"))
        a = aylik[(d.year, d.month, bl, k)]; a[0] += v; a[1] += q
        h = hafta[(d.year, d.month, _blok(d.day), bl, k)]; h[0] += v; h[1] += q
        h2 = hafta[(d.year, d.month, 50 + d.day, bl, k)]; h2[0] += v; h2[1] += q   # günlük (aynı gün kıyası için)
        dn = dan[(d.year, d.month, bl, _danisman(r.get("H")), k)]
        dn[0] += v; dn[1] += q
    # 3) güncel stok (marka sayfaları)
    stok = []
    for bl, sayfa in Q4_SAYFA.items():
        for no, r in b._sayfa_akis(dosya, sayfa, {"E", "F", "G"}):
            k = (r.get("E") or "").strip()
            if no >= 3 and k and k != "Kriter":
                stok.append([bl, k, _n(r.get("F")), _n(r.get("G"))])
    return {"son": son.strftime("%d.%m.%Y") if son else "",
            "hedef": hedef,
            "aylik": [[y, m, bl, k, round(v[0], 2), v[1]] for (y, m, bl, k), v in sorted(aylik.items())],
            "hafta": [[y, m, w, bl, k, round(v[0], 2), v[1]] for (y, m, w, bl, k), v in sorted(hafta.items()) if w < 50 or (m >= 10)],
            "dan": [[y, m, bl, dn, k, round(v[0], 2), v[1]] for (y, m, bl, dn, k), v in sorted(dan.items())],
            "stok": stok, "kriter": Q4_KRITER, "blok": {k: v for k, v in Q4_BLOK.items()}}


# ================================================================== #
#  Ek satış sayfası — ek-satis-dashboard/build_ek_satis.py ile aynı okuma
#  Kaynak: ek_satis.xlsx "Konsolide Hedef" (marka blokları Plan / Fiili)
# ================================================================== #
import datetime as dt
import sys
import unicodedata
SHEET = "Konsolide Hedef"
LEFT_COLS = 17  # sol taraftaki marka blokları (S sütunundan sonrası konsolide/parametre)

# Marka sırası ve gösterim adları (blok başlığı -> ad)
BRAND_MAP = {
    "PEUGEOT": "Peugeot",
    "CITROEN": "Citroën",
    "OPEL": "Opel",
    "FIATARJ": "Fiat, Alfa Romeo & Jeep",
    "HONDA": "Honda",
    "BMW": "BMW & MINI",
    "JAECOO": "Omoda & Jaecoo",
}
MONTHS = {"OCAK": 1, "SUBAT": 2, "MART": 3, "NISAN": 4, "MAYIS": 5, "HAZIRAN": 6,
          "TEMMUZ": 7, "AGUSTOS": 8, "EYLUL": 9, "EKIM": 10, "KASIM": 11, "ARALIK": 12}

# Çalışma günü: Cumartesi dahil, Pazar ve resmi tatiller hariç (düzenlenebilir)
HOLIDAYS = {dt.date(2026, *md) for md in [
    (1, 1), (3, 20), (3, 21), (3, 22), (4, 23), (5, 1), (5, 19),
    (5, 27), (5, 28), (5, 29), (5, 30), (7, 15), (8, 30), (10, 29)]}


def norm(v):
    """Türkçe karakter / boşluk / büyük-küçük harf farklarını yok say."""
    if v is None:
        return ""
    s = unicodedata.normalize("NFKD", str(v))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("ı", "I").replace("İ", "I")
    return "".join(s.upper().split())


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def workday(d):
    return d.weekday() != 6 and d not in HOLIDAYS


def month_bounds(year, m):
    start = dt.date(year, m, 1)
    nxt = dt.date(year + (m == 12), m % 12 + 1, 1)
    return start, nxt - dt.timedelta(days=1)


def count_workdays(a, b):
    n, d = 0, a
    while d <= b:
        n += workday(d)
        d += dt.timedelta(days=1)
    return n


def ek_panel_uret(path, today=None):
    """ek_satis.xlsx 'Konsolide Hedef' → ek satış sayfası verisi (ek-satis-dashboard ile aynı hesap)."""
    import openpyxl
    today = today or dt.date.today()
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if SHEET not in wb.sheetnames:
        raise SystemExit(f"HATA: '{SHEET}' sayfası bulunamadı. Sayfalar: {wb.sheetnames}")
    rows = [list(r) for r in wb[SHEET].iter_rows(values_only=True)]
    wb.close()

    # Dönem yılı: 'Başlangıç tarihi' parametresinden
    year = today.year
    for r in rows:
        for i, c in enumerate(r):
            if norm(c) == "BASLANGICTARIHI":
                for nxt in r[i + 1:]:
                    if isinstance(nxt, (dt.datetime, dt.date)):
                        year = nxt.year
                        break

    brands, month_names, warns = {}, None, []
    for ri, r in enumerate(rows):
        t = norm(r[0] if r else None)
        if not t.endswith("SSH") or ri + 1 >= len(rows):
            continue
        key = t[:-3]
        title, sub = r[:LEFT_COLS], rows[ri + 1][:LEFT_COLS]
        mcols = [(i, MONTHS[norm(c)]) for i, c in enumerate(title) if norm(c) in MONTHS]
        if not mcols:
            continue
        end_all = next((i for i, c in enumerate(title) if norm(c) == "TOPLAM"), len(title))
        spans = []
        for j, (c0, mno) in enumerate(mcols):
            c1 = mcols[j + 1][0] if j + 1 < len(mcols) else end_all
            lab = {norm(x): c0 + k for k, x in enumerate(sub[c0:c1]) if x}
            spans.append((mno, title[c0], lab))
        if month_names is None:
            month_names = [(m, str(n)) for m, n, _ in spans]
        items, total_plan = [], {}
        for rr in rows[ri + 2:]:
            k = rr[0]
            if k is None:
                break
            if norm(k) == "TOPLAM":
                for mno, _, lab in spans:
                    if "PLAN" in lab:
                        total_plan[mno] = num(rr[lab["PLAN"]]) or 0
                break
            ms = []
            for mno, _, lab in spans:
                p = (num(rr[lab["PLAN"]]) if "PLAN" in lab else 0) or 0
                f = num(rr[lab["FIILI"]]) if "FIILI" in lab else None
                ms.append({"p": p, "f": f})
            items.append({"k": str(k).strip(), "m": ms})
        for i, (mno, _, _) in enumerate(spans):
            s = sum(it["m"][i]["p"] for it in items)
            if mno in total_plan and abs(s - total_plan[mno]) > 1:
                warns.append(f"{key} {mno}. ay: kalem toplamı {s:,.0f} ≠ sayfadaki Toplam {total_plan[mno]:,.0f}")
        brands[key] = {"key": key, "ad": BRAND_MAP.get(key, key.title()), "items": items}

    if not brands or not month_names:
        raise SystemExit("HATA: Marka blokları okunamadı (başlıkların 'xxx SSH' formatında olduğundan emin olun).")

    order = [k for k in BRAND_MAP if k in brands] + [k for k in brands if k not in BRAND_MAP]
    months = []
    for mno, ad in month_names:
        s, e = month_bounds(year, mno)
        durum = "kapandi" if today > e else ("gelecek" if today < s else "devam")
        tot = count_workdays(s, e)
        gec = tot if durum == "kapandi" else (0 if durum == "gelecek" else count_workdays(s, today))
        months.append({"ad": ad, "no": mno, "durum": durum, "toplamGun": tot, "gecenGun": gec,
                       "oran": gec / tot if tot else 0})
    q = (month_names[0][0] - 1) // 3 + 1
    donem = f"{year} Q{q}" if len(month_names) == 3 else f"{year} {month_names[0][1]}–{month_names[-1][1]}"
    data = {"meta": {"donem": donem, "veriTarihi": today.strftime("%d.%m.%Y"),
                     "uretim": dt.datetime.now().strftime("%d.%m.%Y %H:%M")},
            "months": months, "brands": [brands[k] for k in order]}
    for w in warns:
        print("  ⚠ " + w)
    return data


