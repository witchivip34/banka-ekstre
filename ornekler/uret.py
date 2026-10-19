"""Her banka için SENTETİK ham ekstre dosyası üretir.

Dosyalar gerçek internet şubesi çıktılarının başlık/sütun düzenini taklit eder; içerikteki
hesap bilgileri, açıklamalar ve tutarlar tamamen uydurmadır. Üretim deterministiktir
(sabit tohum), yani aynı komut her seferinde aynı dosyaları üretir.

Kullanım:
    python ornekler/uret.py            # dosyaları bu klasöre yazar
    python ornekler/uret.py hedef/     # başka klasöre yazar
"""
from __future__ import annotations

import random
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

import xlwt
from openpyxl import Workbook

ACIKLAMALAR_GIREN = [
    "HAVALE GELEN", "POS TAHSİLAT", "EFT GELEN - ÖRNEK TİCARET", "FAST GELEN", "KREDİ KARTI SATIŞ TAHSİLATI",
]
ACIKLAMALAR_CIKAN = [
    "MARKET ALIŞVERİŞİ", "KİRA ÖDEMESİ", "ELEKTRİK FATURASI", "SU FATURASI", "İNTERNET FATURASI",
    "TEDARİKÇİ ÖDEMESİ", "HESAP İŞLETİM ÜCRETİ", "AKARYAKIT ALIMI", "VERGİ ÖDEMESİ",
]
ADETLER = {"ziraat": 10, "yapikredi": 12, "garanti": 9, "isbankasi": 11, "turkiyefinans": 8, "vakifbank": 10}
DOSYA_ADLARI = {
    "ziraat": "ziraat-ornek.xlsx",
    "yapikredi": "yapikredi-ornek.xls",
    "garanti": "garanti-ornek.xls",
    "isbankasi": "isbankasi-ornek.xls",
    "turkiyefinans": "turkiyefinans-ornek.xlsx",
    "vakifbank": "vakifbank-ornek.xlsx",
}


@dataclass
class OrnekIslem:
    zaman: datetime
    aciklama: str
    referans: str
    tutar: float
    bakiye: float


def ornek_islemler(banka: str) -> list[OrnekIslem]:
    """Banka için deterministik uydurma işlem listesi (ilk iki işlemden biri giren, biri çıkan)."""
    rnd = random.Random(f"banka-ekstre-{banka}")
    zaman = datetime(2026, 1 + rnd.randrange(6), 1, 9, 0)
    bakiye = round(rnd.uniform(5_000, 50_000), 2)
    sonuc = []
    for k in range(ADETLER[banka]):
        giren = k == 0 or (k != 1 and rnd.random() < 0.4)
        if giren:
            tutar = round(rnd.uniform(250, 25_000), 2)
            aciklama = rnd.choice(ACIKLAMALAR_GIREN)
        else:
            tutar = -round(rnd.uniform(15, 8_000), 2)
            aciklama = rnd.choice(ACIKLAMALAR_CIKAN)
        bakiye = round(bakiye + tutar, 2)
        zaman += timedelta(hours=rnd.randrange(3, 60), minutes=rnd.randrange(60))
        sonuc.append(OrnekIslem(zaman, aciklama, f"{rnd.randrange(10**9, 10**10)}", tutar, bakiye))
    return sonuc


# ─────────────────────────────── yazıcılar ───────────────────────────────

def _xls_yaz(yol: Path, satirlar: list[list]) -> None:
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Sayfa1")
    for r, satir in enumerate(satirlar):
        for c, deger in enumerate(satir):
            if deger != "":
                ws.write(r, c, deger)
    wb.save(str(yol))


def _xlsx_yaz(yol: Path, satirlar: list[list]) -> None:
    wb = Workbook()
    ws = wb.active
    for satir in satirlar:
        ws.append([None if d == "" else d for d in satir])
    wb.save(yol)


def _dolgu(satir: list, uzunluk: int) -> list:
    return satir + [""] * (uzunluk - len(satir))


# ─────────────────────────── banka şablonları ────────────────────────────

def ziraat(islemler: list[OrnekIslem]) -> list[list]:
    s = [["Sayın ÖRNEK MÜŞTERİ"], ["01.01.2026 - 31.01.2026 tarihleri arası hesap hareketleri"],
         [], ["Hesap Numarası", "0000-00000000-0000"], ["Hesap Türü", "Vadesiz TL"], [],
         ["Hesap Hareketleri"], [], [], [], [],
         ["Tarih", "Fiş No", "Açıklama", "İşlem Tutarı", "Bakiye"]]
    for i in islemler:
        s.append([i.zaman.strftime("%d.%m.%Y"), f"{int(i.referans) % 100000:05d}", i.aciklama, i.tutar, i.bakiye])
    return s


def yapikredi(islemler: list[OrnekIslem]) -> list[list]:
    s = [["Hesap Hareketleri"], ["Müşteri Adı Soyadı", "ÖRNEK MÜŞTERİ"], ["Müşteri Numarası", "00000000"],
         ["Şube", "ÖRNEK ŞUBE"], ["Hesap Adı", "Vadesiz TL"], ["Hesap No", "0000000-0"],
         ["Kullanılabilir Bakiye", "0,00 TL"], ["Tarih Aralığı", "01/01/2026 - 31/01/2026"], [], [],
         ["Tarih", "Saat", "İşlem", "Kanal", "Referans No", "Açıklama", "İşlem Tutarı", "Bakiye", "Döviz Cinsi"]]
    for i in islemler:
        islem = "Gelen EFT" if i.tutar > 0 else "Ödeme"
        s.append([i.zaman.strftime("%d/%m/%Y"), i.zaman.strftime("%H:%M"), islem, "İnternet",
                  i.referans, i.aciklama, i.tutar, i.bakiye, "TL"])
    return s


def garanti(islemler: list[OrnekIslem]) -> list[list]:
    etiketler = ["Para Transferi", "Diğer", "Fatura", "Alışveriş"]
    s = [["Şirket Ünvanı", "ÖRNEK MÜŞTERİ"], ["Hesap", "0000-0000000"], ["Şube", "ÖRNEK ŞUBE"],
         ["Başlangıç", "01/01/2026"], ["Bitiş", "31/01/2026"], ["İndirme Tarihi", "01/02/2026"],
         ["Bakiye", "0,00 TL"], ["Kullanılabilir Bakiye", "0,00 TL"], [], [], [], [], [], [],
         ["Tarih", "Açıklama", "Etiket", "Tutar", "Bakiye", "Dekont No"]]
    for k, i in enumerate(islemler):
        s.append([i.zaman.strftime("%d/%m/%Y"), i.aciklama, etiketler[k % len(etiketler)],
                  i.tutar, i.bakiye, i.referans])
    return s


def isbankasi(islemler: list[OrnekIslem]) -> list[list]:
    n = 15
    bas = [[""] * n for _ in range(15)]
    bas[0][0] = "ÖRNEK VERGİ DAİRESİ"
    bas[5][6] = "HESAP ÖZETİ"
    bas[7][2] = "Sayın ÖRNEK MÜŞTERİ"
    bas[8][2], bas[8][4] = "Müşteri numarası:", "000000000"
    bas[9][2], bas[9][6], bas[9][9] = "Vergi No:", "İşlem Yeri:", "İnternet"
    bas[10][2], bas[10][4] = "Hesap No:", "0000 0000000"
    bas[11][2] = "Hesap Şube Kodu/Adı:"
    bas[12][2], bas[12][4] = "Ürün-Döviz Tipi:", "Vadesiz Hesap"
    bas[13][2] = "Mevcut Bakiye:"
    bas[14][2] = "Net Bakiye:"
    baslik = _dolgu(["Tarih/Saat", "Valör", "Kanal/Şube", "İşlem Tutarı*", "Bakiye", "Ek Hesap\nBakiyesi",
                     "İşlem", "İşlem Tipi", "Açıklama"], 14) + ["Referans"]
    s = bas + [baslik]
    for i in islemler:
        tip = ("FA", "FAST") if i.tutar > 0 else ("OD", "ÖDEME")
        s.append(_dolgu([i.zaman.strftime("%d/%m/%Y-%H:%M:%S"), i.zaman.strftime("%d/%m/%Y"), "İnternet",
                         i.tutar, i.bakiye, 0.0, tip[0], tip[1], i.aciklama], 14) + [i.referans])
    s += [[], ["(*) Çekilen tutar '-' ile gösterilmiştir."], ["Bu belge örnek amaçlıdır."]]
    return s


def turkiyefinans(islemler: list[OrnekIslem]) -> list[list]:
    s = [["Müşteri", "ÖRNEK MÜŞTERİ"], ["Hesap", "000000-0"], ["Döviz Cinsi", "TL"],
         ["Başlangıç", "01.01.2026"], ["Bitiş", "31.01.2026"], [],
         ["İşlem Tarihi", "İşlem Referansı", "Açıklama", "Tutar", "Hesap Bakiyesi", "Döviz Cinsi"]]
    for i in islemler:
        s.append([i.zaman.strftime("%d.%m.%Y %H:%M"), i.referans, i.aciklama, i.tutar, i.bakiye, "TL"])
    return s


def vakifbank(islemler: list[OrnekIslem]) -> list[list]:
    s = [_dolgu(["HESAP SAHİBİ", "ÖRNEK MÜŞTERİ", "", "", "", "", "TARİH VE SAAT", "01.02.2026 10:00"], 8),
         _dolgu(["MÜŞTERİ NO", "00000000", "", "", "", "", "BAKİYE", "0,00"], 8),
         _dolgu(["HESAP", "0000000000", "", "", "", "", "EK HESAP LİMİTİ", "0,00"], 8),
         _dolgu(["HESAP TÜRÜ", "VADESİZ TL", "", "", "", "", "TARİH ARALIĞI", "01.01.2026 - 31.01.2026"], 8),
         _dolgu(["ŞUBE", "ÖRNEK ŞUBE", "", "", "", "", "RUMUZ", ""], 8),
         ["HESAP NO", "FİŞ NO", "HAREKET TARIH", "İŞLEM TARİHİ", "KART NO", "İŞLEM", "TUTAR", "BAKİYE", "KANAL",
          "İŞLEM NO", "REFERANS", "HAVALE", "REF NO", "TCKN", "VKN", "B/A", "AÇIKLAMA"]]
    for k, i in enumerate(islemler):
        zaman = i.zaman.strftime("%d.%m.%Y %H:%M")
        s.append(["00000000000000000", float(1000 + k), zaman, zaman, "", "Hesap Hareketi", i.tutar, i.bakiye,
                  "İnternet", i.referans, "", "", "", "", "", "A" if i.tutar > 0 else "B", i.aciklama])
    s += [[""], ["Ticaret Unvanı: ÖRNEK BANKA A.Ş."], [""], ["Bu belge örnek amaçlıdır."]]
    return s


SABLONLAR = {
    "ziraat": ziraat, "yapikredi": yapikredi, "garanti": garanti,
    "isbankasi": isbankasi, "turkiyefinans": turkiyefinans, "vakifbank": vakifbank,
}


def uret(hedef: Path) -> list[Path]:
    """Tüm bankalar için örnek dosyaları `hedef` klasörüne yazar."""
    hedef.mkdir(parents=True, exist_ok=True)
    yollar = []
    for banka, sablon in SABLONLAR.items():
        yol = hedef / DOSYA_ADLARI[banka]
        satirlar = sablon(ornek_islemler(banka))
        (_xls_yaz if yol.suffix == ".xls" else _xlsx_yaz)(yol, satirlar)
        yollar.append(yol)
    return yollar


if __name__ == "__main__":
    klasor = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent
    for p in uret(klasor):
        print(p)
