"""Türkiye İş Bankası — internet şubesi "Hesap Özeti" (.xls).

Tarih sütunu `GG/AA/YYYY-SS:DD:ss` biçimindedir; saat kısmı atılır.
"""
from __future__ import annotations

from ..model import Islem
from ._ortak import SutunDuzeni, baslik_satiri_bul, standart_parse

AD = "isbankasi"
GORUNEN_AD = "Türkiye İş Bankası"
UZANTI = ".xls"

# Ham sütunlar: 0=Tarih/Saat | 1=Valör | 2=Kanal/Şube | 3=İşlem Tutarı* (±) | 4=Bakiye
#               5=Ek Hesap Bakiyesi | 6=İşlem | 7=İşlem Tipi | 8=Açıklama | 14=Referans
DUZEN = SutunDuzeni(tarih=0, referans=14, aciklama=8, tutar=3, bakiye=4)
IMZA = {0: "tarihsaat", 1: "valor", 3: "islemtutari", 8: "aciklama"}


def tani(rows: list[list]) -> bool:
    """Başlık satırı İş Bankası imzasıyla eşleşiyor mu?"""
    return baslik_satiri_bul(rows, IMZA) is not None


def parse(rows: list[list], kaynak_dosya: str) -> list[Islem]:
    """Ham satırları ortak `Islem` listesine çevirir."""
    return standart_parse(rows, kaynak_dosya, AD, DUZEN, IMZA)
