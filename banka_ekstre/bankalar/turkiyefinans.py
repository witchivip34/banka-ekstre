"""Türkiye Finans — internet şubesi hesap hareketleri (.xlsx).

Tarih sütununda saat de bulunur (`GG.AA.YYYY SS:DD`); saat kısmı atılır.
"""
from __future__ import annotations

from ..model import Islem
from ._ortak import SutunDuzeni, baslik_satiri_bul, standart_parse

AD = "turkiyefinans"
GORUNEN_AD = "Türkiye Finans"
UZANTI = ".xlsx"

# Ham sütunlar: 0=İşlem Tarihi | 1=İşlem Referansı | 2=Açıklama | 3=Tutar (±)
#               4=Hesap Bakiyesi | 5=Döviz Cinsi
DUZEN = SutunDuzeni(tarih=0, referans=1, aciklama=2, tutar=3, bakiye=4)
IMZA = {0: "islemtarihi", 1: "islemreferansi", 2: "aciklama", 3: "tutar"}


def tani(rows: list[list]) -> bool:
    """Başlık satırı Türkiye Finans imzasıyla eşleşiyor mu?"""
    return baslik_satiri_bul(rows, IMZA) is not None


def parse(rows: list[list], kaynak_dosya: str) -> list[Islem]:
    """Ham satırları ortak `Islem` listesine çevirir."""
    return standart_parse(rows, kaynak_dosya, AD, DUZEN, IMZA)
