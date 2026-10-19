"""Ziraat Bankası — internet şubesi "Hesap Hareketleri" (.xlsx).

Ek/alt hesap ekstreleri de aynı düzende gelir ve bu modülle okunur.
"""
from __future__ import annotations

from ..model import Islem
from ._ortak import SutunDuzeni, baslik_satiri_bul, standart_parse

AD = "ziraat"
GORUNEN_AD = "Ziraat Bankası"
UZANTI = ".xlsx"

# Ham sütunlar: 0=Tarih | 1=Fiş No | 2=Açıklama | 3=İşlem Tutarı (±) | 4=Bakiye
DUZEN = SutunDuzeni(tarih=0, referans=1, aciklama=2, tutar=3, bakiye=4)
IMZA = {0: "tarih", 1: "fisno", 2: "aciklama", 3: "islemtutari", 4: "bakiye"}


def tani(rows: list[list]) -> bool:
    """Başlık satırı Ziraat imzasıyla eşleşiyor mu?"""
    return baslik_satiri_bul(rows, IMZA) is not None


def parse(rows: list[list], kaynak_dosya: str) -> list[Islem]:
    """Ham satırları ortak `Islem` listesine çevirir."""
    return standart_parse(rows, kaynak_dosya, AD, DUZEN, IMZA)
