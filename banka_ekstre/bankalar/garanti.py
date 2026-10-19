"""Garanti BBVA — internet şubesi hesap hareketleri (.xls).

Ek hesap (ör. e-ticaret/POS tahsilat hesabı) ekstreleri de aynı düzende gelir.
"""
from __future__ import annotations

from ..model import Islem
from ._ortak import SutunDuzeni, baslik_satiri_bul, standart_parse

AD = "garanti"
GORUNEN_AD = "Garanti BBVA"
UZANTI = ".xls"

# Ham sütunlar: 0=Tarih | 1=Açıklama | 2=Etiket | 3=Tutar (±) | 4=Bakiye | 5=Dekont No
DUZEN = SutunDuzeni(tarih=0, referans=5, aciklama=1, tutar=3, bakiye=4)
IMZA = {0: "tarih", 1: "aciklama", 2: "etiket", 3: "tutar", 5: "dekontno"}


def tani(rows: list[list]) -> bool:
    """Başlık satırı Garanti imzasıyla eşleşiyor mu?"""
    return baslik_satiri_bul(rows, IMZA) is not None


def parse(rows: list[list], kaynak_dosya: str) -> list[Islem]:
    """Ham satırları ortak `Islem` listesine çevirir."""
    return standart_parse(rows, kaynak_dosya, AD, DUZEN, IMZA)
