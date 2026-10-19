"""Yapı Kredi — internet şubesi "Hesap Hareketleri" (.xls)."""
from __future__ import annotations

from ..model import Islem
from ._ortak import SutunDuzeni, baslik_satiri_bul, standart_parse

AD = "yapikredi"
GORUNEN_AD = "Yapı Kredi"
UZANTI = ".xls"

# Ham sütunlar: 0=Tarih | 1=Saat | 2=İşlem | 3=Kanal | 4=Referans No | 5=Açıklama
#               6=İşlem Tutarı (±) | 7=Bakiye | 8=Döviz Cinsi
DUZEN = SutunDuzeni(tarih=0, referans=4, aciklama=5, tutar=6, bakiye=7)
IMZA = {0: "tarih", 1: "saat", 4: "referansno", 5: "aciklama", 6: "islemtutari"}


def tani(rows: list[list]) -> bool:
    """Başlık satırı Yapı Kredi imzasıyla eşleşiyor mu?"""
    return baslik_satiri_bul(rows, IMZA) is not None


def parse(rows: list[list], kaynak_dosya: str) -> list[Islem]:
    """Ham satırları ortak `Islem` listesine çevirir."""
    return standart_parse(rows, kaynak_dosya, AD, DUZEN, IMZA)
