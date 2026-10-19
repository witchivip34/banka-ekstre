"""VakıfBank — internet şubesi hesap hareketleri (.xlsx).

Başlık satırından önce birkaç satırlık hesap özeti bloğu bulunur; tarih sütununda saat
de vardır (`GG.AA.YYYY SS:DD`).
"""
from __future__ import annotations

from ..model import Islem
from ._ortak import SutunDuzeni, baslik_satiri_bul, standart_parse

AD = "vakifbank"
GORUNEN_AD = "VakıfBank"
UZANTI = ".xlsx"

# Ham sütunlar: 0=Hesap No | 1=Fiş No | 2=Hareket Tarih | 3=İşlem Tarihi | 4=Kart No
#               5=İşlem | 6=Tutar (±) | 7=Bakiye | 8=Kanal | 9=İşlem No | 10=Referans
#               11=Havale | 12=Ref No | 13-14=kimlik/vergi no | 15=B/A | 16=Açıklama
DUZEN = SutunDuzeni(tarih=2, referans=9, aciklama=16, tutar=6, bakiye=7)
IMZA = {2: "harekettarih", 6: "tutar", 9: "islemno", 16: "aciklama"}


def tani(rows: list[list]) -> bool:
    """Başlık satırı VakıfBank imzasıyla eşleşiyor mu?"""
    return baslik_satiri_bul(rows, IMZA) is not None


def parse(rows: list[list], kaynak_dosya: str) -> list[Islem]:
    """Ham satırları ortak `Islem` listesine çevirir."""
    return standart_parse(rows, kaynak_dosya, AD, DUZEN, IMZA)
