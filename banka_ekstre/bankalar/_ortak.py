"""Banka modüllerinin paylaştığı başlık tespiti ve satır ayrıştırma mantığı."""
from __future__ import annotations

from dataclasses import dataclass

from ..metin import metne, sade, sayi_mi, tarih_normalle
from ..model import Islem

# Başlık satırı ilk kaç satırda aranır
BASLIK_ARAMA_SINIRI = 60


@dataclass(frozen=True)
class SutunDuzeni:
    """Bir bankanın ham ekstresindeki sütun indeksleri (0 tabanlı)."""

    tarih: int
    referans: int | None
    aciklama: int
    tutar: int
    bakiye: int | None = None

    @property
    def gerekli_uzunluk(self) -> int:
        """Bir veri satırının en az kaç hücre içermesi gerektiği (tarih ve tutar için)."""
        return max(self.tarih, self.tutar) + 1


def baslik_satiri_bul(satirlar: list[list], imza: dict[int, str]) -> int | None:
    """`imza` ({sütun: sade başlık}) ile eşleşen ilk satırın indeksini döndürür."""
    for i, satir in enumerate(satirlar[:BASLIK_ARAMA_SINIRI]):
        if all(c < len(satir) and sade(satir[c]) == ad for c, ad in imza.items()):
            return i
    return None


def _hucre(satir: list, indeks: int | None) -> object:
    if indeks is None or indeks >= len(satir):
        return ""
    return satir[indeks]


def satirlari_ayristir(
    satirlar: list[list],
    duzen: SutunDuzeni,
    banka: str,
    kaynak_dosya: str,
    baslangic: int = 0,
) -> list[Islem]:
    """Tarih sütunu geçerli tarih, tutar sütunu sayı olan satırları `Islem`e çevirir."""
    sonuc: list[Islem] = []
    for i in range(baslangic, len(satirlar)):
        satir = satirlar[i]
        if len(satir) < duzen.gerekli_uzunluk:
            continue
        tarih = tarih_normalle(satir[duzen.tarih])
        tutar = satir[duzen.tutar]
        if tarih is None or not sayi_mi(tutar):
            continue
        bakiye = _hucre(satir, duzen.bakiye)
        sonuc.append(Islem(
            tarih=tarih,
            aciklama=metne(_hucre(satir, duzen.aciklama)),
            referans=metne(_hucre(satir, duzen.referans)),
            tutar=float(tutar),
            bakiye=float(bakiye) if sayi_mi(bakiye) else None,
            banka=banka,
            kaynak_dosya=kaynak_dosya,
            satir_no=i + 1,
        ))
    return sonuc


def standart_parse(
    satirlar: list[list],
    kaynak_dosya: str,
    banka: str,
    duzen: SutunDuzeni,
    imza: dict[int, str],
) -> list[Islem]:
    """Başlık bulunursa ondan sonraki, bulunmazsa tüm satırları ayrıştırır."""
    baslik = baslik_satiri_bul(satirlar, imza)
    baslangic = 0 if baslik is None else baslik + 1
    return satirlari_ayristir(satirlar, duzen, banka, kaynak_dosya, baslangic)
