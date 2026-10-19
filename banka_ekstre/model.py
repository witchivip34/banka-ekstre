"""Bankadan bağımsız ortak işlem modeli."""
from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import date


@dataclass
class Islem:
    """Ekstredeki tek bir hesap hareketi.

    `tutar` işaretlidir: pozitif = hesaba giren, negatif = hesaptan çıkan.
    """

    tarih: str  # DD.MM.YYYY
    aciklama: str
    referans: str
    tutar: float
    bakiye: float | None
    banka: str
    kaynak_dosya: str
    satir_no: int  # ham dosyadaki 1 tabanlı satır numarası

    @property
    def giren(self) -> float:
        """Giren tutar (pozitif ya da 0)."""
        return self.tutar if self.tutar > 0 else 0.0

    @property
    def cikan(self) -> float:
        """Çıkan tutarın mutlak değeri (pozitif ya da 0)."""
        return -self.tutar if self.tutar < 0 else 0.0

    @property
    def tarih_nesnesi(self) -> date:
        """`tarih` alanını `datetime.date` olarak döndürür."""
        gun, ay, yil = self.tarih.split(".")
        return date(int(yil), int(ay), int(gun))

    def sozluk(self) -> dict:
        """Alanları sözlük olarak döndürür."""
        return asdict(self)


def toplamlar(islemler: list[Islem]) -> tuple[float, float]:
    """(giren toplamı, çıkan toplamı) — ikisi de pozitif."""
    giren = sum(i.giren for i in islemler)
    cikan = sum(i.cikan for i in islemler)
    return round(giren, 2), round(cikan, 2)


def tarih_araligi(islemler: list[Islem]) -> tuple[str, str] | None:
    """En erken ve en geç işlem tarihini DD.MM.YYYY olarak döndürür."""
    if not islemler:
        return None
    tarihler = sorted(islemler, key=lambda i: i.tarih_nesnesi)
    return tarihler[0].tarih, tarihler[-1].tarih
