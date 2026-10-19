"""Tek bir ekstre dosyasını okuyup bankasını tespit eden üst düzey API."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .bankalar import modul
from .model import Islem, toplamlar
from .readers import read_rows
from .tespit import banka_adini_cozumle, banka_tani

UZANTILAR = (".xls", ".xlsx")


class TespitHatasi(ValueError):
    """Dosyanın bankası bulunamadı."""


@dataclass
class Ekstre:
    """Bir dosyadan okunan işlemler."""

    dosya: Path
    banka: str
    islemler: list[Islem] = field(default_factory=list)

    @property
    def giren_toplam(self) -> float:
        return toplamlar(self.islemler)[0]

    @property
    def cikan_toplam(self) -> float:
        return toplamlar(self.islemler)[1]


def ekstre_oku(yol: str | Path, banka: str | None = None) -> Ekstre:
    """Dosyayı okur; `banka` verilmezse içerikten/dosya adından tespit eder."""
    yol = Path(yol)
    satirlar = read_rows(yol)
    banka_id = banka_adini_cozumle(banka) if banka else banka_tani(yol, satirlar)
    if banka_id is None:
        raise TespitHatasi(f"{yol.name}: banka tespit edilemedi (--banka ile belirtin)")
    islemler = modul(banka_id).parse(satirlar, yol.name)
    return Ekstre(dosya=yol, banka=banka_id, islemler=islemler)


def dosyalari_bul(girdi: str | Path) -> list[Path]:
    """Girdi dosyaysa kendisini, klasörse içindeki .xls/.xlsx dosyalarını (sıralı) döndürür."""
    girdi = Path(girdi)
    if girdi.is_dir():
        return sorted(
            p for p in girdi.iterdir()
            if p.is_file() and p.suffix.lower() in UZANTILAR and not p.name.startswith("~$")
        )
    if girdi.is_file():
        return [girdi]
    raise FileNotFoundError(f"Girdi bulunamadı: {girdi}")
