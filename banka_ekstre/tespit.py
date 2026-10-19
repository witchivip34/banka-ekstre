"""Bir ekstre dosyasının hangi bankaya ait olduğunu bulur."""
from __future__ import annotations

from pathlib import Path

from .bankalar import BANKALAR, BANKA_MODULLERI
from .metin import sade, sade_parcalar
from .readers import read_rows

# Dosya adında aranan banka kelimeleri (sadeleştirilmiş, Türkçe karaktersiz).
# Sıra önemlidir: uzun/özgül kelimeler önce denenir.
AD_KELIMELERI: list[tuple[str, str]] = [
    ("turkiyefinans", "turkiyefinans"),
    ("tfinans", "turkiyefinans"),
    ("isbankasi", "isbankasi"),
    ("isbank", "isbankasi"),
    ("yapikredi", "yapikredi"),
    ("vakifbank", "vakifbank"),
    ("vakif", "vakifbank"),
    ("ziraat", "ziraat"),
    ("garanti", "garanti"),
]
# Yalnız ayrı bir parça olarak geçtiğinde kabul edilen kısa kelimeler ("t.iş-..." gibi)
KISA_PARCALAR: dict[str, str] = {"is": "isbankasi", "tis": "isbankasi"}


def icerikten_tani(rows: list[list]) -> str | None:
    """Başlık imzalarına bakarak banka kimliğini döndürür."""
    for m in BANKALAR:
        if m.tani(rows):
            return m.AD
    return None


def addan_tani(dosya_adi: str) -> str | None:
    """Dosya adındaki banka kelimesinden banka kimliğini döndürür."""
    return metinden_tani(Path(dosya_adi).stem)


def metinden_tani(kok: str) -> str | None:
    """Serbest metindeki ("t.finans", "İş Bankası" gibi) banka kelimesini bulur."""
    birlesik = sade(kok)
    for kelime, banka in AD_KELIMELERI:
        if kelime in birlesik:
            return banka
    for parca in sade_parcalar(kok):
        if parca in KISA_PARCALAR:
            return KISA_PARCALAR[parca]
    return None


def banka_adini_cozumle(ad: str) -> str:
    """`--banka` değerini (ör. "İş Bankası", "t.finans") banka kimliğine çevirir."""
    if sade(ad) in BANKA_MODULLERI:
        return sade(ad)
    bulunan = metinden_tani(ad)
    if bulunan is None:
        gecerli = ", ".join(BANKA_MODULLERI)
        raise ValueError(f"Bilinmeyen banka: {ad!r} (geçerli: {gecerli})")
    return bulunan


def banka_tani(path: str | Path, rows: list[list] | None = None) -> str | None:
    """Önce içerik imzası, bulunamazsa dosya adı; ikisi de yoksa None."""
    path = Path(path)
    if rows is None:
        rows = read_rows(path)
    return icerikten_tani(rows) or addan_tani(path.name)
