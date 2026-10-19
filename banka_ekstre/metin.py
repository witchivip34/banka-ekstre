"""Metin, tarih ve sayı yardımcıları."""
from __future__ import annotations

import re

_TR_SADE = str.maketrans({
    "İ": "i", "I": "i", "ı": "i", "Ç": "c", "ç": "c", "Ğ": "g", "ğ": "g",
    "Ö": "o", "ö": "o", "Ş": "s", "ş": "s", "Ü": "u", "ü": "u",
    "Â": "a", "â": "a", "Î": "i", "î": "i", "Û": "u", "û": "u",
})

_TARIH_RE = re.compile(r"^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{4})")


def sade(metin: object) -> str:
    """Türkçe karakterleri ASCII'ye indirir, küçük harfe çevirir, harf/rakam dışını atar.

    Örn. "İşlem Tutarı*" → "islemtutari".
    """
    s = str(metin).translate(_TR_SADE).lower()
    return re.sub(r"[^a-z0-9]", "", s)


def sade_parcalar(metin: object) -> list[str]:
    """Metni sadeleştirip harf/rakam dışı karakterlerden parçalara böler."""
    s = str(metin).translate(_TR_SADE).lower()
    return [p for p in re.split(r"[^a-z0-9]+", s) if p]


def metne(deger: object) -> str:
    """Hücre değerini metne çevirir (tam sayı float'lar ".0" olmadan)."""
    if deger is None:
        return ""
    if isinstance(deger, float) and deger.is_integer():
        return str(int(deger))
    return str(deger).strip()


def sayi_mi(deger: object) -> bool:
    """Değer gerçek bir sayı mı (bool hariç)?"""
    return isinstance(deger, (int, float)) and not isinstance(deger, bool)


def tarih_normalle(deger: object) -> str | None:
    """`10/01/2026`, `10.01.2026 18:30`, `10/01/2026-18:30:00` → `10.01.2026`.

    Başında tarih yoksa None döner.
    """
    m = _TARIH_RE.match(metne(deger))
    if not m:
        return None
    gun, ay, yil = m.groups()
    return f"{int(gun):02d}.{int(ay):02d}.{yil}"
