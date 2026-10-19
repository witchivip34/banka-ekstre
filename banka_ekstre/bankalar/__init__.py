"""Desteklenen bankaların modülleri ve kayıt defteri.

Yeni banka eklemek için: bu klasöre `AD`, `GORUNEN_AD`, `UZANTI`, `tani`, `parse`
içeren bir modül yazıp aşağıdaki `BANKALAR` listesine ekleyin.
"""
from __future__ import annotations

from types import ModuleType

from . import garanti, isbankasi, turkiyefinans, vakifbank, yapikredi, ziraat

BANKALAR: list[ModuleType] = [ziraat, yapikredi, garanti, isbankasi, turkiyefinans, vakifbank]
BANKA_MODULLERI: dict[str, ModuleType] = {m.AD: m for m in BANKALAR}


def modul(ad: str) -> ModuleType:
    """Banka kimliğinden modülü döndürür; bilinmiyorsa KeyError."""
    return BANKA_MODULLERI[ad]
