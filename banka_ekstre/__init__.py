"""banka-ekstre: Türk bankalarının ekstre dosyalarını ortak bir işlem listesine çevirir."""
from __future__ import annotations

from .ekstre import Ekstre, ekstre_oku
from .model import Islem

__version__ = "0.1.0"
__all__ = ["Ekstre", "Islem", "ekstre_oku", "__version__"]
