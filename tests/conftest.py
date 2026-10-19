"""Ortak test yardımcıları: sentetik örnek dosyalar ve beklenen işlemler."""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

KOK = Path(__file__).resolve().parents[1]
ORNEKLER = KOK / "ornekler"
sys.path.insert(0, str(KOK))

_spec = importlib.util.spec_from_file_location("ornek_uret", ORNEKLER / "uret.py")
uret = importlib.util.module_from_spec(_spec)
sys.modules["ornek_uret"] = uret
_spec.loader.exec_module(uret)

BANKALAR = list(uret.DOSYA_ADLARI)


@pytest.fixture(scope="session")
def ornek_klasoru(tmp_path_factory) -> Path:
    """Örnekleri geçici klasöre yeniden üretir (depodaki dosyalardan bağımsız)."""
    hedef = tmp_path_factory.mktemp("ornekler")
    uret.uret(hedef)
    return hedef


@pytest.fixture
def ornek_yolu(ornek_klasoru):
    def _yol(banka: str) -> Path:
        return ornek_klasoru / uret.DOSYA_ADLARI[banka]
    return _yol
