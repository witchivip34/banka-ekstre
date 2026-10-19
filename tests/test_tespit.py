"""Banka tespiti: içerik imzası, dosya adı ve elle zorlama."""
from __future__ import annotations

import shutil

import pytest

from banka_ekstre import ekstre_oku
from banka_ekstre.ekstre import TespitHatasi
from banka_ekstre.readers import read_rows
from banka_ekstre.tespit import addan_tani, banka_adini_cozumle, banka_tani, icerikten_tani
from conftest import BANKALAR


@pytest.mark.parametrize("banka", BANKALAR)
def test_icerikten_tespit_dosya_adi_yaniltici_olsa_da(banka, ornek_yolu, tmp_path):
    # Dosya adı başka bir bankayı işaret etse bile içerik imzası önceliklidir.
    kaynak = ornek_yolu(banka)
    yaniltici = "garanti" if banka != "garanti" else "ziraat"
    hedef = tmp_path / f"{yaniltici}-ekstre{kaynak.suffix}"
    shutil.copy(kaynak, hedef)
    assert icerikten_tani(read_rows(hedef)) == banka
    assert banka_tani(hedef) == banka


@pytest.mark.parametrize("ad, beklenen", [
    ("ziraat-mart1.xlsx", "ziraat"),
    ("ZİRAAT_ek_hesap.xlsx", "ziraat"),
    ("yapıkredi-ocak.xls", "yapikredi"),
    ("Yapi Kredi 2026.xls", "yapikredi"),
    ("garanti-eticaret-2.xls", "garanti"),
    ("t.iş-ocak1.xls", "isbankasi"),
    ("t.iş bankası-mart.xls", "isbankasi"),
    ("İŞ BANKASI.xls", "isbankasi"),
    ("isbankasi.xls", "isbankasi"),
    ("t.finans-ocak.xlsx", "turkiyefinans"),
    ("Türkiye Finans.xlsx", "turkiyefinans"),
    ("vakıfbank-nisan.xlsx", "vakifbank"),
    ("VAKIFBANK.xlsx", "vakifbank"),
    ("ekstre.xlsx", None),
    ("listesi.xls", None),
])
def test_addan_tespit(ad, beklenen):
    assert addan_tani(ad) == beklenen


def test_icerik_yoksa_dosya_adina_duser(tmp_path):
    satirlar = [["bilinmeyen", "düzen"], ["01.02.2026", "x", "y", 10.0]]
    assert icerikten_tani(satirlar) is None
    yol = tmp_path / "vakıfbank-ozel.xlsx"
    assert banka_tani(yol, rows=satirlar) == "vakifbank"


def test_tespit_edilemeyen_dosya_hata_verir(tmp_path):
    from openpyxl import Workbook
    wb = Workbook()
    wb.active.append(["a", "b"])
    yol = tmp_path / "ekstre.xlsx"
    wb.save(yol)
    with pytest.raises(TespitHatasi):
        ekstre_oku(yol)


def test_banka_elle_zorlanabilir(ornek_yolu, tmp_path):
    hedef = tmp_path / "dosya.xls"
    shutil.copy(ornek_yolu("garanti"), hedef)
    e = ekstre_oku(hedef, banka="Garanti BBVA")
    assert e.banka == "garanti" and len(e.islemler) == 9


@pytest.mark.parametrize("girdi, beklenen", [
    ("ziraat", "ziraat"), ("İş Bankası", "isbankasi"), ("t.finans", "turkiyefinans"), ("VakıfBank", "vakifbank"),
])
def test_banka_adini_cozumle(girdi, beklenen):
    assert banka_adini_cozumle(girdi) == beklenen


def test_bilinmeyen_banka_adi():
    with pytest.raises(ValueError):
        banka_adini_cozumle("olmayanbank")
