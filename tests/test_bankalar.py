"""Her banka için sentetik dosyanın ayrıştırılması."""
from __future__ import annotations

import re

import pytest

from banka_ekstre import ekstre_oku
from banka_ekstre.bankalar import BANKA_MODULLERI
from banka_ekstre.readers import read_rows
from conftest import BANKALAR, ORNEKLER, uret


@pytest.mark.parametrize("banka", BANKALAR)
def test_parse_satir_sayisi_isaret_tarih(banka, ornek_yolu):
    beklenen = uret.ornek_islemler(banka)
    e = ekstre_oku(ornek_yolu(banka))
    assert e.banka == banka
    assert len(e.islemler) == len(beklenen) == uret.ADETLER[banka]
    for islem, b in zip(e.islemler, beklenen):
        assert re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", islem.tarih)
        assert islem.tarih == b.zaman.strftime("%d.%m.%Y")
        assert islem.tutar == pytest.approx(b.tutar)
        assert islem.bakiye == pytest.approx(b.bakiye)
        assert islem.aciklama == b.aciklama
        assert (islem.giren > 0) == (b.tutar > 0)
        assert (islem.cikan > 0) == (b.tutar < 0)
        assert islem.giren >= 0 and islem.cikan >= 0
        assert islem.kaynak_dosya == uret.DOSYA_ADLARI[banka]
        assert islem.satir_no > 1


@pytest.mark.parametrize("banka", BANKALAR)
def test_referans_alinir(banka, ornek_yolu):
    e = ekstre_oku(ornek_yolu(banka))
    beklenen = uret.ornek_islemler(banka)
    for islem, b in zip(e.islemler, beklenen):
        if banka == "ziraat":  # Ziraat'te referans = kısa fiş no
            assert islem.referans == f"{int(b.referans) % 100000:05d}"
        else:
            assert islem.referans == b.referans


@pytest.mark.parametrize("banka", BANKALAR)
def test_depodaki_ornekler_guncel(banka):
    """Depoya konmuş örnek dosyalar uret.py çıktısıyla aynı işlemleri içerir."""
    e = ekstre_oku(ORNEKLER / uret.DOSYA_ADLARI[banka])
    assert [i.tutar for i in e.islemler] == pytest.approx([b.tutar for b in uret.ornek_islemler(banka)])


@pytest.mark.parametrize("banka", BANKALAR)
def test_uzanti_sabiti(banka):
    assert uret.DOSYA_ADLARI[banka].endswith(BANKA_MODULLERI[banka].UZANTI)


def test_baslik_ve_dipnot_satirlari_atlanir(ornek_yolu):
    satirlar = read_rows(ornek_yolu("isbankasi"))
    e = ekstre_oku(ornek_yolu("isbankasi"))
    assert len(satirlar) > len(e.islemler) + 15


def test_xlsx_tarih_hucresi_metne_cevrilir(tmp_path):
    """Gerçek tarih hücresi (seri sayı + tarih biçimi) DD.MM.YYYY metni olarak okunur."""
    from datetime import datetime

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["Tarih", "Fiş No", "Açıklama", "İşlem Tutarı", "Bakiye"])
    ws.append([datetime(2026, 3, 7), "1", "HAVALE GELEN", 150.0, 150.0])
    ws.append([datetime(2026, 3, 8), "2", "KİRA ÖDEMESİ", -100.0, 50.0])
    yol = tmp_path / "tarihli.xlsx"
    wb.save(yol)
    assert read_rows(yol)[1][0] == "07.03.2026"
    e = ekstre_oku(yol)
    assert e.banka == "ziraat"
    assert [i.tarih for i in e.islemler] == ["07.03.2026", "08.03.2026"]
    assert read_rows(yol)[1][3] == 150.0  # tutar sayı olarak kalır
