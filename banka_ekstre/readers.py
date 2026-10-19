"""Ham ekstre dosyalarını satır listesine (list[list]) çeviren okuyucular.

- `.xls`  (Excel 97-2003, BIFF): xlrd ile okunur.
- `.xlsx` (Office Open XML): openpyxl yerine doğrudan zip + XML ile okunur.
  Bazı bankaların ürettiği .xlsx dosyalarında stil sayfası (styles.xml) standart dışıdır
  ve openpyxl "Fill" / stylesheet hatasıyla açmayı reddeder; ham XML okumak bu dosyaları
  da açar.

Tarih biçimli hücreler `DD.MM.YYYY` metnine çevrilir; diğer sayılar float, metinler str olur.
"""
from __future__ import annotations

import re
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

import xlrd

Hucre = object
Satir = list
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Excel'in yerleşik tarih/saat sayı biçimi kimlikleri
_YERLESIK_TARIH_BICIMLERI = set(range(14, 23)) | {45, 46, 47}


def read_rows(yol: str | Path) -> list[Satir]:
    """Dosya uzantısına göre ilk çalışma sayfasının tüm satırlarını döndürür."""
    yol = Path(yol)
    uzanti = yol.suffix.lower()
    if uzanti == ".xls":
        return read_xls_rows(yol)
    if uzanti == ".xlsx":
        return read_xlsx_rows(yol)
    raise ValueError(f"Desteklenmeyen uzantı: {yol.name} (yalnız .xls / .xlsx)")


# ─────────────────────────────── .xls ────────────────────────────────

def read_xls_rows(yol: Path) -> list[Satir]:
    """xlrd ile .xls okur; tarih hücreleri DD.MM.YYYY metnine çevrilir."""
    kitap = xlrd.open_workbook(str(yol), formatting_info=False)
    sayfa = kitap.sheet_by_index(0)
    satirlar: list[Satir] = []
    for r in range(sayfa.nrows):
        satir = []
        for c in range(sayfa.ncols):
            deger = sayfa.cell_value(r, c)
            if sayfa.cell_type(r, c) == xlrd.XL_CELL_DATE:
                try:
                    deger = xlrd.xldate.xldate_as_datetime(deger, kitap.datemode).strftime("%d.%m.%Y")
                except Exception:
                    pass
            satir.append(deger)
        satirlar.append(satir)
    return satirlar


# ─────────────────────────────── .xlsx ───────────────────────────────

def _sutun_harfinden_indeks(harfler: str) -> int:
    n = 0
    for ch in harfler:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _metin(eleman: ET.Element) -> str:
    return "".join(t.text or "" for t in eleman.iter() if t.tag.endswith("}t"))


def _paylasilan_metinler(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    with zf.open("xl/sharedStrings.xml") as f:
        kok = ET.parse(f).getroot()
    return [_metin(si) for si in kok.findall("x:si", NS)]


def _bicim_tarih_mi(kod: str) -> bool:
    """Özel sayı biçimi kodu tarih/saat içeriyor mu?"""
    sade = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", kod.lower())
    return bool(re.search(r"[dy]", sade)) or bool(re.search(r"m.*[/.\-].*", sade))


def _tarih_stilleri(zf: zipfile.ZipFile) -> set[int]:
    """Tarih biçimli hücre stillerinin (cellXfs) indekslerini döndürür."""
    if "xl/styles.xml" not in zf.namelist():
        return set()
    try:
        with zf.open("xl/styles.xml") as f:
            kok = ET.parse(f).getroot()
    except ET.ParseError:
        return set()
    ozel = {}
    numfmts = kok.find("x:numFmts", NS)
    if numfmts is not None:
        for nf in numfmts.findall("x:numFmt", NS):
            try:
                ozel[int(nf.get("numFmtId", "-1"))] = nf.get("formatCode", "")
            except ValueError:
                continue
    sonuc: set[int] = set()
    xfs = kok.find("x:cellXfs", NS)
    if xfs is None:
        return sonuc
    for i, xf in enumerate(xfs.findall("x:xf", NS)):
        try:
            nid = int(xf.get("numFmtId", "0"))
        except ValueError:
            continue
        if nid in _YERLESIK_TARIH_BICIMLERI or (nid in ozel and _bicim_tarih_mi(ozel[nid])):
            sonuc.add(i)
    return sonuc


def _tarih_1904_mu(zf: zipfile.ZipFile) -> bool:
    try:
        with zf.open("xl/workbook.xml") as f:
            kok = ET.parse(f).getroot()
    except (KeyError, ET.ParseError):
        return False
    pr = kok.find("x:workbookPr", NS)
    return pr is not None and pr.get("date1904") in ("1", "true")


def _ilk_sayfa_yolu(zf: zipfile.ZipFile) -> str:
    """workbook.xml'deki ilk sayfanın dosya yolunu bulur; olmazsa ilk worksheet dosyası."""
    adlar = zf.namelist()
    try:
        with zf.open("xl/workbook.xml") as f:
            wb = ET.parse(f).getroot()
        with zf.open("xl/_rels/workbook.xml.rels") as f:
            rels = ET.parse(f).getroot()
        ilk = wb.find("x:sheets/x:sheet", NS)
        rid = ilk.get(f"{{{NS_REL}}}id") if ilk is not None else None
        for rel in rels.findall(f"{{{NS_PKG_REL}}}Relationship"):
            if rel.get("Id") == rid:
                hedef = rel.get("Target", "").lstrip("/")
                yol = hedef if hedef.startswith("xl/") else "xl/" + hedef
                if yol in adlar:
                    return yol
    except (KeyError, ET.ParseError, AttributeError):
        pass
    return next(n for n in adlar if n.startswith("xl/worksheets/") and n.endswith(".xml"))


def _seri_tarihten(deger: float, tarih_1904: bool) -> str:
    baslangic = datetime(1904, 1, 1) if tarih_1904 else datetime(1899, 12, 30)
    return (baslangic + timedelta(days=deger)).strftime("%d.%m.%Y")


def read_xlsx_rows(yol: Path) -> list[Satir]:
    """xlsx'i zip + XML ile okur.

    Hücrelerde `r=` (A1 adresi) olmayabilir; o durumda hücreler ardışık kabul edilir.
    """
    with zipfile.ZipFile(yol) as zf:
        paylasilan = _paylasilan_metinler(zf)
        tarih_stilleri = _tarih_stilleri(zf)
        tarih_1904 = _tarih_1904_mu(zf)
        with zf.open(_ilk_sayfa_yolu(zf)) as f:
            kok = ET.parse(f).getroot()

    sheet_data = kok.find("x:sheetData", NS)
    if sheet_data is None:
        return []
    satirlar: list[Satir] = []
    for satir_el in sheet_data.findall("x:row", NS):
        satir_no = satir_el.get("r")
        # Boş satırlar dosyada yer almayabilir; satır numarasını korumak için doldur.
        if satir_no and satir_no.isdigit():
            while len(satirlar) < int(satir_no) - 1:
                satirlar.append([])
        hucreler = _satir_hucreleri(satir_el, paylasilan, tarih_stilleri, tarih_1904)
        if not hucreler:
            satirlar.append([])
            continue
        satirlar.append([hucreler.get(i, "") for i in range(max(hucreler) + 1)])
    return satirlar


def _satir_hucreleri(satir_el, paylasilan, tarih_stilleri, tarih_1904) -> dict[int, Hucre]:
    hucreler: dict[int, Hucre] = {}
    sira = 0
    for c in satir_el.findall("x:c", NS):
        adres = c.get("r")
        m = re.match(r"([A-Z]+)", adres) if adres else None
        sutun = _sutun_harfinden_indeks(m.group(1)) if m else sira
        sira = sutun + 1
        hucreler[sutun] = _hucre_degeri(c, paylasilan, tarih_stilleri, tarih_1904)
    return hucreler


def _hucre_degeri(c, paylasilan, tarih_stilleri, tarih_1904) -> Hucre:
    tur = c.get("t")
    v = c.find("x:v", NS)
    satir_ici = c.find("x:is", NS)
    if v is not None:
        ham = v.text
    elif satir_ici is not None:
        return _metin(satir_ici)
    else:
        return ""
    if tur == "s":
        try:
            return paylasilan[int(ham)]
        except (ValueError, IndexError, TypeError):
            return ham or ""
    if tur in ("str", "inlineStr", "e"):
        return ham or ""
    if tur == "b":
        return bool(int(ham))
    try:
        sayi = float(ham)
    except (TypeError, ValueError):
        return ham or ""
    stil = c.get("s")
    if stil is not None and stil.isdigit() and int(stil) in tarih_stilleri:
        return _seri_tarihten(sayi, tarih_1904)
    return sayi
