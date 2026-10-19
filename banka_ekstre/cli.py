"""`banka-ekstre` komut satırı arayüzü."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .bankalar import BANKA_MODULLERI
from .cikti import bordro_f1f6
from .cikti import csv as csv_cikti
from .cikti import xlsx as xlsx_cikti
from .ekstre import Ekstre, dosyalari_bul, ekstre_oku
from .model import Islem, tarih_araligi, toplamlar

FORMATLAR = {"csv": csv_cikti, "xlsx": xlsx_cikti, "f1f6": bordro_f1f6}


# ─────────────────────────────── yardımcılar ───────────────────────────────

def para(x: float) -> str:
    """1234567.8 → "1.234.567,80"."""
    return f"{x:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def _ekstreleri_oku(girdi: str, banka: str | None) -> tuple[list[Ekstre], int]:
    """Girdi altındaki dosyaları okur; (başarılı ekstreler, hata sayısı) döndürür."""
    ekstreler: list[Ekstre] = []
    hata = 0
    for yol in dosyalari_bul(girdi):
        try:
            ekstreler.append(ekstre_oku(yol, banka))
        except Exception as e:  # tek dosyadaki hata diğerlerini durdurmasın
            print(f"  HATA  {yol.name}: {e}", file=sys.stderr)
            hata += 1
    return ekstreler, hata


def _ozet_tablosu(satirlar: list[tuple[str, str, int, float, float]]) -> None:
    genislik = max([len("dosya")] + [len(s[0]) for s in satirlar])
    print(f"\n{'dosya':<{genislik}}  {'banka':<13} {'satır':>6} {'giren':>18} {'çıkan':>18}")
    print("-" * (genislik + 59))
    for ad, banka, n, g, c in satirlar:
        print(f"{ad:<{genislik}}  {banka:<13} {n:>6} {para(g):>18} {para(c):>18}")
    if len(satirlar) > 1:
        print("-" * (genislik + 59))
        print(f"{'TOPLAM':<{genislik}}  {'':<13} {sum(s[2] for s in satirlar):>6} "
              f"{para(sum(s[3] for s in satirlar)):>18} {para(sum(s[4] for s in satirlar)):>18}")


def _islem_satiri(i: Islem) -> str:
    return f"    {i.tarih}  {para(i.tutar):>14}  {i.referans[:18]:<18}  {i.aciklama[:50]}"


# ─────────────────────────────── komutlar ───────────────────────────────

def komut_tani(args: argparse.Namespace) -> int:
    ekstreler, hata = _ekstreleri_oku(args.girdi, args.banka)
    for e in ekstreler:
        print(f"{e.dosya.name}: banka={e.banka}  satır={len(e.islemler)}")
        for i in e.islemler[:3]:
            print(_islem_satiri(i))
    return 1 if hata else 0


def komut_ozet(args: argparse.Namespace) -> int:
    ekstreler, hata = _ekstreleri_oku(args.girdi, args.banka)
    tum: list[Islem] = []
    for e in ekstreler:
        aralik = tarih_araligi(e.islemler)
        aralik_metin = f"{aralik[0]} – {aralik[1]}" if aralik else "-"
        g, c = toplamlar(e.islemler)
        print(f"{e.dosya.name}: {e.banka}, {len(e.islemler)} işlem, {aralik_metin}, "
              f"giren {para(g)}, çıkan {para(c)}, net {para(g - c)}")
        tum.extend(e.islemler)
    if len(ekstreler) > 1:
        aralik = tarih_araligi(tum)
        g, c = toplamlar(tum)
        print(f"\nTOPLAM: {len(tum)} işlem, {aralik[0] if aralik else '-'} – {aralik[1] if aralik else '-'}, "
              f"giren {para(g)}, çıkan {para(c)}, net {para(g - c)}")
    return 1 if hata else 0


def _f1f6_yaz(islemler: list[Islem], hedef: Path, banka: str, args, profil) -> Path:
    yol = bordro_f1f6.yaz(islemler, hedef, banka, profil, args.ilk_satirlar_karisik)
    if args.excel_resave:
        bordro_f1f6.excel_ile_yeniden_kaydet(yol, profil.yerlesim(banka).kalin_baslik)
    return yol


def komut_cevir(args: argparse.Namespace) -> int:
    ekstreler, hata = _ekstreleri_oku(args.girdi, args.banka)
    if not ekstreler:
        print("Çevrilecek dosya yok.", file=sys.stderr)
        return 1
    cikti_dir = Path(args.cikti)
    modul = FORMATLAR[args.format]
    profil = bordro_f1f6.profil_yukle(args.profil) if args.format == "f1f6" else None
    if args.format != "f1f6" and (args.ilk_satirlar_karisik or args.excel_resave or args.profil):
        print("  [UYARI] --profil / --ilk-satirlar-karisik / --excel-resave yalnız f1f6 içindir.", file=sys.stderr)

    yazilanlar: list[Path] = []
    if args.birlestir:
        yazilanlar += _birlesik_yaz(ekstreler, cikti_dir, args, modul, profil)
    else:
        for e in ekstreler:
            hedef = cikti_dir / f"{e.dosya.stem}-{args.format}{modul.UZANTI}"
            if args.format == "f1f6":
                yazilanlar.append(_f1f6_yaz(e.islemler, hedef, e.banka, args, profil))
            else:
                yazilanlar.append(modul.yaz(e.islemler, hedef))

    _ozet_tablosu([(e.dosya.name, e.banka, len(e.islemler), *toplamlar(e.islemler)) for e in ekstreler])
    print("\nYazılan dosyalar:")
    for y in yazilanlar:
        print(f"  {y}")
    return 1 if hata else 0


def _birlesik_yaz(ekstreler, cikti_dir: Path, args, modul, profil) -> list[Path]:
    """csv/xlsx: tek dosya. f1f6: kolon yerleşimi bankaya bağlı olduğundan banka başına bir dosya."""
    if args.format != "f1f6":
        tum = [i for e in ekstreler for i in e.islemler]
        return [modul.yaz(tum, cikti_dir / f"birlesik-{args.format}{modul.UZANTI}")]
    yollar = []
    for banka in dict.fromkeys(e.banka for e in ekstreler):
        islemler = [i for e in ekstreler if e.banka == banka for i in e.islemler]
        hedef = cikti_dir / f"birlesik-{banka}-f1f6{modul.UZANTI}"
        yollar.append(_f1f6_yaz(islemler, hedef, banka, args, profil))
    return yollar


# ─────────────────────────────── argparse ───────────────────────────────

def parser_olustur() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="banka-ekstre",
        description="Türk bankalarının ekstre dosyalarını (.xls/.xlsx) ortak bir işlem listesine çevirir.",
    )
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    alt = p.add_subparsers(dest="komut", required=True)
    bankalar = ", ".join(BANKA_MODULLERI)

    def ortak(sp: argparse.ArgumentParser) -> None:
        sp.add_argument("girdi", help="ekstre dosyası ya da klasör")
        sp.add_argument("--banka", help=f"tespiti atla, bankayı zorla ({bankalar})")

    t = alt.add_parser("tani", help="dosyanın bankasını, satır sayısını ve ilk 3 işlemi göster")
    ortak(t)
    t.set_defaults(func=komut_tani)

    o = alt.add_parser("ozet", help="toplamlar ve tarih aralığı")
    ortak(o)
    o.set_defaults(func=komut_ozet)

    c = alt.add_parser("cevir", help="ekstreleri csv / xlsx / f1f6 biçimine çevir")
    ortak(c)
    c.add_argument("--cikti", required=True, help="çıktı klasörü")
    c.add_argument("--format", choices=sorted(FORMATLAR), default="csv", help="çıktı profili (varsayılan: csv)")
    c.add_argument("--profil", help="f1f6 için kendi JSON profil dosyanız")
    c.add_argument("--birlestir", action="store_true",
                   help="tüm girdileri tek dosyada birleştir (f1f6'da banka başına bir dosya)")
    c.add_argument("--ilk-satirlar-karisik", action="store_true",
                   help="f1f6: ilk 5 satırda hem giren hem çıkan işlem olsun")
    c.add_argument("--excel-resave", action="store_true",
                   help="f1f6: Excel ile 97-2003 biçiminde yeniden kaydet (Windows + pywin32)")
    c.set_defaults(func=komut_cevir)
    return p


def main(argv: list[str] | None = None) -> int:
    for akis in (sys.stdout, sys.stderr):
        try:
            akis.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    args = parser_olustur().parse_args(argv)
    try:
        return args.func(args)
    except (FileNotFoundError, ValueError) as e:
        print(f"HATA: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
