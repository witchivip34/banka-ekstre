# banka-ekstre

> **English summary:** `banka-ekstre` converts internet-banking statement exports (.xls/.xlsx) from six Turkish banks
> into one common transaction list and writes it as CSV, XLSX or an "F1–F6" accounting-import template.
> Everything runs locally, offline. See [README.en.md](README.en.md).

Türk bankalarının internet şubesinden indirilen ekstre dosyalarını (.xls/.xlsx) **tek bir ortak işlem listesine**
çevirir; CSV, XLSX ve muhasebe programı içe aktarım profilleri (F1–F6) üretir.

## Sorun

Her banka hesap hareketlerini başka bir düzende verir: biri .xls, biri .xlsx; tarih bir yerde `07/03/2026`, başka
yerde `07.03.2026 14:25`; tutar sütunu kimisinde 4., kimisinde 7. sütun; üstte değişen sayıda hesap bilgisi satırı,
altta dipnotlar. Ay sonunda bu dosyaları muhasebe programına aktarmak çoğu zaman elle kopyala-yapıştır demektir.

`banka-ekstre` her bankanın düzenini bilir, dosyayı içeriğinden tanır ve hepsini aynı yapıya çevirir:

| alan | açıklama |
|------|----------|
| `tarih` | `GG.AA.YYYY` |
| `aciklama` | bankanın işlem açıklaması |
| `referans` | fiş / dekont / işlem numarası |
| `tutar` | işaretli: **+ giren**, **− çıkan** |
| `bakiye` | işlem sonrası bakiye (varsa) |
| `banka`, `kaynak_dosya`, `satir_no` | izlenebilirlik için |

## Kurulum

Python 3.10 veya üstü gerekir.

```bash
git clone https://github.com/witchivip34-dev/banka-ekstre.git
cd banka-ekstre
pip install -e .            # xlrd, xlwt, openpyxl kurulur
pip install -e ".[excel]"   # isteğe bağlı: Windows'ta --excel-resave için pywin32
```

## Kullanım

```bash
# 1) Dosyanın (ya da klasördeki tüm dosyaların) bankasını, satır sayısını ve ilk 3 işlemini göster
banka-ekstre tani ornekler/

# 2) Bir klasördeki tüm ekstreleri tek bir Excel dosyasında birleştir
banka-ekstre cevir ornekler/ --cikti cikti/ --format xlsx --birlestir

# 3) Muhasebe programına aktarılacak F1–F6 dosyalarını üret (dosya başına bir .xls)
banka-ekstre cevir ornekler/ --cikti cikti/ --format f1f6 --ilk-satirlar-karisik
```

Toplamlar ve tarih aralığı için: `banka-ekstre ozet <dosya|klasör>`.
Komut kurulu değilse `python -m banka_ekstre ...` ile de çalışır.

`cevir` seçenekleri:

| seçenek | anlamı |
|---------|--------|
| `--format csv\|xlsx\|f1f6` | çıktı profili (varsayılan `csv`) |
| `--banka X` | tespiti atla, bankayı elle belirt (`ziraat`, `yapikredi`, `garanti`, `isbankasi`, `turkiyefinans`, `vakifbank`; "İş Bankası", "t.finans" gibi yazımlar da kabul edilir) |
| `--birlestir` | tüm girdileri tek dosyaya yaz (`birlesik-<format>.<uzantı>`); f1f6'da kolon yerleşimi bankaya bağlı olduğundan banka başına bir dosya (`birlesik-<banka>-f1f6.xls`) |
| `--profil p.json` | f1f6 için kendi profil dosyanız |
| `--ilk-satirlar-karisik` | f1f6: ilk 5 satırda hem giren hem çıkan işlem olmasını sağlar (aşağıya bakın) |
| `--excel-resave` | f1f6: dosyayı Excel ile 97-2003 biçiminde yeniden kaydeder (yalnız Windows + Excel + pywin32; yoksa uyarı verip atlar) |

Çıktı dosya adı: `<girdi-adı>-<format>.<uzantı>`, ör. `ziraat-ornek.xlsx` → `ziraat-ornek-csv.csv`,
`ziraat-ornek-f1f6.xls`. Komut sonunda dosya, banka, satır, giren toplamı ve çıkan toplamı içeren bir özet tablo basılır.

### Python'dan kullanım

```python
from banka_ekstre import ekstre_oku

e = ekstre_oku("ornekler/garanti-ornek.xls")
print(e.banka, len(e.islemler), e.giren_toplam, e.cikan_toplam)
for i in e.islemler:
    print(i.tarih, i.tutar, i.aciklama)
```

## Desteklenen bankalar

Banka, önce dosyanın **başlık satırındaki sütun adlarından** tanınır; bulunamazsa dosya adındaki banka kelimesine
bakılır (`ziraat`, `yapikredi`, `garanti`, `isbankasi`/`t.iş`, `tfinans`/`turkiyefinans`, `vakifbank`; büyük-küçük
harf ve Türkçe karakter farkı önemsiz). Aynı bankanın ek/alt hesap ekstreleri aynı düzende geldiği için aynı modülle okunur.

Ham sütun düzeni (0 tabanlı sütun indeksi):

| Banka | Modül | Uzantı | Tarih | Referans | Açıklama | Tutar (±) | Bakiye | Tanıma imzası (başlık satırı) |
|-------|-------|--------|-------|----------|----------|-----------|--------|-------------------------------|
| Ziraat Bankası | `ziraat` | .xlsx | 0 | 1 (Fiş No) | 2 | 3 | 4 | Tarih · Fiş No · Açıklama · İşlem Tutarı · Bakiye |
| Yapı Kredi | `yapikredi` | .xls | 0 | 4 (Referans No) | 5 | 6 | 7 | Tarih · Saat · … · Referans No · Açıklama · İşlem Tutarı |
| Garanti BBVA | `garanti` | .xls | 0 | 5 (Dekont No) | 1 | 3 | 4 | Tarih · Açıklama · Etiket · Tutar · … · Dekont No |
| Türkiye İş Bankası | `isbankasi` | .xls | 0 (`GG/AA/YYYY-SS:DD:ss`) | 14 (Referans) | 8 | 3 | 4 | Tarih/Saat · Valör · … · İşlem Tutarı* · … · Açıklama |
| Türkiye Finans | `turkiyefinans` | .xlsx | 0 (`GG.AA.YYYY SS:DD`) | 1 (İşlem Referansı) | 2 | 3 | 4 | İşlem Tarihi · İşlem Referansı · Açıklama · Tutar |
| VakıfBank | `vakifbank` | .xlsx | 2 (Hareket Tarih) | 9 (İşlem No) | 16 | 6 | 7 | … · HAREKET TARIH · … · TUTAR · … · İŞLEM NO · … · AÇIKLAMA |

Veri satırı sayılan satırlar: tarih sütununda geçerli bir tarih **ve** tutar sütununda sayı olanlar. Üstteki hesap
bilgisi blokları, ara başlıklar ve dipnotlar kendiliğinden atlanır.

`ornekler/` klasöründe her banka için **sentetik** (tamamen uydurma) bir ham dosya vardır; `python ornekler/uret.py`
ile yeniden üretilebilir.

## Çıktı profilleri

| format | dosya | içerik |
|--------|-------|--------|
| `csv` | `.csv` | UTF-8 BOM, `;` ayraç, ondalık ayırıcı virgül (Türkçe Excel çift tıkla doğru açar). Sütunlar: `tarih;banka;referans;aciklama;giren;cikan;bakiye;kaynak_dosya`. `giren` ve `cikan` pozitif yazılır, boş olan hücre boş kalır. |
| `xlsx` | `.xlsx` | Aynı sütunlar; başlık kalın, tutar sütunları `#,##0.00` sayı biçiminde, ilk satır sabit. |
| `f1f6` | `.xls` (97-2003) | Muhasebe programlarının "F1–F6" tipi banka hareketi içe aktarım şablonu (aşağıda). |

## F1–F6 profilini kendi programınıza uyarlama

Pek çok Türk muhasebe programı banka hareketlerini başlığı `F1 F2 F3 F4 F5 F6` olan 6 kolonlu bir .xls dosyasından
alır; hangi kolonun tarih, açıklama, giren veya çıkan olduğunu aktarım ekranında eşleştirirsiniz. Varsayılan yerleşim
`banka_ekstre/profiller/f1f6_varsayilan.json` dosyasındadır:

| Banka | F1 | F2 | F3 | F4 | F5 | F6 | çıkan işareti |
|-------|----|----|----|----|----|----|---------------|
| ziraat | tarih | boş | referans | açıklama | **giren** | **çıkan** | mutlak |
| yapikredi | tarih | boş | referans | açıklama | **giren** | **çıkan** | negatif |
| turkiyefinans | tarih | boş | referans | açıklama | **giren** | **çıkan** | negatif |
| garanti | tarih | boş | açıklama | **giren** | **çıkan** | referans | mutlak |
| isbankasi | tarih | boş | **giren** | **çıkan** | açıklama | referans | mutlak |
| vakifbank | tarih | boş | **giren** | **çıkan** | referans | açıklama | negatif |

Tutar kolonları `0.00` (sayı), diğerleri `@` (metin) biçiminde yazılır; sayfa adı `Sayfa`.

Kendi programınız farklı bir düzen bekliyorsa bu dosyayı kopyalayıp düzenleyin ve `--profil benim.json` ile verin:

```json
{
  "sayfa_adi": "Sayfa",
  "baslik": ["F1", "F2", "F3", "F4", "F5", "F6"],
  "bankalar": {
    "ziraat": {
      "kolonlar": ["tarih", "", "referans", "aciklama", "giren", "cikan"],
      "cikan_isareti": "mutlak",
      "kalin_baslik": [0, 1, 2, 3, 4, 5]
    }
  },
  "varsayilan": {
    "kolonlar": ["tarih", "", "referans", "aciklama", "giren", "cikan"],
    "cikan_isareti": "mutlak"
  }
}
```

- `kolonlar`: başlık sayısı kadar alan adı. Kullanılabilecek alanlar: `tarih`, `aciklama`, `referans`, `giren`,
  `cikan`, `bakiye`, `banka` ve `""` (boş kolon). Başlık sayısı 6 olmak zorunda değildir.
- `cikan_isareti`: `"mutlak"` → çıkan tutar pozitif yazılır (`125.40`); `"negatif"` → eksi işaretiyle (`-125.40`).
- `kalin_baslik`: başlık satırında kalın yazılacak kolonlar (0 tabanlı). Bazı programlar buna duyarlıdır.
- `varsayilan`: profilde adı geçmeyen bankalar için kullanılır.

**`--ilk-satirlar-karisik` neden var?** Bazı muhasebe programları giren/çıkan kolonlarını dosyanın ilk satırlarına
bakarak otomatik algılar; ilk satırların hepsi tek yönlüyse (ör. hepsi giren) diğer kolonu "yok" sayıp o tutarları
içe almaz. Bu seçenek, ilk 5 satır tek yönlüyse dosyadaki ilk ters yönlü işlemi 5. sıraya taşır (yer değiştirme;
hiçbir satır eklenmez ya da silinmez, toplamlar aynı kalır).

**`--excel-resave` neden var?** xlwt'nin ürettiği .xls dosyasını bazı programlar kararsız okur; aynı dosyanın Excel'in
kendisi tarafından 97-2003 biçiminde (`FileFormat=56`) yeniden kaydedilmesi bu sorunu giderir. Hücre içeriği değişmez.

## Yeni banka ekleme (5 adım)

1. Bankanın ham dosyasını `python -c "from banka_ekstre.readers import read_rows; [print(i, r) for i, r in enumerate(read_rows('dosya.xlsx')[:20])]"`
   ile açıp başlık satırını ve tarih / referans / açıklama / tutar / bakiye sütun indekslerini bulun.
2. `banka_ekstre/bankalar/` altına mevcut bir modülü (ör. `garanti.py`) kopyalayıp `AD`, `GORUNEN_AD`, `UZANTI`,
   `DUZEN = SutunDuzeni(...)` ve başlık imzası `IMZA = {sütun: "sadelestirilmis_ad"}` sabitlerini doldurun
   (imza Türkçe karaktersiz, küçük harf, boşluksuz yazılır: "İşlem Tutarı" → `"islemtutari"`).
3. Modülü `banka_ekstre/bankalar/__init__.py` içindeki `BANKALAR` listesine ekleyin; dosya adından tanınması için
   `banka_ekstre/tespit.py` içindeki `AD_KELIMELERI` listesine bir kelime ekleyin.
4. `banka_ekstre/profiller/f1f6_varsayilan.json` içine bankanın F1–F6 yerleşimini ekleyin (eklemezseniz `varsayilan` kullanılır).
5. `ornekler/uret.py` içine sentetik bir şablon ekleyip `python ornekler/uret.py` çalıştırın, `python -m pytest -q` ile doğrulayın.

## Sınırlamalar

- Bankalar internet şubesi çıktı düzenini haber vermeden değiştirebilir. Başlık imzası tutmazsa dosya adından tanıma
  devreye girer; sütunlar kaydıysa yanlış okuma olmaması için `tani` çıktısını kontrol edin.
- 2026'nın ilk yarısında indirilmiş TL vadesiz hesap ekstreleriyle test edildi. Döviz hesapları, kredi kartı ekstreleri
  ve PDF çıktılar desteklenmez.
- POS tahsilatlarında açıklamadaki "Tutar: X Komisyon: Y" gibi ayrıntılar ayrıştırılmaz; açıklama olduğu gibi aktarılır.
- `openpyxl` bazı bankaların .xlsx dosyalarında standart dışı stil sayfası yüzünden açılış hatası (stylesheet / Fill)
  veriyor. Bu yüzden .xlsx dosyaları doğrudan zip + XML olarak okunur; openpyxl yalnız xlsx **yazmak** için kullanılır.

## Gizlilik

Dosyalar yalnız yerelde okunur ve yazılır; araç hiçbir ağ bağlantısı kurmaz, telemetri yoktur. Gerçek ekstrelerinizi
depoya eklememeniz için `.gitignore` `ornekler/` dışındaki tüm `.xls`/`.xlsx` dosyalarını ve `cikti/` klasörünü dışlar.

## Geliştirme

```bash
pip install -e ".[test]"
python -m pytest -q
```

## Lisans

[MIT](LICENSE)
