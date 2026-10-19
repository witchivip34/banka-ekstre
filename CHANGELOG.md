# Değişiklik günlüğü

## 0.1.0 — 2026-10-15

İlk sürüm.

- Altı banka: Ziraat Bankası, Yapı Kredi, Garanti BBVA, Türkiye İş Bankası, Türkiye Finans, VakıfBank.
- Ortak işlem modeli (`Islem`): tarih, açıklama, referans, işaretli tutar, bakiye, banka, kaynak dosya, satır no.
- Banka tespiti: önce başlık satırı imzası, sonra dosya adı; `--banka` ile elle zorlama.
- `.xlsx` için openpyxl'den bağımsız zip + XML okuyucu (tarih biçimli hücreler `GG.AA.YYYY` olur).
- Çıktılar: CSV (UTF-8 BOM, `;`), XLSX, F1–F6 muhasebe içe aktarım şablonu (.xls) ve JSON profil dosyası.
- F1–F6 seçenekleri: `--profil`, `--ilk-satirlar-karisik`, `--excel-resave` (Windows + pywin32).
- Komutlar: `tani`, `cevir` (`--birlestir` dahil), `ozet`.
- Sentetik örnek ekstreler (`ornekler/uret.py`) ve pytest test takımı.
