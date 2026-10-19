# banka-ekstre (English)

`banka-ekstre` reads the account-statement exports (.xls / .xlsx) that Turkish internet-banking portals produce and
turns them into one common transaction list: date (`DD.MM.YYYY`), description, reference, signed amount
(+ credit / − debit), balance, bank, source file and row number. The full documentation is in Turkish: [README.md](README.md).

**Supported banks:** Ziraat Bankası, Yapı Kredi, Garanti BBVA, Türkiye İş Bankası, Türkiye Finans, VakıfBank.
The bank is detected from the column headers of the file, falling back to a bank keyword in the file name.

**Install:** `pip install -e .` (Python ≥ 3.10; depends on `xlrd`, `xlwt`, `openpyxl`; optional `pywin32` via `.[excel]`).

**Commands:**

```bash
banka-ekstre tani ornekler/                                   # detect bank, row count, first 3 transactions
banka-ekstre ozet ornekler/                                   # totals and date range
banka-ekstre cevir ornekler/ --cikti out/ --format xlsx --birlestir
banka-ekstre cevir ornekler/ --cikti out/ --format f1f6 --ilk-satirlar-karisik
```

**Output formats:**

- `csv` — UTF-8 with BOM, `;` separator, comma decimal (opens correctly in Turkish-locale Excel).
- `xlsx` — same columns, bold header, `#,##0.00` amounts.
- `f1f6` — a 6-column `.xls` (Excel 97-2003) template with headers `F1…F6`, used by many Turkish accounting packages
  to import bank movements. The column layout per bank and the sign of debits come from a JSON profile
  (`banka_ekstre/profiller/f1f6_varsayilan.json`); pass your own with `--profil my.json`.
  `--ilk-satirlar-karisik` makes sure the first 5 rows contain both a credit and a debit (some importers auto-detect
  columns from the first rows). `--excel-resave` re-saves the file through Excel (Windows + pywin32 only).

**Privacy:** everything runs locally; no network access. The sample files in `ornekler/` are fully synthetic.

**Limitations:** banks may change their export layout without notice; tested with TRY current-account exports from the
first half of 2026.

License: MIT.
