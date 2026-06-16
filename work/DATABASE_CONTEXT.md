# Database Context

Use this file as the starting point for future tasks about prices, materials, installation, montage, systems, or equipment.

## Active Database

`C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite`

## Source Workbooks

- `C:\Users\admla\Downloads\02.06 ЦВСИГ.xlsx`
- `C:\Users\admla\Desktop\РВС.xls`
- `C:\Users\admla\Desktop\Протей\Протей.xlsx`
- `C:\Users\admla\Downloads\Дагестан 23.04.xls`
- `C:\Users\admla\Documents\Codex\2026-04-24\excel\reliable_price_database_v2.csv`
- `C:\Users\admla\Downloads\ЦВСИГ (16.06) - 18.xlsx`
- `C:\Users\admla\Downloads\КП Технология - Слаботочные системы от 16.06.xlsx`

Snapshot dates: `2026-06-02`, `2026-05-26`, `2026-04-23`, `2026-04-24`, and `2026-06-16`.

## What Is Stored

- 2 960 price positions / reference rows.
- 5 267 price records.
- Materials: 2 797 records, total `187 531 621.93` RUB.
- Installation: 2 470 records, total `251 401 276.44` RUB.
- Unit-price reference rows from `Надежная база 2026-04-24` do not have total sums, so they do not change material/work total totals above.
- Systems include `АПС`, `ОТС`, `СКС`, `СКУД`, `СОТ`, `РВС`, `Протей`, `Дагестан / ...`, `КП Технология / ...`, plus `Надежная база 2026-04-24 / AUTO_OK`, `/ CHECK`, `/ MANUAL`.
- User clarified on 2026-06-08: all uploaded source prices are `руб. с НДС`, even if a source column header says `без НДС`. Current `prices.vat_note` values follow this rule for imported user documents.

## Tables

- `price_snapshots`: source file metadata and import date.
- `price_lines`: one row per original Excel position.
- `prices`: separate records for `material` and `installation`.

## Views To Use First

`v_price_lines` gives one row per position:

- `price_date`
- `system`
- `position`
- `name`
- `type_code_article`
- `manufacturer`
- `unit`
- `quantity`
- `material_unit_price_rub`
- `material_total_price_rub`
- `material_vat_note`
- `installation_unit_price_rub`
- `installation_total_price_rub`
- `installation_vat_note`
- `note`
- `excel_row`

`v_prices` gives normalized price records:

- `price_kind` is `material` or `installation`.
- `unit_price_rub` and `total_price_rub` are RUB values.
- `vat_note` states the active VAT interpretation.

## Quick Commands

Show summary:

```powershell
& 'C:\Users\admla\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\work\query_prices.py' summary
```

Search by name, code, manufacturer, or system:

```powershell
& 'C:\Users\admla\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\work\query_prices.py' search "С2000"
```

Run a SQL query:

```powershell
& 'C:\Users\admla\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\work\query_prices.py' sql "SELECT * FROM v_price_lines WHERE system = 'АПС' LIMIT 5"
```

## Rule For Future Work

Do not parse the Excel workbooks again unless the user asks to update or re-import prices. Query the SQLite database first.
