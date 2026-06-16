# Chat Workspace Notes

For tasks in this chat that mention prices, materials, installation, montage, equipment, systems, or the price database, use this database first:

`C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite`

It currently includes prices imported from:

- `C:\Users\admla\Downloads\02.06 ЦВСИГ.xlsx`
- `C:\Users\admla\Desktop\РВС.xls`
- `C:\Users\admla\Desktop\Протей\Протей.xlsx`
- `C:\Users\admla\Downloads\Дагестан 23.04.xls`
- `C:\Users\admla\Documents\Codex\2026-04-24\excel\reliable_price_database_v2.csv`
- `C:\Users\admla\Downloads\ЦВСИГ (16.06) - 18.xlsx`
- `C:\Users\admla\Downloads\КП Технология - Слаботочные системы от 16.06.xlsx`

The 2026-04-24 reliable CSV is a trusted unit-price reference. It stores recommended unit prices with VAT and status groups `AUTO_OK`, `CHECK`, and `MANUAL`; use status and note fields when judging confidence.

VAT rule confirmed by the user on 2026-06-08: all uploaded price documents should be treated as prices with VAT, even when a source column header says "без НДС". Do not multiply or divide historical source prices because of source header text.

Quick context and query helper:

- `work\DATABASE_CONTEXT.md`
- `work\query_prices.py`

Do not re-import the source workbooks unless the user asks to update or refresh the database from Excel.
