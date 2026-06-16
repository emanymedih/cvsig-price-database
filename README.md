# CVSIG Price Database Workspace

Рабочий проект с собранной базой цен по материалам и монтажу.

## Главный файл базы

`outputs\cvsig_prices_2026-06-02.sqlite`

База уже включает источники:

- `02.06 ЦВСИГ.xlsx`
- `РВС.xls`
- `Протей.xlsx`
- `Дагестан 23.04.xls`
- `reliable_price_database_v2.csv`
- `ЦВСИГ (16.06) - 18.xlsx`

Все цены из пользовательских источников считаются ценами с НДС, даже если исходный заголовок говорит "без НДС".

## Быстрый просмотр базы

```powershell
& 'C:\Users\admla\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\work\query_prices.py' summary
```

Поиск:

```powershell
& 'C:\Users\admla\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\work\query_prices.py' search "С2000"
```

SQL только для чтения:

```powershell
& 'C:\Users\admla\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '.\work\query_prices.py' sql "SELECT * FROM v_price_lines LIMIT 10"
```

## Важные файлы

- `AGENTS.md`: инструкция для Codex в этом проекте.
- `work\DATABASE_CONTEXT.md`: описание базы и источников.
- `work\query_prices.py`: быстрый помощник для поиска и SQL-запросов.
- `work\import_cvsig_20260616_prices.py`: импорт файла `ЦВСИГ (16.06) - 18.xlsx`.
