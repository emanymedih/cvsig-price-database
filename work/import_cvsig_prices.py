from __future__ import annotations

import datetime as dt
import hashlib
import json
import sqlite3
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl


SOURCE = Path(r"C:\Users\admla\Downloads\02.06 ЦВСИГ.xlsx")
OUTPUT = Path(r"C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite")
PRICE_DATE = "2026-06-02"


def money_to_kopeks(value):
    if value is None or value == "":
        return None
    amount = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return int(amount * 100)


def num_or_none(value):
    if value is None or value == "":
        return None
    return float(value)


def text_or_none(value):
    if value is None:
        return None
    value = str(value).strip()
    return value or None


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def iter_price_rows(path: Path):
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for sheet in workbook.worksheets:
        for excel_row, row in enumerate(sheet.iter_rows(min_row=4, values_only=True), start=4):
            position = row[0]
            if not isinstance(position, (int, float)):
                continue

            quantity = num_or_none(row[5])
            material_unit = money_to_kopeks(row[6])
            material_total = money_to_kopeks(row[7])
            install_unit = money_to_kopeks(row[8])
            install_total = money_to_kopeks(row[9])

            has_price = any(
                value is not None
                for value in (material_unit, material_total, install_unit, install_total)
            )
            if not has_price:
                continue

            yield {
                "system": sheet.title,
                "excel_row": excel_row,
                "position": int(position) if float(position).is_integer() else float(position),
                "name": text_or_none(row[1]),
                "type_code_article": text_or_none(row[2]),
                "manufacturer": text_or_none(row[3]),
                "unit": text_or_none(row[4]),
                "quantity": quantity,
                "material_unit_kopeks": material_unit,
                "material_total_kopeks": material_total,
                "installation_unit_kopeks": install_unit,
                "installation_total_kopeks": install_total,
                "note": text_or_none(row[10]),
            }


SCHEMA = """
PRAGMA foreign_keys = ON;

DROP VIEW IF EXISTS v_price_lines;
DROP VIEW IF EXISTS v_prices;
DROP TABLE IF EXISTS prices;
DROP TABLE IF EXISTS price_lines;
DROP TABLE IF EXISTS price_snapshots;

CREATE TABLE price_snapshots (
    id INTEGER PRIMARY KEY,
    source_file TEXT NOT NULL,
    source_sha256 TEXT NOT NULL,
    source_last_modified TEXT NOT NULL,
    price_date TEXT NOT NULL,
    imported_at TEXT NOT NULL,
    source_headers_json TEXT NOT NULL
);

CREATE TABLE price_lines (
    id INTEGER PRIMARY KEY,
    snapshot_id INTEGER NOT NULL REFERENCES price_snapshots(id) ON DELETE CASCADE,
    system TEXT NOT NULL,
    excel_row INTEGER NOT NULL,
    position TEXT NOT NULL,
    name TEXT,
    type_code_article TEXT,
    manufacturer TEXT,
    unit TEXT,
    quantity REAL,
    note TEXT,
    UNIQUE (snapshot_id, system, excel_row)
);

CREATE TABLE prices (
    id INTEGER PRIMARY KEY,
    line_id INTEGER NOT NULL REFERENCES price_lines(id) ON DELETE CASCADE,
    price_kind TEXT NOT NULL CHECK (price_kind IN ('material', 'installation')),
    unit_price_kopeks INTEGER,
    total_price_kopeks INTEGER,
    currency TEXT NOT NULL DEFAULT 'RUB',
    vat_note TEXT NOT NULL DEFAULT 'по исходной таблице: руб. без НДС',
    UNIQUE (line_id, price_kind)
);

CREATE VIEW v_prices AS
SELECT
    s.price_date,
    l.system,
    l.position,
    l.name,
    l.type_code_article,
    l.manufacturer,
    l.unit,
    l.quantity,
    p.price_kind,
    p.unit_price_kopeks / 100.0 AS unit_price_rub,
    p.total_price_kopeks / 100.0 AS total_price_rub,
    p.currency,
    p.vat_note,
    l.note,
    l.excel_row
FROM prices p
JOIN price_lines l ON l.id = p.line_id
JOIN price_snapshots s ON s.id = l.snapshot_id;

CREATE VIEW v_price_lines AS
SELECT
    s.price_date,
    l.system,
    l.position,
    l.name,
    l.type_code_article,
    l.manufacturer,
    l.unit,
    l.quantity,
    MAX(CASE WHEN p.price_kind = 'material' THEN p.unit_price_kopeks END) / 100.0 AS material_unit_price_rub,
    MAX(CASE WHEN p.price_kind = 'material' THEN p.total_price_kopeks END) / 100.0 AS material_total_price_rub,
    MAX(CASE WHEN p.price_kind = 'material' THEN p.vat_note END) AS material_vat_note,
    MAX(CASE WHEN p.price_kind = 'installation' THEN p.unit_price_kopeks END) / 100.0 AS installation_unit_price_rub,
    MAX(CASE WHEN p.price_kind = 'installation' THEN p.total_price_kopeks END) / 100.0 AS installation_total_price_rub,
    MAX(CASE WHEN p.price_kind = 'installation' THEN p.vat_note END) AS installation_vat_note,
    l.note,
    l.excel_row
FROM price_lines l
JOIN price_snapshots s ON s.id = l.snapshot_id
LEFT JOIN prices p ON p.line_id = l.id
GROUP BY l.id;
"""


def main():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    if OUTPUT.exists():
        OUTPUT.unlink()

    rows = list(iter_price_rows(SOURCE))
    source_stat = SOURCE.stat()
    source_headers = {
        "columns": [
            "поз.",
            "Наименование",
            "Тип / код / артикул",
            "Изготовитель",
            "Ед.",
            "Кол-во",
            "Стоимость материала: ед. расценка, руб. без НДС",
            "Стоимость материала: итого, руб. без НДС",
            "Стоимость монтажа: ед. расценка, руб. без НДС",
            "Стоимость монтажа: итого, руб. без НДС",
            "Примечание",
        ]
    }

    with sqlite3.connect(OUTPUT) as conn:
        conn.executescript(SCHEMA)
        cursor = conn.execute(
            """
            INSERT INTO price_snapshots (
                source_file,
                source_sha256,
                source_last_modified,
                price_date,
                imported_at,
                source_headers_json
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                str(SOURCE),
                file_sha256(SOURCE),
                dt.datetime.fromtimestamp(source_stat.st_mtime).isoformat(timespec="seconds"),
                PRICE_DATE,
                dt.datetime.now().isoformat(timespec="seconds"),
                json.dumps(source_headers, ensure_ascii=False),
            ),
        )
        snapshot_id = cursor.lastrowid

        for row in rows:
            cursor = conn.execute(
                """
                INSERT INTO price_lines (
                    snapshot_id,
                    system,
                    excel_row,
                    position,
                    name,
                    type_code_article,
                    manufacturer,
                    unit,
                    quantity,
                    note
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    snapshot_id,
                    row["system"],
                    row["excel_row"],
                    str(row["position"]),
                    row["name"],
                    row["type_code_article"],
                    row["manufacturer"],
                    row["unit"],
                    row["quantity"],
                    row["note"],
                ),
            )
            line_id = cursor.lastrowid

            if row["material_unit_kopeks"] is not None or row["material_total_kopeks"] is not None:
                conn.execute(
                    """
                    INSERT INTO prices (line_id, price_kind, unit_price_kopeks, total_price_kopeks)
                    VALUES (?, 'material', ?, ?)
                    """,
                    (line_id, row["material_unit_kopeks"], row["material_total_kopeks"]),
                )

            if row["installation_unit_kopeks"] is not None or row["installation_total_kopeks"] is not None:
                conn.execute(
                    """
                    INSERT INTO prices (line_id, price_kind, unit_price_kopeks, total_price_kopeks)
                    VALUES (?, 'installation', ?, ?)
                    """,
                    (line_id, row["installation_unit_kopeks"], row["installation_total_kopeks"]),
                )

        checks = {
            "output": str(OUTPUT),
            "lines": conn.execute("SELECT COUNT(*) FROM price_lines").fetchone()[0],
            "price_records": conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0],
            "by_kind": [
                {
                    "price_kind": price_kind,
                    "records": records,
                    "total_rub": total_rub,
                }
                for price_kind, records, total_rub in conn.execute(
                    "SELECT price_kind, COUNT(*), SUM(total_price_kopeks) / 100.0 FROM prices GROUP BY price_kind"
                ).fetchall()
            ],
            "by_system": conn.execute(
                """
                SELECT system, COUNT(DISTINCT price_lines.id),
                       SUM(CASE WHEN price_kind = 'material' THEN total_price_kopeks ELSE 0 END) / 100.0 AS material_total,
                       SUM(CASE WHEN price_kind = 'installation' THEN total_price_kopeks ELSE 0 END) / 100.0 AS installation_total
                FROM price_lines
                JOIN prices ON prices.line_id = price_lines.id
                GROUP BY system
                ORDER BY system
                """
            ).fetchall(),
        }

    print(json.dumps(checks, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
