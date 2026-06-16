from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
import sqlite3
import sys
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path


sys.path.insert(0, str(Path(__file__).with_name("vendor")))
import xlrd  # noqa: E402


SOURCE = Path(r"C:\Users\admla\Downloads\Дагестан 23.04.xls")
DB_PATH = Path(r"C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite")
SYSTEM_PREFIX = "Дагестан"
PRICE_DATE = "2026-04-23"
VAT_NOTE = "по исходной таблице: руб. с НДС"


def money_to_kopeks(value):
    if value is None or value == "":
        return None
    try:
        amount = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None
    return int(amount * 100)


def num_or_none(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def text_or_none(value):
    if value is None:
        return None
    value = str(value).strip()
    return value or None


def clean_position(value, previous_position, child_index):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            return str(int(value)), str(int(value)), 0
        return str(value), str(value), 0

    if previous_position:
        next_child = child_index + 1
        return f"{previous_position}.{next_child}", previous_position, next_child
    return f"row-{child_index}", previous_position, child_index + 1


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def infer_nonprice_note(*values):
    notes = []
    for value in values:
        if isinstance(value, str) and value.strip() and not re.search(r"\d", value):
            notes.append(value.strip())
    return "; ".join(notes) if notes else None


def iter_rows():
    workbook = xlrd.open_workbook(str(SOURCE))
    for sheet in workbook.sheets():
        previous_position = None
        child_index = 0
        for row_index in range(sheet.nrows):
            if row_index < 3:
                continue

            row = sheet.row_values(row_index)
            while len(row) < 11:
                row.append("")

            name = text_or_none(row[1])
            if not name:
                continue

            material_unit = money_to_kopeks(row[6])
            material_total = money_to_kopeks(row[7])
            installation_unit = money_to_kopeks(row[8])
            installation_total = money_to_kopeks(row[9])
            if all(
                value is None
                for value in (material_unit, material_total, installation_unit, installation_total)
            ):
                continue

            position, previous_position, child_index = clean_position(
                row[0], previous_position, child_index
            )
            if child_index == 0:
                previous_position = position

            source_note = text_or_none(row[10])
            nonprice_note = infer_nonprice_note(row[6], row[7], row[8], row[9])
            notes = [part for part in (source_note, nonprice_note, sheet.name) if part]

            yield {
                "system": f"{SYSTEM_PREFIX} / {sheet.name}",
                "excel_row": row_index + 1,
                "position": position,
                "name": name,
                "type_code_article": text_or_none(row[2]),
                "manufacturer": text_or_none(row[3]),
                "unit": text_or_none(row[4]),
                "quantity": num_or_none(row[5]),
                "material_unit_kopeks": material_unit,
                "material_total_kopeks": material_total,
                "installation_unit_kopeks": installation_unit,
                "installation_total_kopeks": installation_total,
                "note": "; ".join(notes) if notes else None,
            }


def main():
    rows = list(iter_rows())
    source_stat = SOURCE.stat()
    source_hash = file_sha256(SOURCE)
    headers = {
        "columns": [
            "поз.",
            "Наименование",
            "Тип / код / артикул",
            "Изготовитель",
            "Ед.",
            "Кол-во",
            "Стоимость материала: ед. расценка, руб. с НДС",
            "Стоимость материала: итого, руб. с НДС",
            "Стоимость монтажа: ед. расценка, руб. с НДС",
            "Стоимость монтажа: итого, руб. с НДС",
            "Примечание",
        ],
        "system_prefix": SYSTEM_PREFIX,
        "position_note": "Rows without explicit position are stored as child rows under the previous numbered position.",
    }

    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        previous = conn.execute(
            "SELECT id FROM price_snapshots WHERE source_file = ? OR source_sha256 = ?",
            (str(SOURCE), source_hash),
        ).fetchall()
        for (snapshot_id,) in previous:
            conn.execute("DELETE FROM price_snapshots WHERE id = ?", (snapshot_id,))

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
                source_hash,
                dt.datetime.fromtimestamp(source_stat.st_mtime).isoformat(timespec="seconds"),
                PRICE_DATE,
                dt.datetime.now().isoformat(timespec="seconds"),
                json.dumps(headers, ensure_ascii=False),
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
                    row["position"],
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
                    INSERT INTO prices (
                        line_id,
                        price_kind,
                        unit_price_kopeks,
                        total_price_kopeks,
                        vat_note
                    )
                    VALUES (?, 'material', ?, ?, ?)
                    """,
                    (
                        line_id,
                        row["material_unit_kopeks"],
                        row["material_total_kopeks"],
                        VAT_NOTE,
                    ),
                )

            if row["installation_unit_kopeks"] is not None or row["installation_total_kopeks"] is not None:
                conn.execute(
                    """
                    INSERT INTO prices (
                        line_id,
                        price_kind,
                        unit_price_kopeks,
                        total_price_kopeks,
                        vat_note
                    )
                    VALUES (?, 'installation', ?, ?, ?)
                    """,
                    (
                        line_id,
                        row["installation_unit_kopeks"],
                        row["installation_total_kopeks"],
                        VAT_NOTE,
                    ),
                )

        result = {
            "inserted_lines": conn.execute(
                "SELECT COUNT(*) FROM price_lines WHERE snapshot_id = ?", (snapshot_id,)
            ).fetchone()[0],
            "inserted_prices": conn.execute(
                """
                SELECT price_kind, COUNT(*), ROUND(SUM(total_price_kopeks) / 100.0, 2)
                FROM prices
                JOIN price_lines ON price_lines.id = prices.line_id
                WHERE snapshot_id = ?
                GROUP BY price_kind
                ORDER BY price_kind
                """,
                (snapshot_id,),
            ).fetchall(),
            "lines_by_system": conn.execute(
                """
                SELECT system, COUNT(*)
                FROM price_lines
                WHERE snapshot_id = ?
                GROUP BY system
                ORDER BY system
                """,
                (snapshot_id,),
            ).fetchall(),
            "deleted_previous_snapshots": len(previous),
        }

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
