from __future__ import annotations

import csv
import datetime as dt
import hashlib
import json
import sqlite3
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path


SOURCE = Path(r"C:\Users\admla\Documents\Codex\2026-04-24\excel\reliable_price_database_v2.csv")
DB_PATH = Path(r"C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite")
SYSTEM = "Надежная база 2026-04-24"
PRICE_DATE = "2026-04-24"
VAT_NOTE = "по надежной базе 2026-04-24: руб. с НДС"


def money_to_kopeks(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        amount = Decimal(str(value).strip()).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None
    return int(amount * 100)


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


def build_note(row: dict, csv_row: int) -> str:
    parts = [
        f"status={row.get('auto_fill_status', '')}",
        f"review_flags={row.get('review_flags', '')}" if row.get("review_flags") else "",
        f"sources_count={row.get('sources_count', '')}" if row.get("sources_count") else "",
        f"occurrences_count={row.get('occurrences_count', '')}" if row.get("occurrences_count") else "",
        f"best_confidence={row.get('best_row_confidence', '')}" if row.get("best_row_confidence") else "",
        f"source={row.get('best_source_file', '')}" if row.get("best_source_file") else "",
        f"sheet={row.get('best_source_sheet', '')}" if row.get("best_source_sheet") else "",
        f"source_row={row.get('best_source_row', '')}" if row.get("best_source_row") else "",
        f"csv_row={csv_row}",
    ]
    return "; ".join(part for part in parts if part)


def main():
    source_hash = file_sha256(SOURCE)
    source_stat = SOURCE.stat()
    rows = []
    with SOURCE.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for csv_row, row in enumerate(reader, start=2):
            material = money_to_kopeks(row.get("recommended_material_price"))
            installation = money_to_kopeks(row.get("recommended_install_price"))
            if material is None and installation is None:
                continue
            rows.append((csv_row, row, material, installation))

    headers = {
        "source_kind": "reliable aggregate csv",
        "csv_columns": list(rows[0][1].keys()) if rows else [],
        "rule": "User marked this folder as a reliable price source on 2026-06-08.",
        "price_semantics": "unit prices with VAT",
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

        for csv_row, row, material, installation in rows:
            status = text_or_none(row.get("auto_fill_status"))
            system = f"{SYSTEM} / {status}" if status else SYSTEM
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
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL, ?)
                """,
                (
                    snapshot_id,
                    system,
                    csv_row,
                    str(csv_row - 1),
                    text_or_none(row.get("name")),
                    text_or_none(row.get("article")),
                    text_or_none(row.get("manufacturer")),
                    text_or_none(row.get("unit")),
                    build_note(row, csv_row),
                ),
            )
            line_id = cursor.lastrowid

            if material is not None:
                conn.execute(
                    """
                    INSERT INTO prices (
                        line_id,
                        price_kind,
                        unit_price_kopeks,
                        total_price_kopeks,
                        vat_note
                    )
                    VALUES (?, 'material', ?, NULL, ?)
                    """,
                    (line_id, material, VAT_NOTE),
                )
            if installation is not None:
                conn.execute(
                    """
                    INSERT INTO prices (
                        line_id,
                        price_kind,
                        unit_price_kopeks,
                        total_price_kopeks,
                        vat_note
                    )
                    VALUES (?, 'installation', ?, NULL, ?)
                    """,
                    (line_id, installation, VAT_NOTE),
                )

        result = {
            "inserted_lines": conn.execute(
                "SELECT COUNT(*) FROM price_lines WHERE snapshot_id = ?", (snapshot_id,)
            ).fetchone()[0],
            "inserted_prices": conn.execute(
                """
                SELECT price_kind, COUNT(*)
                FROM prices
                JOIN price_lines ON price_lines.id = prices.line_id
                WHERE snapshot_id = ?
                GROUP BY price_kind
                ORDER BY price_kind
                """,
                (snapshot_id,),
            ).fetchall(),
            "lines_by_status": conn.execute(
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
