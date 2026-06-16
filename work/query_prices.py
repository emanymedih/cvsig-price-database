from __future__ import annotations

import argparse
import csv
import json
import sqlite3
import sys
from pathlib import Path


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

CONFIG_PATH = Path(__file__).with_name("db_config.json")


def load_config() -> dict:
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def connect() -> sqlite3.Connection:
    config = load_config()
    db_path = config["database_path"]
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def print_rows(rows: list[sqlite3.Row]) -> None:
    if not rows:
        print("No rows.")
        return
    writer = csv.writer(sys.stdout, lineterminator="\n")
    writer.writerow(rows[0].keys())
    for row in rows:
        writer.writerow([row[key] for key in row.keys()])


def cmd_summary(_: argparse.Namespace) -> None:
    with connect() as conn:
        summary = {
            "integrity": conn.execute("PRAGMA integrity_check").fetchone()[0],
            "line_count": conn.execute("SELECT COUNT(*) FROM price_lines").fetchone()[0],
            "price_count": conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0],
            "totals_by_kind": [
                dict(row)
                for row in conn.execute(
                    """
                    SELECT
                        price_kind,
                        COUNT(*) AS records,
                        ROUND(SUM(total_price_kopeks) / 100.0, 2) AS total_rub
                    FROM prices
                    GROUP BY price_kind
                    ORDER BY price_kind
                    """
                )
            ],
            "lines_by_system": [
                dict(row)
                for row in conn.execute(
                    """
                    SELECT
                        system,
                        COUNT(*) AS lines,
                        ROUND(SUM(material_total_price_rub), 2) AS material_total_rub,
                        ROUND(SUM(installation_total_price_rub), 2) AS installation_total_rub
                    FROM v_price_lines
                    GROUP BY system
                    ORDER BY system
                    """
                )
            ],
        }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


def cmd_search(args: argparse.Namespace) -> None:
    term = f"%{args.term}%"
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT
                system,
                position,
                name,
                type_code_article,
                manufacturer,
                unit,
                quantity,
                material_unit_price_rub,
                material_vat_note,
                installation_unit_price_rub,
                installation_vat_note
            FROM v_price_lines
            WHERE
                system LIKE ?
                OR name LIKE ?
                OR type_code_article LIKE ?
                OR manufacturer LIKE ?
            ORDER BY system, CAST(position AS REAL), excel_row
            LIMIT ?
            """,
            (term, term, term, term, args.limit),
        ).fetchall()
    print_rows(rows)


def cmd_system(args: argparse.Namespace) -> None:
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT
                position,
                name,
                type_code_article,
                manufacturer,
                unit,
                quantity,
                material_unit_price_rub,
                material_total_price_rub,
                installation_unit_price_rub,
                installation_total_price_rub
            FROM v_price_lines
            WHERE system = ?
            ORDER BY CAST(position AS REAL), excel_row
            """,
            (args.system,),
        ).fetchall()
    print_rows(rows)


def cmd_sql(args: argparse.Namespace) -> None:
    sql = args.query.strip()
    if not sql.lower().startswith("select"):
        raise SystemExit("Only SELECT queries are allowed through this helper.")
    with connect() as conn:
        rows = conn.execute(sql).fetchmany(args.limit)
    print_rows(rows)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Quick access to the CVSIG prices database.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    summary = subparsers.add_parser("summary", help="Show database summary.")
    summary.set_defaults(func=cmd_summary)

    search = subparsers.add_parser("search", help="Search by system, name, article, or manufacturer.")
    search.add_argument("term")
    search.add_argument("--limit", type=int, default=50)
    search.set_defaults(func=cmd_search)

    system = subparsers.add_parser("system", help="List all rows for a system, e.g. АПС.")
    system.add_argument("system")
    system.set_defaults(func=cmd_system)

    sql = subparsers.add_parser("sql", help="Run a read-only SELECT query.")
    sql.add_argument("query")
    sql.add_argument("--limit", type=int, default=500)
    sql.set_defaults(func=cmd_sql)

    return parser


def main() -> None:
    args = build_parser().parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
