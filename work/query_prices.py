"""Read-only helper for the CVSIG prices database.

Usage (from any directory):
    python3 work/query_prices.py summary
    python3 work/query_prices.py search "кабель кпсвэв" [--limit 30]
    python3 work/query_prices.py system "АПС" [--limit 200]
    python3 work/query_prices.py sql "SELECT ... FROM v_price_lines ..." [--limit 50]

The database path comes from work/db_config.json ("database_path", relative to
that file) and can be overridden with the CVSIG_DB environment variable.
The database is always opened read-only.
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import sqlite3
import sys
from pathlib import Path


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

CONFIG_PATH = Path(__file__).with_name("db_config.json")


def db_path() -> Path:
    override = os.environ.get("CVSIG_DB")
    if override:
        path = Path(override)
    else:
        config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        path = Path(config["database_path"])
        if not path.is_absolute():
            path = CONFIG_PATH.parent / path
    path = path.resolve()
    if not path.is_file():
        raise SystemExit(f"Database not found: {path}")
    return path


def connect() -> sqlite3.Connection:
    # mode=ro: never creates an empty file and never writes to the database.
    conn = sqlite3.connect(db_path().as_uri() + "?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    # SQLite LIKE/lower() are case-insensitive for ASCII only; this handles Cyrillic.
    conn.create_function("ulower", 1, lambda s: s.lower() if isinstance(s, str) else s, deterministic=True)
    return conn


def print_rows(rows: list[sqlite3.Row], limit: int | None = None, more: bool = False) -> None:
    if not rows:
        print("No rows.")
        return
    writer = csv.writer(sys.stdout, lineterminator="\n")
    writer.writerow(rows[0].keys())
    for row in rows:
        writer.writerow([row[key] for key in row.keys()])
    if more:
        print(f"... truncated at {limit} rows; narrow the query or raise --limit.")


def fetch_limited(conn: sqlite3.Connection, sql: str, params: tuple, limit: int) -> tuple[list[sqlite3.Row], bool]:
    rows = conn.execute(sql, params).fetchmany(limit + 1)
    return rows[:limit], len(rows) > limit


def cmd_summary(_: argparse.Namespace) -> None:
    conn = connect()
    summary = {
        "database": str(db_path()),
        "line_count": conn.execute("SELECT COUNT(*) FROM price_lines").fetchone()[0],
        "price_count": conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0],
        "snapshots": [
            dict(row)
            for row in conn.execute(
                """
                SELECT s.id, s.price_date, COUNT(l.id) AS lines,
                       replace(s.source_file, rtrim(s.source_file, replace(s.source_file, '\\', '')), '') AS source
                FROM price_snapshots s LEFT JOIN price_lines l ON l.snapshot_id = s.id
                GROUP BY s.id ORDER BY s.id
                """
            )
        ],
        "totals_by_kind": [
            dict(row)
            for row in conn.execute(
                """
                SELECT price_kind, COUNT(*) AS records,
                       ROUND(SUM(total_price_kopeks) / 100.0, 2) AS total_rub
                FROM prices GROUP BY price_kind ORDER BY price_kind
                """
            )
        ],
        "lines_by_system": {
            row["system"]: row["lines"]
            for row in conn.execute(
                "SELECT system, COUNT(*) AS lines FROM v_price_lines GROUP BY system ORDER BY system"
            )
        },
    }
    conn.close()
    print(json.dumps(summary, ensure_ascii=False, indent=1))


def cmd_search(args: argparse.Namespace) -> None:
    # Every word must occur in name, article, manufacturer or system (any order, any case).
    words = [w for w in args.term.lower().split() if w]
    if not words:
        raise SystemExit("Empty search term.")
    haystack = "ulower(coalesce(name,'') || ' ' || coalesce(type_code_article,'') || ' ' || coalesce(manufacturer,'') || ' ' || system)"
    where = " AND ".join(f"{haystack} LIKE ?" for _ in words)
    sql = f"""
        SELECT price_date, system, name, type_code_article, manufacturer, unit,
               material_unit_price_rub AS mat_rub, installation_unit_price_rub AS inst_rub
        FROM v_price_lines
        WHERE {where}
        ORDER BY price_date DESC, system, excel_row
    """
    conn = connect()
    rows, more = fetch_limited(conn, sql, tuple(f"%{w}%" for w in words), args.limit)
    conn.close()
    print_rows(rows, args.limit, more)


def cmd_system(args: argparse.Namespace) -> None:
    sql = """
        SELECT position, name, type_code_article, manufacturer, unit, quantity,
               material_unit_price_rub AS mat_rub, material_total_price_rub AS mat_total,
               installation_unit_price_rub AS inst_rub, installation_total_price_rub AS inst_total
        FROM v_price_lines
        WHERE system = ?
        ORDER BY CAST(position AS REAL), excel_row
    """
    conn = connect()
    rows, more = fetch_limited(conn, sql, (args.system,), args.limit)
    conn.close()
    print_rows(rows, args.limit, more)


def cmd_sql(args: argparse.Namespace) -> None:
    sql = args.query.strip()
    if not sql.lower().startswith(("select", "with")):
        raise SystemExit("Only SELECT / WITH queries are allowed.")
    conn = connect()  # read-only connection, so writes fail anyway
    rows, more = fetch_limited(conn, sql, (), args.limit)
    conn.close()
    print_rows(rows, args.limit, more)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Read-only access to the CVSIG prices database.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    summary = subparsers.add_parser("summary", help="Snapshots, counts and lines per system.")
    summary.set_defaults(func=cmd_summary)

    search = subparsers.add_parser("search", help="All words must match name/article/manufacturer/system; newest first.")
    search.add_argument("term")
    search.add_argument("--limit", type=int, default=30)
    search.set_defaults(func=cmd_search)

    system = subparsers.add_parser("system", help="Rows of one exact system, e.g. АПС.")
    system.add_argument("system")
    system.add_argument("--limit", type=int, default=200)
    system.set_defaults(func=cmd_system)

    sql = subparsers.add_parser("sql", help="Read-only SELECT/WITH query; ulower(x) lowercases Cyrillic.")
    sql.add_argument("query")
    sql.add_argument("--limit", type=int, default=50)
    sql.set_defaults(func=cmd_sql)

    return parser


def main() -> None:
    args = build_parser().parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
