import json
import sqlite3


DB_PATH = r"C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite"


with sqlite3.connect(DB_PATH) as conn:
    summary = {
        "integrity": conn.execute("PRAGMA integrity_check").fetchone()[0],
        "objects": conn.execute(
            "SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view') ORDER BY type, name"
        ).fetchall(),
        "line_count": conn.execute("SELECT COUNT(*) FROM price_lines").fetchone()[0],
        "price_count": conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0],
        "totals": conn.execute(
            """
            SELECT price_kind, COUNT(*), ROUND(SUM(total_price_kopeks) / 100.0, 2)
            FROM prices
            GROUP BY price_kind
            ORDER BY price_kind
            """
        ).fetchall(),
        "systems": conn.execute(
            "SELECT system, COUNT(*) FROM price_lines GROUP BY system ORDER BY system"
        ).fetchall(),
        "sample": conn.execute(
            """
            SELECT system, position, name, material_unit_price_rub, installation_unit_price_rub
            FROM v_price_lines
            ORDER BY system, CAST(position AS REAL)
            LIMIT 3
            """
        ).fetchall(),
    }

print(json.dumps(summary, ensure_ascii=True, indent=2))
