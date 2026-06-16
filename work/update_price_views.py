import sqlite3


DB_PATH = r"C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite"


SQL = """
DROP VIEW IF EXISTS v_price_lines;
DROP VIEW IF EXISTS v_prices;

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


with sqlite3.connect(DB_PATH) as conn:
    conn.executescript(SQL)
    print(conn.execute("PRAGMA integrity_check").fetchone()[0])
