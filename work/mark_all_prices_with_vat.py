import sqlite3


DB_PATH = r"C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite"
VAT_NOTE = "по уточнению пользователя: руб. с НДС; исходный заголовок мог быть неверным"


with sqlite3.connect(DB_PATH) as conn:
    conn.execute("UPDATE prices SET vat_note = ?", (VAT_NOTE,))
    print(conn.execute("PRAGMA integrity_check").fetchone()[0])
    print(conn.execute("SELECT COUNT(*) FROM prices WHERE vat_note = ?", (VAT_NOTE,)).fetchone()[0])
