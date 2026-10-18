"""Parse a supplier invoice PDF (счёт) into suppliers/invoices/<номер>.csv.

    python3 work/parse_invoice.py СЧЁТ.pdf            # writes suppliers/invoices/<номер>.csv

Only the goods lines are stored (no bank details or buyer data). Prices are as in the invoice
(с НДС). The total of the parsed lines is checked against "Итого" in the invoice.
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from decimal import Decimal
from pathlib import Path

import pdfplumber

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "suppliers" / "invoices"
MONEY = re.compile(r"^\d[\d\s]*,\d{2}$")
QTY = re.compile(r"^([\d\s]+(?:,\d+)?)\s*(\S+)$")
FIELDS = ["invoice", "date", "supplier", "line", "name", "code", "maker", "unit", "qty", "price_rub", "sum_rub"]


def money(s: str) -> Decimal:
    return Decimal(re.sub(r"\s", "", s).replace(",", "."))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("pdf", type=Path)
    args = ap.parse_args()
    with pdfplumber.open(args.pdf) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        rows = [r for p in pdf.pages for t in p.extract_tables() for r in t]
    m = re.search(r"Сч[её]т\s*№\s*(\S+)\s+от\s+(\d{2})\.(\d{2})\.(\d{4})", text)
    if not m:
        raise SystemExit("Не найден номер и дата счёта")
    number, date = m.group(1), f"{m.group(4)}-{m.group(3)}-{m.group(2)}"
    sm = re.search(r"Продавец\s*\n\s*(.+?)(?:\s+\d{20}|\n)", text)
    supplier = (sm.group(1).strip() if sm else "").replace('"', "")
    total = re.search(r"Итого,\s*руб:?\s*([\d\s]+,\d{2})", text)

    lines = []
    for r in rows:
        cells = [(c or "").strip() for c in r]
        price_i = next((i for i, c in enumerate(cells) if MONEY.match(c)), None)
        if price_i is None or price_i + 2 >= len(cells) or not MONEY.match(cells[price_i + 2]):
            continue
        q = QTY.match(cells[price_i + 1])
        if not q:
            continue
        name = next((c for c in cells[:price_i] if len(c) > 3 and not c.startswith("1 р.д")), "")
        name = re.sub(r"\s*\n\s*", " ", name)
        mk = re.search(r"\(([^()]*(?:\([^()]*\))?[^()]*)\)\s*$", name)
        maker = mk.group(1) if mk else ""
        body = name[:mk.start()].strip() if mk else name
        cm = re.search(r"\s(\S*\d\S*)$", body)  # supplier/manufacturer code at the end of the name
        code = cm.group(1) if cm and len(re.sub(r"\D", "", cm.group(1))) >= 4 else ""
        lines.append({
            "invoice": number, "date": date, "supplier": supplier, "line": len(lines) + 1, "name": body,
            "code": code, "maker": maker, "unit": q.group(2).rstrip("."),
            "qty": str(money(q.group(1) + ",0") if "," not in q.group(1) else money(q.group(1))).rstrip("0").rstrip("."),
            "price_rub": str(money(cells[price_i])), "sum_rub": str(money(cells[price_i + 2])),
        })
    got = sum(Decimal(x["sum_rub"]) for x in lines)
    print(f"Счёт {number} от {date}, {supplier}: строк {len(lines)}, сумма {got}")
    if total and money(total.group(1)) != got:
        raise SystemExit(f"Сумма строк {got} не равна итогу счёта {money(total.group(1))}: разбор неполный")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / f"{number}.csv"
    with out.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader()
        w.writerows(lines)
    print("->", out.relative_to(ROOT))


if __name__ == "__main__":
    sys.exit(main())
