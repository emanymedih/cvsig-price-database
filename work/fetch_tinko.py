"""Download TINKO "Цены и остатки" and keep only the rows our estimates need.

Runs in GitHub Actions (.github/workflows/tinko-prices.yml); needs openpyxl.

The full list (~85k items) is NOT stored in this public repository: it is the
supplier's database (ГК РФ ст. 1334). Only rows matching suppliers/tinko/watch.csv
terms or codes pinned in mapping/items.csv are written to suppliers/tinko/prices.csv.

Local run on an already downloaded file:  python3 work/fetch_tinko.py --source price.xlsx
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import tempfile
import time
import urllib.request
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pricelib import key  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "suppliers" / "tinko"
WATCH = OUT_DIR / "watch.csv"
MAPPING = ROOT / "mapping" / "items.csv"
URL = "https://www.tinko.ru/upload/exchange/public/price-stock/price.xlsx"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36"
COLUMNS = [  # (output column, source header)
    ("code", "Код"),
    ("article", "Артикул"),
    ("article_extra", "Артикул доп"),
    ("name", "Наименование"),
    ("trade_name", "Торговое наименование"),
    ("unit", "Ед"),
    ("availability", "Срок поставки"),
    ("retail_rub", "Цена розница"),
    ("currency", "Валюта"),
    ("manufacturer", "Производитель"),
    ("category", "Функционал"),
]
MAX_PER_TERM = 40        # a vague term must not pull a large part of the supplier database
MIN_SOURCE_ROWS = 50000  # sanity check: never overwrite good data with a broken download


def download(dest: Path) -> None:
    last = None
    for attempt in range(3):
        try:
            req = urllib.request.Request(URL, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=180) as resp:
                dest.write_bytes(resp.read())
            return
        except Exception as exc:  # network hiccup: retry
            last = exc
            time.sleep(10 * (attempt + 1))
    raise SystemExit(f"Cannot download {URL}: {last}")


def text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def money(v) -> str:
    if v is None or v == "":
        return ""
    d = Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(d.normalize(), "f")


def read_source(path: Path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    title, header, rows = "", None, []
    for row in ws.iter_rows(values_only=True):
        if header is None:
            cells = [text(c) for c in row]
            for c in cells:
                if "Цены и остатки" in c:
                    title = c
            if "Код" in cells and "Цена розница" in cells:
                header = {name: i for i, name in enumerate(cells) if name}
            continue
        if row[header["Код"]] is None:
            continue
        rows.append({out: row[header[src]] if header[src] < len(row) else None for out, src in COLUMNS})
    if header is None:
        raise SystemExit("Header row (Код / Цена розница) not found")
    missing = [src for _, src in COLUMNS if src not in header]
    if missing:
        raise SystemExit(f"Columns missing in the source: {missing}")
    return title, rows


def source_timestamp(title: str) -> str:
    m = re.search(r"на\s+(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})", title)
    if not m:
        raise SystemExit(f"Cannot read export time from title: {title!r}")
    d, mo, y, hh, mm, ss = m.groups()
    return f"{y}-{mo}-{d}T{hh}:{mm}:{ss}+03:00"


def read_watch() -> tuple[list[str], str]:
    raw = WATCH.read_bytes() if WATCH.exists() else b""
    terms = []
    if raw:
        for r in csv.DictReader(raw.decode("utf-8").splitlines()):
            t = (r.get("term") or "").strip()
            if t and len(key(t)) >= 3:
                terms.append(t)
    return terms, hashlib.sha256(raw).hexdigest()


def read_pinned_codes() -> set[str]:
    if not MAPPING.exists():
        return set()
    with MAPPING.open(encoding="utf-8") as f:
        return {(r.get("tinko_code") or "").strip() for r in csv.DictReader(f)} - {""}


def select(rows, terms, pinned):
    term_keys = {}
    for t in terms:
        term_keys.setdefault(key(t), t)
    hits = {tk: [] for tk in term_keys}
    keep = {}
    for r in rows:
        code = text(r["code"])
        fields = (key(r["trade_name"]), key(r["article"]), key(r["article_extra"]))
        hay = "|".join(fields) + "|" + key(r["name"])
        for tk in term_keys:
            if tk in hay:
                hits[tk].append((tk in fields, code, r))
        if code in pinned:
            keep[code] = r
    capped = []
    for tk, lst in hits.items():
        lst.sort(key=lambda x: (not x[0], int(x[1]) if x[1].isdigit() else 0))  # exact field matches first
        if len(lst) > MAX_PER_TERM:
            capped.append(term_keys[tk])
        for _, code, r in lst[:MAX_PER_TERM]:
            keep[code] = r
    return keep, sorted(capped)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--source", type=Path, help="use a local price.xlsx instead of downloading")
    args = ap.parse_args()

    with tempfile.TemporaryDirectory() as tmp:
        src = args.source or Path(tmp) / "price.xlsx"
        if not args.source:
            download(src)
        title, rows = read_source(src)
    if len(rows) < MIN_SOURCE_ROWS:
        raise SystemExit(f"Only {len(rows)} rows in the source: refusing to overwrite data")

    terms, watch_sha = read_watch()
    keep, capped = select(rows, terms, read_pinned_codes())

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_rows = sorted(keep.items(), key=lambda kv: (0, int(kv[0])) if kv[0].isdigit() else (1, kv[0]))
    with (OUT_DIR / "prices.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow([c for c, _ in COLUMNS])
        for code, r in out_rows:
            w.writerow([code if c == "code" else money(r[c]) if c == "retail_rub" else text(r[c]) for c, _ in COLUMNS])

    meta = {
        "source": URL,
        "source_title": title,
        "source_timestamp": source_timestamp(title),
        "source_rows": len(rows),
        "selected_rows": len(out_rows),
        "watch_terms": len(terms),
        "watch_sha256": watch_sha,
        "capped_terms": capped,
        "price_column": "Цена розница",
        "vat": "Розничная цена Тинко включает НДС (подтверждено пользователем 18.10.2026)",
    }
    (OUT_DIR / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
