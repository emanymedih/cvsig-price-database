"""Read retail prices from product pages listed in suppliers/pages/urls.txt.

Runs in GitHub Actions (.github/workflows/supplier-pages.yml); standard library only.
Parses schema.org markup (itemprop="price" / JSON-LD "price"), as on satro-paladin.com.
One URL per line in urls.txt; text after "|" is a comment; "#" starts a comment line.
"""
from __future__ import annotations

import csv
import datetime as dt
import html
import re
import time
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIR = ROOT / "suppliers" / "pages"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36"
FIELDS = ["url", "supplier", "name", "price_rub", "currency", "price_label", "availability", "checked_date", "error"]


def read_urls() -> list[str]:
    path = DIR / "urls.txt"
    urls = []
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.split("|", 1)[0].strip()
            if line and not line.startswith("#") and line not in urls:
                urls.append(line)
    return urls


def fetch(url: str) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "ru"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        charset = resp.headers.get_content_charset() or "utf-8"
        return resp.read().decode(charset, errors="replace")


def first(patterns, page: str) -> str:
    for p in patterns:
        m = re.search(p, page, re.S | re.I)
        if m:
            return html.unescape(m.group(1)).strip()
    return ""


def parse(page: str) -> dict:
    price = first([r'itemprop="price"[^>]*content="([\d\s.,]+)"', r'"price"\s*:\s*"?([\d.]+)'], page)
    price = price.replace(" ", "").replace(",", ".")
    name = first([r'<h1[^>]*>(.*?)</h1>', r'<meta property="og:title" content="([^"]+)"', r"<title>(.*?)</title>"], page)
    name = re.sub(r"<[^>]+>", " ", name)
    return {
        "name": re.sub(r"\s+", " ", name).strip(),
        "price_rub": price,
        "currency": first([r'itemprop="priceCurrency"[^>]*content="([A-Z]{3})"', r'"priceCurrency"\s*:\s*"([A-Z]{3})"'], page),
        "price_label": "Розничная цена" if "Розничная цена" in page else "",
        "availability": first([r'itemprop="availability"[^>]*(?:content|href)="[^"]*/(\w+)"', r'"availability"\s*:\s*"[^"]*/(\w+)"'], page),
    }


def main() -> None:
    today = dt.date.today().isoformat()
    rows = []
    for i, url in enumerate(read_urls()):
        if i:
            time.sleep(2)  # be polite
        row = {"url": url, "supplier": urllib.parse.urlparse(url).netloc.removeprefix("www."), "checked_date": today}
        try:
            row.update(parse(fetch(url)))
            if not row.get("price_rub"):
                row["error"] = "цена на странице не найдена"
        except Exception as exc:
            row["error"] = f"{type(exc).__name__}: {exc}"[:200]
        rows.append(row)
    DIR.mkdir(parents=True, exist_ok=True)
    with (DIR / "prices.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader()
        for row in sorted(rows, key=lambda r: r["url"]):
            w.writerow({k: row.get(k, "") for k in FIELDS})
    for row in rows:
        print(row["url"], row.get("price_rub"), row.get("error", ""))


if __name__ == "__main__":
    main()
