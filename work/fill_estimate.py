"""Fill an estimate (смета / КП) with current retail material prices and installation prices.

  python3 work/fill_estimate.py plan  ESTIMATE.xlsx                 # add its items to the TINKO watch list
  python3 work/fill_estimate.py fill  ESTIMATE.xlsx [-o OUT.xlsx]   # write prices + sheet "Источники"
  python3 work/fill_estimate.py watch "С2000-КДЛ" "PR.012031"       # add arbitrary terms to the watch list

Rules agreed with the user (2026-10-18):
  * material = TINKO retail price incl. VAT (suppliers/tinko/prices.csv, refreshed by GitHub Actions);
    if TINKO has no such item: a product-page price (suppliers/pages/prices.csv) linked in mapping/items.csv;
  * installation = the freshest price in the SQLite base; several prices on that date -> the higher one (flagged);
  * nothing is guessed: without a confident match the cell stays empty and is highlighted.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import json
import re
import sqlite3
import statistics
import sys
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pricelib import (  # noqa: E402
    contains_seq, distinctive, key, paren_keys, same_maker, strip_paren, strip_variant, tokens, trailing_variants,
    unit_factor,
)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
DB = ROOT / "outputs" / "cvsig_prices_2026-06-02.sqlite"
TINKO_DIR = ROOT / "suppliers" / "tinko"
WATCH = TINKO_DIR / "watch.csv"
PAGES = ROOT / "suppliers" / "pages" / "prices.csv"
INVOICES = ROOT / "suppliers" / "invoices"
MAPPING = ROOT / "mapping" / "items.csv"
MSK = dt.timezone(dt.timedelta(hours=3))
YELLOW = PatternFill("solid", fgColor="FFC7CE")  # no price found (light red, Excel "Bad")
ORANGE = PatternFill("solid", fgColor="FFEB9C")  # price written, needs a check (light yellow, Excel "Neutral")
HEAD_FILL = PatternFill("solid", fgColor="FCE8DA")  # header fill of the estimate template
THIN = Side(style="thin", color="000000")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
# Excel "Финансовый" (accounting) format without a currency sign (user request 18.10.2026); dash for zero.
FIN = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'
FONT_NAME = "Times New Roman"
TINKO_URL = "https://www.tinko.ru/catalog/product/{:06d}/"
OK, CHECK, AMBIG, NONE, KEPT = "ОК", "Проверить", "Неоднозначно", "Не найдено", "Из файла"


# ---------------------------------------------------------------- estimate

@dataclass
class Item:
    row: int
    pos: str
    name: str
    article_cell: str
    article: str
    code: str
    maker: str
    unit: str
    qty: float

    @property
    def text(self) -> str:
        return f"{self.name} {self.article_cell}"


def split_article(cell) -> tuple[str, str]:
    """'КПСнг(А)-FRHF 1×2×0,75; код УТ0005907' -> ('КПСнг(А)-FRHF 1×2×0,75', 'УТ0005907')."""
    cell = str(cell or "").strip()
    main = cell.split(";")[0].strip()
    m = re.search(r"код\s*[:№]?\s*([^\s;,]*\d[^\s;,]*)", cell, re.I)
    return main, (m.group(1) if m else "")


def detect_layout(ws) -> dict:
    header = None
    for r in range(1, 21):
        vals = {c: str(ws.cell(r, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)}
        if any("наименование" in v for v in vals.values()) and any(v.startswith("кол") for v in vals.values()):
            header = r
            break
    if header is None:
        raise SystemExit("Не найдена строка заголовков (Наименование / Кол-во)")

    def find(pred):
        return next((c for c, v in vals.items() if v and pred(v)), None)

    sub = header + 1
    sub_vals = {c: str(ws.cell(sub, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)}

    def unit_price_col(block):
        if block is None:
            return None
        for c in (block, block + 1):
            if "за ед" in sub_vals.get(c, ""):
                return c
        return block

    lay = {
        "header": header,
        "first": header + (2 if any("за ед" in v for v in sub_vals.values()) else 1),
        "pos": find(lambda v: v.startswith("поз") or v.startswith("№")),
        "name": find(lambda v: "наименование" in v),
        "article": find(lambda v: "артикул" in v or "марка" in v or v.startswith("тип")),
        "maker": find(lambda v: "изготовитель" in v or "производитель" in v),
        "unit": find(lambda v: v.startswith("ед")),
        "qty": find(lambda v: v.startswith("кол")),
        "mat": unit_price_col(find(lambda v: "материал" in v or "оборудован" in v)),
        "inst": unit_price_col(find(lambda v: "монтаж" in v or "работ" in v)),
    }
    missing = [k for k in ("name", "article", "unit", "qty", "mat", "inst") if not lay[k]]
    if missing:
        raise SystemExit(f"Не найдены колонки: {missing}")
    return lay


def read_items(ws, lay) -> list[Item]:
    items = []
    for r in range(lay["first"], ws.max_row + 1):
        first = str(ws.cell(r, 1).value or "").strip().lower()
        if first.startswith("итого"):
            break
        name, qty = ws.cell(r, lay["name"]).value, ws.cell(r, lay["qty"]).value
        if not name or not isinstance(qty, (int, float)):
            continue
        cell = ws.cell(r, lay["article"]).value
        art, code = split_article(cell)
        items.append(Item(
            row=r,
            pos=str(ws.cell(r, lay["pos"]).value or "").strip() if lay["pos"] else "",
            name=str(name).strip(),
            article_cell=str(cell or "").strip(),
            article=art,
            code=code,
            maker=str(ws.cell(r, lay["maker"]).value or "").strip() if lay["maker"] else "",
            unit=str(ws.cell(r, lay["unit"]).value or "").strip(),
            qty=float(qty),
        ))
    return items


def open_estimate(path: Path, sheet: str | None):
    wb = openpyxl.load_workbook(path)  # formulas are kept
    sheets = [wb[sheet]] if sheet else [s for s in wb.worksheets if s.title != "Источники"]
    for ws in sheets:
        try:
            lay = detect_layout(ws)
        except SystemExit:
            if sheet:
                raise
            continue
        return wb, ws, lay
    raise SystemExit("Ни на одном листе не найдена таблица сметы")


# ---------------------------------------------------------------- watch list

def watch_terms(it: Item) -> list[str]:
    terms = []
    if it.code and len(key(it.code)) >= 4:
        terms.append(it.code)
    if it.article and len(key(it.article)) >= 3 and (distinctive(tokens(it.article)) or not it.code):
        terms.append(it.article)
    sp = strip_paren(it.article)
    if sp != it.article and len(key(sp)) >= 4 and distinctive(tokens(sp)):
        terms.append(sp)
    return terms


def read_watch() -> list[dict]:
    if not WATCH.exists():
        return []
    with WATCH.open(encoding="utf-8") as f:
        return list(csv.DictReader(f))


def add_watch(terms: list[tuple[str, str]]) -> list[str]:
    rows = read_watch()
    have = {key(r["term"]) for r in rows}
    new = []
    for term, comment in terms:
        k = key(term)
        if len(k) >= 3 and k not in have:
            have.add(k)
            rows.append({"term": term, "comment": comment})
            new.append(term)
    if new:
        WATCH.parent.mkdir(parents=True, exist_ok=True)
        with WATCH.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["term", "comment"], lineterminator="\n")
            w.writeheader()
            w.writerows(rows)
    return new


# ---------------------------------------------------------------- data sources

def load_tinko():
    meta = json.loads((TINKO_DIR / "meta.json").read_text(encoding="utf-8"))
    rows = []
    with (TINKO_DIR / "prices.csv").open(encoding="utf-8") as f:
        for r in csv.DictReader(f):
            r["k_trade"], r["k_art"], r["k_art2"] = key(r["trade_name"]), key(r["article"]), key(r["article_extra"])
            r["k_trade_vars"] = {key(v) for v in trailing_variants(r["trade_name"])}
            r["k_name"] = key(r["name"])
            r["paren"] = paren_keys(r["name"]) | paren_keys(r["trade_name"])
            r["toks"] = tokens(r["trade_name"]) + ["|"] + tokens(r["name"])
            r["price"] = float(r["retail_rub"]) if r["retail_rub"] else None
            rows.append(r)
    return meta, rows


def check_fresh(meta: dict, max_age_h: float, allow_stale: bool) -> str:
    ts = dt.datetime.fromisoformat(meta["source_timestamp"])
    age = (dt.datetime.now(MSK) - ts).total_seconds() / 3600
    msg = f"Прайс Тинко: «{meta['source_title']}», возраст {age:.0f} ч"
    if age > max_age_h and not allow_stale:
        raise SystemExit(msg + f": старше {max_age_h:g} ч. Запустите workflow tinko-prices или добавьте --allow-stale")
    return msg


def load_invoices() -> list[dict]:
    """Goods lines of supplier invoices (work/parse_invoice.py); a newer invoice wins."""
    rows = []
    for path in sorted(INVOICES.glob("*.csv")):
        with path.open(encoding="utf-8") as f:
            for r in csv.DictReader(f):
                r["k_name"], r["toks"] = key(r["name"]), tokens(r["name"])
                r["k_code"] = key(r["code"])
                r["price"] = float(r["price_rub"])
                r["ref"] = f"{r['invoice']}:{r['line']}"
                rows.append(r)
    return rows


def invoice_candidates(it: Item, inv: list[dict], me: dict, auto: set | None = None):
    """Pinned line from the mapping; automatic matching only within the invoices named by --invoice,
    so an invoice for one object never changes prices of another estimate unnoticed."""
    if me.get("invoice"):
        return [r for r in inv if r["ref"] == me["invoice"]], "справочник"
    inv = [r for r in inv if r["invoice"] in (auto or set())]
    ck, ak, at = key(it.code), key(it.article), tokens(it.article)
    if len(ck) >= 4:
        c = [r for r in inv if ck == r["k_code"] or (ck in r["k_name"] and contains_seq(r["toks"], tokens(it.code)))]
        if c:
            return c, "по коду"
    if len(ak) >= 4 and distinctive(at):
        c = [r for r in inv if ak in r["k_name"] and contains_seq(r["toks"], at)]
        if c:
            return c, "по артикулу"
    return [], ""


def invoice_result(it: Item, inv: list[dict], me: dict, auto: set | None = None) -> Result | None:
    cands, how = invoice_candidates(it, inv, me, auto)
    if not cands:
        if me.get("invoice"):
            return Result(notes=[f"строки счёта {me['invoice']} из справочника нет в suppliers/invoices"])
        return None
    newest = max(c["date"] for c in cands)
    cands = [c for c in cands if c["date"] == newest]
    conv = []
    for r in cands:
        if me.get("material_factor"):
            f, note, sure = parse_factor(me["material_factor"]), f"коэффициент из справочника × {me['material_factor']}", True
        else:
            f, note, sure = unit_factor(r["unit"], it.unit, r["name"], it.text)
        if f is not None:
            conv.append((r, round(r["price"] * f, 2), note, sure))
    if not conv:
        return None
    prices = sorted({c[1] for c in conv})
    res = Result(how=how)
    if len(prices) > 1:
        res.status = AMBIG
        res.alts = "; ".join(f"[{c[0]['ref']}] {c[0]['name'][:60]} — {fmt(c[1])}" for c in conv[:6])
        res.notes.append("в счёте несколько подходящих строк с разной ценой — выберите (mapping/items.csv, invoice)")
        return res
    r, price, note, sure = conv[0]
    res.price = price
    res.status = OK if (how in ("справочник", "по коду") and sure) or me.get("material_status") == "confirmed" else CHECK
    if note:
        res.notes.append(note)
    if me.get("material_comment"):
        res.notes.append(me["material_comment"])
    res.src = {"supplier": f"{r['supplier']}, счёт {r['invoice']} от {r['date']}", "code": r["code"], "name": r["name"],
               "maker": r["maker"], "unit": r["unit"], "price": r["price"], "factor": None, "availability": "",
               "url": f"строка {r['line']}"}
    return res


def load_pages() -> dict:
    if not PAGES.exists():
        return {}
    with PAGES.open(encoding="utf-8") as f:
        return {r["url"]: r for r in csv.DictReader(f)}


def load_mapping() -> dict:
    out = {}
    if MAPPING.exists():
        with MAPPING.open(encoding="utf-8") as f:
            for r in csv.DictReader(f):
                out[(key(r.get("article")), key(r.get("code")))] = {k: (v or "").strip() for k, v in r.items()}
    return out


def mapping_for(it: Item, mapping: dict) -> dict:
    return mapping.get((key(it.article), key(it.code))) or mapping.get((key(it.article), "")) or {}


def load_base() -> list[dict]:
    conn = sqlite3.connect(DB.resolve().as_uri() + "?mode=ro", uri=True)
    q = """
        SELECT p.price_kind, p.unit_price_kopeks / 100.0, s.id, s.price_date, s.source_file,
               l.system, l.excel_row, l.unit, l.name, l.type_code_article
        FROM prices p
        JOIN price_lines l ON l.id = p.line_id
        JOIN price_snapshots s ON s.id = l.snapshot_id
        WHERE p.unit_price_kopeks > 0
    """
    rows = []
    for kind, price, sid, date, src, system, xrow, unit, name, art in conn.execute(q):
        art = (art or "").strip()
        k_art = key(art) if art not in ("", "—", "-") else ""
        rows.append({
            "kind": kind, "price": price, "sid": sid, "date": date, "source": re.split(r"[\\/]", src)[-1],
            "system": system, "excel_row": xrow, "unit": unit, "name": name or "", "article": art,
            "k_art": k_art, "k_text": key(name) + "|" + k_art, "toks": tokens(name) + ["|"] + tokens(art),
            "paren": paren_keys(name), "raw": f"{name} {art}", "tok_keys": {key(t) for t in tokens(f"{name} {art}")},
        })
    conn.close()
    return rows


# ---------------------------------------------------------------- matching

@dataclass
class Result:
    price: float | None = None
    status: str = NONE
    how: str = ""
    notes: list = field(default_factory=list)
    src: dict = field(default_factory=dict)
    alts: str = ""


def code_conflict(it: Item, text: str) -> bool:
    """Estimate says PR08.3659, the candidate mentions another PR08.xxxx -> different product."""
    m = re.match(r"^([A-Za-zА-Яа-я]+\d*)\.", it.code or "")
    if not m:
        return False
    found = {x.lower() for x in re.findall(re.escape(m.group(1)) + r"\.\d+", text, re.I)}
    return bool(found) and it.code.lower() not in found


def first_word(name: str) -> str:
    """First meaningful word of an estimate name ("Шпилька", "Труба") for size-only articles."""
    for t in tokens(name)[:3]:
        if distinctive([t]) and len(t) >= 4:
            return t
    return ""


def tinko_candidates(it: Item, rows: list[dict]):
    ck, ak = key(it.code), key(it.article)
    if len(ck) >= 4:
        c = [r for r in rows if ck in (r["k_art"], r["k_art2"]) or ck in r["paren"]]
        if c:
            return c, "по коду"
    if len(ak) >= 3:
        c = [r for r in rows if ak in (r["k_trade"], r["k_art"], r["k_art2"]) or ak in r["k_trade_vars"]]
        if c:
            return c, "по артикулу"
    at = tokens(it.article)
    if distinctive(at) and len(ak) >= 5:
        c = [r for r in rows if ak in r["k_trade"] + "|" + r["k_name"] and contains_seq(r["toks"], at)
             and not code_conflict(it, r["name"])]
        if c:
            return c, "по названию"
    w = first_word(it.name)
    if at and not distinctive(at) and len(ak) >= 4 and w:
        c = [r for r in rows if ak in r["k_name"] and w in r["toks"]]
        if c:
            return c, "по названию и размеру"
    return [], ""


def base_candidates(it: Item, rows: list[dict], kind: str, rule: str = ""):
    rows = [r for r in rows if r["kind"] == kind]
    if rule:
        if rule.lower() == "none":
            return [], "справочник: без цены"
        out = []
        for part in rule.split("||"):
            typ, _, val = part.strip().partition(":")
            typ, val = typ.strip().lower(), val.strip()
            if typ == "art":
                out += [r for r in rows if r["k_art"] == key(val)]
            elif typ == "name":
                vk, vt = key(val), tokens(val)
                out += [r for r in rows if vk in r["k_text"] and contains_seq(r["toks"], vt)]
        return list({id(r): r for r in out}.values()), "справочник"
    ck = key(it.code)
    by_code = [r for r in rows if len(ck) >= 4 and (r["k_art"] == ck or ck in r["paren"] or ck in r["tok_keys"])]
    if by_code:  # the supplier code is the most specific identifier
        return by_code, "по коду"
    tried = set()
    for label, art in (("артикул", it.article), ("артикул без скобок", strip_paren(it.article)),
                       ("без исполнения", strip_variant(it.article))):
        ak, at = key(art), tokens(art)
        if len(ak) < 3 or ak in tried:
            continue
        tried.add(ak)
        c = [r for r in rows if r["k_art"] == ak]
        if distinctive(at) and len(ak) >= 4:
            c += [r for r in rows if r["k_art"] != ak and ak in r["k_text"]
                  and (contains_seq(r["toks"], at) or ak in r["tok_keys"])]
        c = [r for r in c if not code_conflict(it, r["raw"])]
        if c:
            return c, label
    at, ak, w = tokens(it.article), key(it.article), first_word(it.name)
    if at and not distinctive(at) and len(ak) >= 4 and w:
        c = [r for r in rows if ak in r["k_text"] and w in r["toks"] and not code_conflict(it, r["raw"])]
        if c:
            return c, "по названию и размеру"
    return [], ""


def parse_factor(text: str) -> float:
    """'0.5', '100', '1/3' -> float."""
    a, _, b = text.replace(",", ".").partition("/")
    return float(a) / float(b) if b else float(a)


def fmt(p: float) -> str:
    return f"{p:,.2f}".replace(",", " ").replace(".00", "")


def base_reference(it: Item, base: list[dict]):
    """Freshest material price for the same article in the base: a sanity check for supplier matches."""
    cands, how = base_candidates(it, base, "material")
    if how not in ("артикул", "по коду"):
        return None
    conv = []
    for r in cands:
        f, _, sure = unit_factor(r["unit"], it.unit, r["name"], it.text)
        if f is not None and sure:
            conv.append((r, r["price"] * f))
    if not conv:
        return None
    d = max(r["date"] for r, _ in conv)
    return {"price": statistics.median(p for r, p in conv if r["date"] == d), "date": d}


def match_material(it: Item, trows: list[dict], pages: dict, me: dict, base: list[dict], capped: set) -> Result:
    res = Result()
    if me.get("tinko_code"):
        cands, how = [r for r in trows if r["code"] == me["tinko_code"]], "справочник"
        if not cands:
            res.notes.append(f"код Тинко {me['tinko_code']} из справочника нет в выгрузке")
    else:
        cands, how = tinko_candidates(it, trows)
    if how not in ("по коду", "по артикулу", "справочник") and any(key(t) in capped for t in watch_terms(it)):
        res.notes.append("запрос слишком общий: у поставщика много похожих позиций, часть могла не попасть в выгрузку")
    if not it.article and not it.code:
        res.notes.append("в смете нет марки/артикула — цену поставщика подобрать нельзя")
    if me.get("material_comment"):
        res.notes.append(me["material_comment"])
    cands = [r for r in cands if r["price"] and (r["currency"] or "руб").lower().startswith("руб")]
    maker_conflict = False
    if it.maker and cands:
        same = [r for r in cands if same_maker(it.maker, r["manufacturer"])]
        if same:
            cands = same
        else:
            maker_conflict = True
            res.notes.append("у поставщика другой производитель: " + ", ".join(sorted({r["manufacturer"] for r in cands})))
    conv = []
    for r in cands:
        if me.get("material_factor"):
            f, note, sure = parse_factor(me["material_factor"]), f"коэффициент из справочника × {me['material_factor']}", True
        else:
            f, note, sure = unit_factor(r["unit"], it.unit, r["name"], it.text)
        if f is not None:
            conv.append((r, f, note, sure))
        elif not me.get("material_factor"):
            res.notes.append(f"код {r['code']}: {note}")
    if len(conv) > 1:
        exact = [c for c in conv if c[0]["k_trade"] == key(it.article)]
        if len(exact) == 1:
            conv = exact
    page = pages.get(me.get("page_url", "")) if me.get("page_url") else None
    page_price = float(page["price_rub"]) if page and page.get("price_rub") else None

    if conv:
        prices = sorted({round(r["price"] * f, 2) for r, f, _, _ in conv})
        if len(prices) > 1:
            res.status, res.how = AMBIG, how
            res.alts = "; ".join(f"[{r['code']}] {r['trade_name']} — {fmt(round(r['price'] * f, 2))}" for r, f, _, _ in conv[:6])
            res.notes.append("несколько подходящих позиций с разной ценой — выберите нужную (mapping/items.csv)")
            return res
        r, f, note, sure = conv[0]
        res.price, res.how = prices[0], how
        res.status = OK if (how in ("по коду", "по артикулу", "справочник") and sure and not maker_conflict
                            and me.get("material_status") != "review") else CHECK
        if note:
            res.notes.append(note)
        if len(conv) > 1:
            res.status = CHECK
            res.notes.append("подходят несколько позиций с одинаковой ценой: " + ", ".join(c[0]["code"] for c in conv))
        if how == "по названию":
            res.notes.append("найдено по вхождению артикула в название")
        elif how == "по названию и размеру":
            res.notes.append("найдено по названию и размеру — артикула в смете нет")
        if me.get("material_status") == "confirmed" and res.status == CHECK:
            res.status = OK
        if (r["availability"] or "").lower().startswith("звон"):
            res.notes.append("на складе нет («Звоните»), цена по действующему прайсу")
        res.src = {
            "supplier": "ТД ТИНКО", "code": r["code"], "name": r["name"], "maker": r["manufacturer"], "unit": r["unit"],
            "price": r["price"], "factor": f, "availability": r["availability"],
            "url": TINKO_URL.format(int(r["code"])) if r["code"].isdigit() else "",
        }
        if page_price:
            res.notes.append(f"сверка: {page['supplier']} {fmt(page_price)} ₽")
        ref = base_reference(it, base)
        if ref and not 0.5 <= res.price / ref["price"] <= 2.0:
            if me.get("material_status") != "confirmed":
                res.status = CHECK
            res.notes.append(f"сильно отличается от базы: {fmt(round(ref['price'], 2))} ₽ на {ref['date']}")
        return res

    if page_price:
        f = parse_factor(me["material_factor"]) if me.get("material_factor") else 1.0
        res.price, res.status, res.how = round(page_price * f, 2), CHECK, "страница товара"
        res.notes.append(f"розница {page['supplier']}; НДС на сайте не указан")
        res.src = {"supplier": page["supplier"], "code": "", "name": page["name"], "maker": "", "unit": "",
                   "price": page_price, "factor": f, "availability": page.get("availability", ""), "url": page["url"],
                   "date": page.get("checked_date", "")}
        return res
    if me.get("page_url") and not page_price:
        res.notes.append("цена со страницы товара ещё не получена (workflow supplier-pages)")
    if not res.notes:
        res.notes.append("у Тинко такой позиции нет")
    return res


def match_installation(it: Item, base: list[dict], me: dict) -> Result:
    res = Result()
    cands, how = base_candidates(it, base, "installation", me.get("install_match", ""))
    res.how = how
    if how == "справочник: без цены":  # no separate installation (e.g. software): an explicit 0 shows as a dash
        res.price, res.status = 0.0, OK
        res.notes.append(me.get("install_comment") or "отдельного монтажа нет")
        return res
    if how == "справочник" and me.get("install_comment"):
        res.notes.append(me["install_comment"])
    conv = []
    for r in cands:
        if me.get("install_factor"):
            f, note, sure = parse_factor(me["install_factor"]), f"коэффициент из справочника × {me['install_factor']}", True
        else:
            f, note, sure = unit_factor(r["unit"], it.unit, r["name"], it.text)
        if f is not None:
            conv.append((r, round(r["price"] * f, 2), note, sure))
    if not conv:
        res.notes.append("в базе нет цены монтажа" if not cands else "в базе есть цены, но в несовместимых единицах")
        return res
    top_date = max(c[0]["date"] for c in conv)
    def reliability(row):  # "Надежная база": AUTO_OK is trusted, CHECK must be verified, MANUAL is weak
        return 2 if "MANUAL" in row["system"] else 1 if "CHECK" in row["system"] else 0
    top = sorted((c for c in conv if c[0]["date"] == top_date), key=lambda c: (reliability(c[0]), -c[1], -c[0]["sid"]))
    r, price, note, sure = top[0]
    res.price = price
    res.status = OK if how in ("артикул", "артикул без скобок", "по коду", "справочник") and sure else CHECK
    if how == "справочник" and me.get("install_status") == "review":
        res.status = CHECK
    if how in ("артикул без скобок", "без исполнения", "по названию и размеру"):
        res.notes.append(f"найдено {how}")
    if note:
        res.notes.append(note)
    distinct = sorted({c[1] for c in top}, reverse=True)
    if len(distinct) > 1:
        res.status = CHECK
        best = [c for c in top if reliability(c[0]) == reliability(r)]
        why = "взята большая" if len({c[1] for c in best}) > 1 else "взята более надёжная строка (AUTO_OK)"
        res.notes.append(f"на {top_date} в базе несколько цен ({', '.join(fmt(p) for p in distinct)}): {why}")
    if "MANUAL" in r["system"]:
        res.status = CHECK
        res.notes.append("строка «Надежной базы» со статусом MANUAL")
    if me.get("install_status") == "confirmed" and res.status == CHECK:
        res.status = OK
    older = sorted({(c[0]["date"], c[1]) for c in conv if c[0]["date"] != top_date}, reverse=True)
    res.alts = "; ".join(f"{fmt(p)} ({d})" for d, p in older[:5])
    res.src = {"system": r["system"], "date": r["date"], "source": r["source"], "excel_row": r["excel_row"],
               "name": r["name"], "article": r["article"], "unit": r["unit"], "price": r["price"]}
    return res


# ---------------------------------------------------------------- output

# (title, width, kind): kind = text | center | money | num
SRC_GROUPS = [("Позиция сметы", 6), ("Материал", 11), ("Монтаж", 8)]
SRC_COLUMNS = [
    ("Строка", 7, "center"), ("Поз.", 6, "center"), ("Наименование (смета)", 34, "text"),
    ("Тип / марка (смета)", 26, "text"), ("Ед.", 6, "center"), ("Кол-во", 9, "num"),
    ("Цена, ₽/ед.", 15, "money"), ("Статус", 12, "center"), ("Поставщик", 12, "center"), ("Код", 9, "center"),
    ("Наименование у поставщика", 36, "text"), ("Производитель", 14, "center"), ("Ед. пост.", 7, "center"),
    ("Цена пост., ₽", 15, "money"), ("Наличие", 11, "center"), ("Ссылка / строка счёта", 30, "text"), ("Комментарий", 44, "text"),
    ("Цена, ₽/ед.", 15, "money"), ("Статус", 12, "center"), ("Источник", 26, "text"), ("Дата цены", 11, "center"),
    ("Строка в источнике", 9, "center"), ("Позиция в источнике", 36, "text"), ("Более ранние цены", 26, "text"),
    ("Комментарий", 40, "text"),
]


def qty_format(value) -> str:
    """Integers without a trailing separator ("400", not "400,"); fractions with up to 3 decimals."""
    return "#,##0" if isinstance(value, (int, float)) and float(value).is_integer() else "#,##0.###"


def _lines(value, width: float, pt: float = 12) -> int:
    """Wrapped line count: greedy word wrap at ~1 char per width unit (Times New Roman 11–12 pt)."""
    if value is None or value == "":
        return 1
    cap = max(int(width * (1.0 if pt <= 11 else 0.93)) - 2, 4)
    total = 0
    for part in str(value).split("\n"):
        lines, cur = 1, 0
        for word in part.split():
            w = len(word)
            if cur == 0:
                cur = w
            elif cur + 1 + w <= cap:
                cur += 1 + w
            else:
                lines, cur = lines + 1, w
            while cur > cap:  # a word longer than the line breaks anywhere
                lines, cur = lines + 1, cur - cap
        total += lines
    return total


def fit_height(ws, row: int, widths: dict, pt: float = 12, min_h: float = 18) -> None:
    lines = max(_lines(ws.cell(row, c).value, w, pt) for c, w in widths.items())
    ws.row_dimensions[row].height = max(min_h, lines * pt * 1.25 + 4)


def _rewire(ws, old: str, new: str) -> None:
    """Point every formula of the sheet from cell `old` to cell `new`."""
    pat = re.compile(r"(?<![A-Z$])" + re.escape(old) + r"(?!\d)")
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("=") and pat.search(c.value):
                c.value = pat.sub(new, c.value)


def _unmerge_row(ws, r: int) -> None:
    for m in [m for m in ws.merged_cells.ranges if m.min_row <= r <= m.max_row]:
        ws.unmerge_cells(str(m))


def compact_expenses(ws, start: int, last_col: int) -> None:
    """"Прочие расходы" and "Сумма НДС": keep only № / наименование / сумма (user request 18.10.2026).

    Rates move to a hidden column right of the table (outside the print area) and the formulas follow
    them, so the totals do not change; the empty «Ставка / База / Ед. / Кол-во» cells disappear.
    """
    hid = last_col + 1
    hid_letter = get_column_letter(hid)
    border_l = ws.cell(start, 1).border
    for r in range(start, ws.max_row + 1):
        vals = {c: str(ws.cell(r, c).value or "").strip().lower() for c in range(1, last_col + 1)}
        rate_col = next((c for c, v in vals.items() if v == "ставка"), None)
        if not rate_col:
            continue
        rows = [r]
        e = r + 1
        while e <= ws.max_row and str(ws.cell(e, 1).value or "").strip()[:1].isdigit():
            rows.append(e)
            e += 1
        for rr in rows:
            if rr != r:
                rate = ws.cell(rr, rate_col)
                ws.cell(rr, hid, rate.value).number_format = rate.number_format
                _rewire(ws, f"{get_column_letter(rate_col)}{rr}", f"{hid_letter}{rr}")
            _unmerge_row(ws, rr)
            for c in range(rate_col, last_col):
                ws.cell(rr, c).value = None
            ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=last_col - 1)
            ws.row_dimensions[rr].height = 20
        ws.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center", indent=1)
        break
    for r in range(start, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        if not a.lower().startswith("сумма ндс"):
            continue
        pct = next((c for c in range(2, last_col) if isinstance(ws.cell(r, c).value, (int, float))
                    and "%" in (ws.cell(r, c).number_format or "")), None)
        if pct:
            rate = ws.cell(r, pct).value
            ws.cell(r, hid, rate).number_format = "0%"
            _rewire(ws, f"{get_column_letter(pct)}{r}", f"{hid_letter}{r}")
            _unmerge_row(ws, r)
            for c in range(2, last_col):
                ws.cell(r, c).value = None
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col - 1)
            ws.cell(r, 1).value = f'="В том числе НДС "&TEXT({hid_letter}{r},"0%")'  # follows the (hidden) rate
            ws.cell(r, 1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
        break
    ws.column_dimensions[hid_letter].hidden = True


def drop_trailing_notes(ws, last_col: int) -> int:
    """Remove the free-text notes under the table (no borders); returns the last row of the table."""
    last_table = max(r for r in range(1, ws.max_row + 1) if ws.cell(r, last_col).border.right.style)
    for r in range(ws.max_row, last_table, -1):
        _unmerge_row(ws, r)
    if ws.max_row > last_table:
        ws.delete_rows(last_table + 1, ws.max_row - last_table)
    return last_table


def style_estimate(ws, lay: dict, items) -> None:
    """Tidy the estimate sheet: accounting money format, alignment, grid, row heights, print titles."""
    last_col = max(c for c in range(1, ws.max_column + 1)
                   if any(ws.cell(r, c).value is not None for r in range(lay["header"], ws.max_row + 1)))
    widths = {c: (ws.column_dimensions[get_column_letter(c)].width or 9) for c in range(1, last_col + 1)}
    money_cols = {lay["mat"], lay["mat"] + 1, lay["inst"], lay["inst"] + 1}
    for row in ws.iter_rows(min_row=lay["header"], max_row=ws.max_row, max_col=last_col):
        for cell in row:
            if "₽" in (cell.number_format or "") or "руб" in (cell.number_format or ""):
                cell.number_format = FIN
    for it in items:
        r = it.row
        for c in range(1, last_col + 1):
            cell = ws.cell(r, c)
            cell.border = GRID
            cell.font = Font(name=FONT_NAME, size=12)
            if c in money_cols:
                cell.number_format = FIN
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif c == lay["qty"]:
                cell.number_format = qty_format(cell.value)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == lay["name"]:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        fit_height(ws, r, {c: widths[c] for c in range(1, last_col + 1) if c not in money_cols})
    # totals and "прочие расходы": percentages centred, text next to them left-aligned, money bold where it was
    for r in range(items[-1].row + 1 if items else lay["first"], ws.max_row + 1):
        for c in range(1, last_col + 1):
            cell = ws.cell(r, c)
            if "%" in (cell.number_format or "") and isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="center", vertical="center")
                nxt = ws.cell(r, c + 1)
                if isinstance(nxt.value, str):
                    nxt.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            elif cell.number_format == FIN:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(cell.value, str) and not cell.value.startswith("=") and c > 1 \
                    and cell.alignment.horizontal in (None, "general", "left"):
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        label = ws.cell(r, 1)
        if isinstance(label.value, str) and not label.value.startswith("="):
            if items and r == items[-1].row + 1:  # the line that closes the item table
                label.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            elif label.font.b and any(m.min_row == r and m.min_col == 1 and 2 <= m.max_col < last_col for m in ws.merged_cells.ranges):
                label.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        # rows of the "прочие расходы" table: an empty block under the price columns becomes one merged cell
        block = range(lay["mat"], lay["inst"] + 2)
        merged_here = any(m.min_row <= r <= m.max_row and m.min_col <= lay["mat"] <= m.max_col for m in ws.merged_cells.ranges)
        if not merged_here and ws.cell(r, 1).value is not None and ws.cell(r, last_col).value is not None \
                and all(ws.cell(r, c).value is None for c in block):
            ws.merge_cells(start_row=r, start_column=block[0], end_row=r, end_column=block[-1])
    tail = items[-1].row + 1 if items else lay["first"]
    compact_expenses(ws, tail, last_col)
    last_row = drop_trailing_notes(ws, last_col)
    total_label = ws.cell(tail, 1)
    if str(total_label.value or "").lower().startswith("итого по заполненным"):
        total_label.value = "Итого"
    # compact heights: title, object line, section bands, totals
    for r in range(1, lay["header"]):
        ws.row_dimensions[r].height = 22 if r == 1 else 16
    item_rows = {it.row for it in items}
    for r in range(lay["first"], last_row + 1):
        if r not in item_rows and ws.row_dimensions[r].height and ws.row_dimensions[r].height > 22:
            ws.row_dimensions[r].height = 20
    ws.print_area = f"A1:{get_column_letter(last_col)}{last_row}"
    ws.print_title_rows = f"{lay['header']}:{lay['first'] - 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    m = ws.page_margins
    m.top = m.bottom = 0.4
    m.left = m.right = 0.3
    m.header = m.footer = 0.2


def write_sources(wb, items, results, meta, estimate_name):
    if "Источники" in wb.sheetnames:
        del wb["Источники"]
    ws = wb.create_sheet("Источники")
    ncol = len(SRC_COLUMNS)
    last = get_column_letter(ncol)
    now = dt.datetime.now(MSK)
    head = [
        (f"Источники цен: «{estimate_name}», заполнено {now:%d.%m.%Y %H:%M} МСК", Font(name=FONT_NAME, size=14, bold=True)),
        (f"Материал — цена по счёту поставщика, если позиция есть в счёте (suppliers/invoices); иначе розничная цена ТД ТИНКО с НДС («{meta['source_title']}»); если у Тинко позиции нет — "
         "розница со страницы товара (ссылка в строке).", None),
        ("Монтаж — самая свежая цена из базы cvsig_prices; если на эту дату цен несколько — взята большая.", None),
        ("Подсветка: светло-жёлтая — цена вписана, но её нужно проверить (причина в комментарии); "
         "светло-красная — цену подобрать не удалось.", None),
    ]
    for i, (line, font) in enumerate(head, 1):
        ws.merge_cells(f"A{i}:{last}{i}")
        cell = ws.cell(i, 1, line)
        cell.font = font or Font(name=FONT_NAME, size=11)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[i].height = 22 if i == 1 else 17
    g, hr = len(head) + 2, len(head) + 3  # group header row, column header row
    col = 1
    for title, span in SRC_GROUPS:
        ws.merge_cells(start_row=g, start_column=col, end_row=g, end_column=col + span - 1)
        ws.cell(g, col, title)
        col += span
    for c, (title, width, _) in enumerate(SRC_COLUMNS, 1):
        ws.cell(hr, c, title)
        ws.column_dimensions[get_column_letter(c)].width = width
    for r in (g, hr):
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name=FONT_NAME, size=11, bold=True)
            cell.fill = HEAD_FILL
            cell.border = GRID
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[g].height = 18
    ws.row_dimensions[hr].height = 32
    widths = {c: w for c, (_, w, _) in enumerate(SRC_COLUMNS, 1)}
    for i, (it, mat, inst) in enumerate(results, hr + 1):
        ms, ins = mat.src, inst.src
        values = [
            it.row, it.pos, it.name, it.article_cell, it.unit, it.qty,
            mat.price, mat.status, ms.get("supplier", ""), ms.get("code", ""), ms.get("name", ""), ms.get("maker", ""),
            ms.get("unit", ""), ms.get("price"), ms.get("availability", ""), ms.get("url", ""),
            "; ".join(mat.notes + ([f"варианты: {mat.alts}"] if mat.alts else [])),
            inst.price, inst.status, ins.get("system", ""), ins.get("date", ""), ins.get("excel_row", ""),
            " / ".join(x for x in (ins.get("name", ""), ins.get("article", "")) if x), inst.alts, "; ".join(inst.notes),
        ]
        for c, v in enumerate(values, 1):
            kind = SRC_COLUMNS[c - 1][2]
            cell = ws.cell(i, c, v if v != "" else None)
            cell.font = Font(name=FONT_NAME, size=11)
            cell.border = GRID
            if kind == "money":
                cell.number_format = FIN
                cell.alignment = Alignment(horizontal="right", vertical="top")
            elif kind == "num":
                cell.number_format = qty_format(v)
                cell.alignment = Alignment(horizontal="center", vertical="top")
            else:
                cell.alignment = Alignment(horizontal="left" if kind == "text" else "center", vertical="top",
                                           wrap_text=True)
        if str(ms.get("url") or "").startswith("http"):
            link = ws.cell(i, 16)
            link.hyperlink = ms["url"]
            link.font = Font(name=FONT_NAME, size=11, color="0563C1", underline="single")
        for c, st in ((8, mat.status), (19, inst.status)):
            if st in (NONE, AMBIG):
                ws.cell(i, c).fill = YELLOW
            elif st == CHECK:
                ws.cell(i, c).fill = ORANGE
        fit_height(ws, i, widths, pt=11, min_h=16)
    ws.freeze_panes = ws.cell(hr + 1, 4)
    ws.auto_filter.ref = f"A{hr}:{last}{hr + len(results)}"
    ws.print_title_rows = f"{g}:{hr}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.sheet_view.zoomScale = 85


def put_price(ws, row, col, res: Result, overwrite: bool):
    cell = ws.cell(row, col)
    if cell.value not in (None, "") and not overwrite:
        res.status, res.notes = KEPT, ["в файле уже была цена — не трогал"] + res.notes
        return
    cell.value = res.price
    if res.price is None:
        cell.fill = YELLOW
    elif res.status != OK:
        cell.fill = ORANGE


# ---------------------------------------------------------------- commands

def cmd_plan(args):
    _, ws, lay = open_estimate(args.estimate, args.sheet)
    items = read_items(ws, lay)
    label = args.estimate.stem
    new = add_watch([(t, label) for it in items for t in watch_terms(it)])
    print(f"Позиций: {len(items)}; новых терминов в {WATCH.relative_to(ROOT)}: {len(new)}")
    for t in new:
        print("  +", t)
    if new:
        print("Дальше: git commit + push suppliers/tinko/watch.csv -> workflow tinko-prices (~2 мин) -> git pull -> fill")
    else:
        print("Все позиции уже в списке. Для свежих цен: gh workflow run tinko-prices.yml -> git pull -> fill")


def cmd_watch(args):
    new = add_watch([(t, args.comment) for t in args.terms])
    print(f"Добавлено: {len(new)}", *new, sep="\n  ")


def cmd_fill(args):
    meta, trows = load_tinko()
    print(check_fresh(meta, args.max_age_hours, args.allow_stale))
    if meta.get("watch_sha256") != hashlib.sha256(WATCH.read_bytes()).hexdigest():
        raise SystemExit("prices.csv собран по старой версии watch.csv: дождитесь workflow tinko-prices и сделайте git pull")
    wb, ws, lay = open_estimate(args.estimate, args.sheet)
    items = read_items(ws, lay)
    watched = {key(r["term"]) for r in read_watch()}
    missing = sorted({t for it in items for t in watch_terms(it) if key(t) not in watched})
    if missing:
        raise SystemExit(f"Нет в списке наблюдения ({len(missing)}): сначала `plan`, push и workflow. Пример: {missing[:5]}")
    pages, mapping, base, invoices = load_pages(), load_mapping(), load_base(), load_invoices()
    capped = {key(t) for t in meta.get("capped_terms", [])}
    results = []
    for it in items:
        me = mapping_for(it, mapping)
        mat = match_material(it, trows, pages, me, base, capped)
        by_invoice = invoice_result(it, invoices, me, set(args.invoice or []))
        if by_invoice is not None and (by_invoice.price is not None or by_invoice.status == AMBIG):
            if mat.price is not None and mat.src.get("supplier") == "ТД ТИНКО":
                by_invoice.notes.append(f"сверка: розница Тинко {fmt(mat.price)} (код {mat.src['code']})")
            mat = by_invoice
        inst = match_installation(it, base, me)
        if mat.price and inst.price and inst.price > 10 * mat.price:
            inst.status = CHECK
            inst.notes.append(f"монтаж в {inst.price / mat.price:.0f} раз дороже материала — проверьте")
        put_price(ws, it.row, lay["mat"], mat, args.overwrite)
        put_price(ws, it.row, lay["inst"], inst, args.overwrite)
        results.append((it, mat, inst))
    style_estimate(ws, lay, items)
    write_sources(wb, items, results, meta, args.estimate.name)
    out = args.output or args.estimate.with_name(f"{args.estimate.stem}_цены_{dt.date.today():%Y-%m-%d}.xlsx")
    wb.save(out)

    def count(idx):
        c = {}
        for r in results:
            c[r[idx].status] = c.get(r[idx].status, 0) + 1
        return ", ".join(f"{k} {v}" for k, v in sorted(c.items()))

    print(f"Позиций: {len(items)}\nМатериал: {count(1)}\nМонтаж:   {count(2)}\nФайл: {out}")
    for it, mat, inst in results:
        if mat.status != OK or inst.status != OK:
            print(f"  {it.pos:5} {it.article[:28]:28} | мат {mat.status}: {'; '.join(mat.notes)[:90]}"
                  f" | монт {inst.status}: {'; '.join(inst.notes)[:70]}")


def main():
    ap = argparse.ArgumentParser(description="Fill an estimate with current retail material and installation prices.")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("plan", help="add the estimate's items to the TINKO watch list")
    p.add_argument("estimate", type=Path)
    p.add_argument("--sheet")
    p.set_defaults(func=cmd_plan)
    w = sub.add_parser("watch", help="add arbitrary search terms to the TINKO watch list")
    w.add_argument("terms", nargs="+")
    w.add_argument("--comment", default="вручную")
    w.set_defaults(func=cmd_watch)
    f = sub.add_parser("fill", help="write prices into a copy of the estimate")
    f.add_argument("estimate", type=Path)
    f.add_argument("-o", "--output", type=Path)
    f.add_argument("--sheet")
    f.add_argument("--overwrite", action="store_true", help="replace prices already present in the file")
    f.add_argument("--invoice", action="append", help="match items automatically against this invoice (number, e.g. ЛП121541)")
    f.add_argument("--max-age-hours", type=float, default=48)
    f.add_argument("--allow-stale", action="store_true")
    f.set_defaults(func=cmd_fill)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
