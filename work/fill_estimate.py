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
from openpyxl.styles import Alignment, Font, PatternFill
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
MAPPING = ROOT / "mapping" / "items.csv"
MSK = dt.timezone(dt.timedelta(hours=3))
YELLOW = PatternFill("solid", fgColor="FFF2CC")  # no price found
ORANGE = PatternFill("solid", fgColor="FCE4D6")  # price written, needs a check
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
    top = sorted((c for c in conv if c[0]["date"] == top_date), key=lambda c: (-c[1], -c[0]["sid"]))
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
        res.notes.append(f"на {top_date} в базе несколько цен ({', '.join(fmt(p) for p in distinct)}): взята большая")
    if "MANUAL" in r["system"]:
        res.status = CHECK
        res.notes.append("строка «Надежной базы» со статусом MANUAL")
    older = sorted({(c[0]["date"], c[1]) for c in conv if c[0]["date"] != top_date}, reverse=True)
    res.alts = "; ".join(f"{fmt(p)} ({d})" for d, p in older[:5])
    res.src = {"system": r["system"], "date": r["date"], "source": r["source"], "excel_row": r["excel_row"],
               "name": r["name"], "article": r["article"], "unit": r["unit"], "price": r["price"]}
    return res


# ---------------------------------------------------------------- output

SRC_COLUMNS = [
    ("Строка", 7), ("Поз.", 6), ("Наименование (смета)", 34), ("Тип / марка (смета)", 26), ("Ед.", 6), ("Кол-во", 8),
    ("Материал, ₽/ед.", 13), ("Статус материала", 13), ("Поставщик", 11), ("Код", 9), ("Наименование у поставщика", 40),
    ("Производитель", 14), ("Ед. пост.", 7), ("Цена пост., ₽", 12), ("Наличие", 12), ("Ссылка", 34),
    ("Комментарий (материал)", 46),
    ("Монтаж, ₽/ед.", 12), ("Статус монтажа", 13), ("Источник монтажа", 34), ("Дата цены", 11), ("Строка в источнике", 9),
    ("Позиция в источнике", 40), ("Более ранние цены", 30), ("Комментарий (монтаж)", 40),
]


def write_sources(wb, items, results, meta, estimate_name):
    if "Источники" in wb.sheetnames:
        del wb["Источники"]
    ws = wb.create_sheet("Источники")
    now = dt.datetime.now(MSK)
    head = [
        f"Источники цен для «{estimate_name}», заполнено {now:%d.%m.%Y %H:%M} МСК",
        f"Материал: розничная цена ТД ТИНКО с НДС — «{meta['source_title']}»; если у Тинко позиции нет — розница со страницы товара (ссылка в строке).",
        "Монтаж: самая свежая цена из базы cvsig_prices; если на эту дату цен несколько — взята большая.",
        "Лист сметы: жёлтая ячейка — цену подобрать не удалось; оранжевая — цена вписана, но её нужно проверить (причина в комментарии).",
    ]
    for i, line in enumerate(head, 1):
        ws.cell(i, 1, line).font = Font(bold=(i == 1))
    hr = len(head) + 2
    for c, (title, width) in enumerate(SRC_COLUMNS, 1):
        cell = ws.cell(hr, c, title)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[get_column_letter(c)].width = width
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
            cell = ws.cell(i, c, v)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(v, float):
                cell.number_format = "#,##0.00"
        if ms.get("url"):
            ws.cell(i, 16).hyperlink = ms["url"]
            ws.cell(i, 16).style = "Hyperlink"
        for col, st in ((8, mat.status), (19, inst.status)):
            if st in (NONE, AMBIG):
                ws.cell(i, col).fill = YELLOW
            elif st == CHECK:
                ws.cell(i, col).fill = ORANGE
    ws.freeze_panes = ws.cell(hr + 1, 3)
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(SRC_COLUMNS))}{hr + len(results)}"


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
    pages, mapping, base = load_pages(), load_mapping(), load_base()
    capped = {key(t) for t in meta.get("capped_terms", [])}
    results = []
    for it in items:
        me = mapping_for(it, mapping)
        mat = match_material(it, trows, pages, me, base, capped)
        inst = match_installation(it, base, me)
        if mat.price and inst.price and inst.price > 10 * mat.price:
            inst.status = CHECK
            inst.notes.append(f"монтаж в {inst.price / mat.price:.0f} раз дороже материала — проверьте")
        put_price(ws, it.row, lay["mat"], mat, args.overwrite)
        put_price(ws, it.row, lay["inst"], inst, args.overwrite)
        results.append((it, mat, inst))
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
    f.add_argument("--max-age-hours", type=float, default=48)
    f.add_argument("--allow-stale", action="store_true")
    f.set_defaults(func=cmd_fill)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
