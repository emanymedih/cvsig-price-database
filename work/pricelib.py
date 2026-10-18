"""Shared helpers for price matching: comparison keys, tokens, units, manufacturers.

Standard library only: used in GitHub Actions (work/fetch_tinko.py)
and locally (work/fill_estimate.py).
"""
from __future__ import annotations

import re

# After lower(): Cyrillic letters that look like Latin ones -> Latin,
# so "С2000" typed with a Cyrillic or a Latin "С" compares equal.
_LOOK = str.maketrans("авекмнорстухё", "abekmhopctyxe")
_QUOTES = "«»\"'“”„‟″`´"
_KEY_DROP = re.compile("[\\s\\-‐‑–—_.,:;()\\[\\]{}/\\\\№+" + re.escape(_QUOTES) + "]")
_TOKEN = re.compile(r"[0-9a-zа-я][0-9a-zа-я.+\-]*")
# Pure sizes like "d20", "25x16", "m6x2000", "ral2004", "80x80x40mm" (already mapped by _LOOK).
_DIM = re.compile(r"^(?:[dm]|ral)?\d+(?:\.\d+)?(?:x\d+(?:\.\d+)?)*(?:mm|cm|m)?$")
_UNITISH = {"mm", "cm", "m", "d", "шt", "yп", "kг"}


def _prep(text) -> str:
    return str(text or "").lower().translate(_LOOK).replace("×", "x")


def key(text) -> str:
    """Compact comparison key: no case, no separators, look-alike letters unified.

    key("С2000-КДЛ") == key("C2000 КДЛ"); key("1х2х0,75") == key("1×2×0.75").
    """
    return _KEY_DROP.sub("", _prep(text))


def tokens(text) -> list[str]:
    """Word tokens that keep hyphenated model names whole: "С2000-КДЛ-С" is one token."""
    s = _prep(text).replace(",", ".")
    s = re.sub("[" + re.escape(_QUOTES) + r"()\[\]{}]", "", s)
    out = []
    for t in _TOKEN.findall(s):
        t = t.strip(".+-")
        if t:
            out.append(t)
    return out


def contains_seq(hay: list[str], needle: list[str]) -> bool:
    """True if needle tokens occur in hay contiguously (whole tokens only)."""
    n = len(needle)
    if not n or n > len(hay):
        return False
    first = needle[0]
    for i in range(len(hay) - n + 1):
        if hay[i] == first and hay[i:i + n] == needle:
            return True
    return False


def distinctive(toks: list[str]) -> bool:
    """An article is distinctive if it is more than sizes/units ("d20 мм" is not)."""
    for t in toks:
        if t in _UNITISH or _DIM.match(t):
            continue
        if re.search(r"[a-zа-я]", t):
            return True
    return False


def strip_paren(text) -> str:
    """Drop a trailing qualifier in brackets: "ДИП-34А-03 (ИП 212-34А)" -> "ДИП-34А-03".

    Brackets glued to a word are part of the model name and stay: "КПСнг(А)-FRHF".
    """
    return re.sub(r"\s+", " ", re.sub(r"\s\([^()]*\)", " ", str(text or ""))).strip()


def strip_variant(text) -> str:
    """Drop "(...)" and "исп.XX": "С2000-СП2 исп.02" -> "С2000-СП2"."""
    s = re.sub(r"\s*исп\.?\s*[\w.\-/]+", " ", strip_paren(text), flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


def paren_keys(text) -> set[str]:
    """Keys of "(...)" parts, e.g. supplier codes in names: "... (PR.012031)"."""
    return {k for k in (key(p) for p in re.findall(r"\(([^()]+)\)", str(text or ""))) if len(k) >= 3}


# ---------------------------------------------------------------- units

_UNIT_CLASSES = {
    "pcs": {"шт", "штука", "штук", "pcs", "ед", "единица"},
    "m": {"м", "мп", "пм", "погм", "метр", "m"},
    "km": {"км", "km"},
    "pack": {"уп", "упак", "упаковка", "ука", "упк", "пачка", "pack"},
    "set": {"кт", "компл", "комплект", "набор", "set"},
    "kg": {"кг", "kg"},
    "t": {"т", "тн", "тон", "тонна"},
    "l": {"л", "литр"},
    "m2": {"м2", "квм"},
    "m3": {"м3", "кубм"},
}


def unit_class(unit) -> str | None:
    s = re.sub(r"[\s.\-]", "", str(unit or "").lower())
    if not s or s.isdigit():  # empty or a shifted number in the unit column
        return None
    for cls, names in _UNIT_CLASSES.items():
        if s in names:
            return cls
    return "?" + s


def pack_size(text) -> int | None:
    s = str(text or "")
    m = re.search(r"\(\s*(\d+)\s*шт", s, re.I) or re.search(r"(\d+)\s*шт\.?\s*/\s*уп", s, re.I)
    if m and int(m.group(1)) > 0:
        return int(m.group(1))
    return None


def unit_factor(src_unit, dst_unit, src_text="", dst_text=""):
    """Multiplier that turns a price per src_unit into a price per dst_unit.

    Returns (factor, note, sure). factor is None when units are incompatible;
    sure=False means the conversion relies on an assumption and must be checked.
    """
    a, b = unit_class(src_unit), unit_class(dst_unit)
    if a is None or b is None:
        return 1.0, "единица измерения не указана", False
    if a == b == "pack":
        ns, nd = pack_size(src_text), pack_size(dst_text)
        if ns and nd and ns != nd:
            return nd / ns, f"упаковка {ns} шт → {nd} шт", False
        return 1.0, "", True
    if a == b:
        return 1.0, "", True
    if a == "km" and b == "m":
        return 0.001, "цена за км ÷ 1000", True
    if a == "m" and b == "km":
        return 1000.0, "цена за м × 1000", True
    if a == "pack" and b == "pcs":
        n = pack_size(src_text)
        if n:
            return 1.0 / n, f"упаковка {n} шт → цена за шт", False
        return None, "упаковка без указания количества", False
    if a == "pcs" and b == "pack":
        n = pack_size(dst_text)
        if n:
            return float(n), f"цена за шт × {n} (упаковка)", False
        return None, "в смете упаковка без указания количества", False
    if {a, b} == {"pcs", "set"}:
        return 1.0, "компл. = шт.", False
    return None, f"единицы несовместимы: {src_unit} → {dst_unit}", False


# ---------------------------------------------------------------- manufacturers

_TRANSLIT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e", "ж": "zh", "з": "z",
    "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r",
    "с": "s", "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "c", "ч": "ch", "ш": "sh", "щ": "sch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
}
_MAKER_PREFIXES = ("ooo", "zao", "oao", "pao", "nvp", "npo", "ao")
_MAKER_ALIASES = {"dkc": "dks"}


def maker_key(name) -> str:
    s = "".join(_TRANSLIT.get(ch, ch) for ch in str(name or "").lower())
    s = re.sub(r"[^a-z0-9]", "", s)
    changed = True
    while changed:
        changed = False
        for p in _MAKER_PREFIXES:
            if s.startswith(p) and len(s) > len(p) + 2:
                s, changed = s[len(p):], True
    return _MAKER_ALIASES.get(s, s)


def same_maker(a, b) -> bool:
    ka, kb = maker_key(a), maker_key(b)
    return bool(ka and kb and (ka in kb or kb in ka))
