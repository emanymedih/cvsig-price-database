from __future__ import annotations

import json
import re
import sqlite3
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl


WORKBOOK_PATH = Path(r"C:\Users\admla\Desktop\Синтек 0806.xlsx")
DB_PATH = Path(r"C:\Users\admla\Documents\Codex\2026-06-02\files-mentioned-by-the-user-02\outputs\cvsig_prices_2026-06-02.sqlite")


STOPWORDS = {
    "и",
    "для",
    "на",
    "с",
    "со",
    "из",
    "без",
    "по",
    "под",
    "к",
    "в",
    "шт",
    "м",
    "п",
    "исп",
    "тип",
    "код",
    "арт",
    "артикул",
    "ооо",
    "зао",
    "нвп",
}


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).lower().replace("ё", "е")
    text = re.sub(r"[«»\"'`´“”„]", " ", text)
    text = text.replace("–", "-").replace("—", "-").replace("−", "-")
    text = re.sub(r"(?<=\d),(?=\d)", ".", text)
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def tokens(value) -> set[str]:
    return {token for token in normalize_text(value).split() if token not in STOPWORDS and len(token) > 1}


def compact(value) -> str:
    return re.sub(r"[^0-9a-zа-я]+", "", normalize_text(value))


def number_tokens(value) -> set[str]:
    return set(re.findall(r"\d+(?:\.\d+)?", normalize_text(value)))


def model_tokens(value) -> set[str]:
    return {
        token
        for token in normalize_text(value).split()
        if any(ch.isdigit() for ch in token) and len(token) >= 2
    }


def money_with_vat(value: float | None, vat_note: str | None) -> float | None:
    if value is None:
        return None
    return round(float(value), 2)


@dataclass
class DbItem:
    system: str
    name: str
    type_code_article: str | None
    manufacturer: str | None
    unit: str | None
    material_unit: float | None
    installation_unit: float | None
    material_vat_note: str | None
    installation_vat_note: str | None
    name_norm: str
    type_norm: str
    manufacturer_norm: str
    name_tokens: set[str]
    type_tokens: set[str]
    type_compact: str


def load_db_items() -> list[DbItem]:
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT
                system,
                name,
                type_code_article,
                manufacturer,
                unit,
                material_unit_price_rub,
                installation_unit_price_rub,
                material_vat_note,
                installation_vat_note
            FROM v_price_lines
            WHERE name IS NOT NULL
            """
        ).fetchall()

    items = []
    for row in rows:
        name = row["name"] or ""
        item_type = row["type_code_article"] or ""
        manufacturer = row["manufacturer"] or ""
        items.append(
            DbItem(
                system=row["system"],
                name=name,
                type_code_article=row["type_code_article"],
                manufacturer=row["manufacturer"],
                unit=row["unit"],
                material_unit=money_with_vat(row["material_unit_price_rub"], row["material_vat_note"]),
                installation_unit=money_with_vat(
                    row["installation_unit_price_rub"], row["installation_vat_note"]
                ),
                material_vat_note=row["material_vat_note"],
                installation_vat_note=row["installation_vat_note"],
                name_norm=normalize_text(name),
                type_norm=normalize_text(item_type),
                manufacturer_norm=normalize_text(manufacturer),
                name_tokens=tokens(name),
                type_tokens=tokens(item_type),
                type_compact=compact(item_type),
            )
        )
    return items


def score_item(target: dict, item: DbItem) -> float:
    target_name_norm = target["name_norm"]
    target_type_norm = target["type_norm"]
    target_manufacturer_norm = target["manufacturer_norm"]
    target_name_tokens = target["name_tokens"]
    target_type_tokens = target["type_tokens"]
    target_type_compact = target["type_compact"]

    name_ratio = SequenceMatcher(None, target_name_norm, item.name_norm).ratio()
    token_overlap = 0.0
    if target_name_tokens and item.name_tokens:
        token_overlap = len(target_name_tokens & item.name_tokens) / len(target_name_tokens | item.name_tokens)

    type_quality = 0.0
    if target_type_compact and item.type_compact:
        if target_type_compact == item.type_compact:
            type_quality = 1.0
        elif target_type_compact in item.type_compact or item.type_compact in target_type_compact:
            short = min(len(target_type_compact), len(item.type_compact))
            long = max(len(target_type_compact), len(item.type_compact))
            type_quality = max(0.55, short / long)
            if short >= 6 and any(ch.isdigit() for ch in target_type_compact + item.type_compact):
                type_quality = max(type_quality, 0.86)
        elif target_type_tokens and item.type_tokens:
            type_quality = len(target_type_tokens & item.type_tokens) / len(target_type_tokens | item.type_tokens)

    manufacturer_bonus = 0.0
    if target_manufacturer_norm and item.manufacturer_norm:
        manufacturer_bonus = 0.05 * SequenceMatcher(
            None, target_manufacturer_norm, item.manufacturer_norm
        ).ratio()

    unit_bonus = 0.03 if target["unit_norm"] and normalize_text(item.unit) == target["unit_norm"] else 0.0
    return round(
        0.35 * name_ratio
        + 0.22 * token_overlap
        + 0.60 * type_quality
        + manufacturer_bonus
        + unit_bonus,
        4,
    )


def match_action(target: dict, item: DbItem, score: float, second_score: float) -> str:
    name_ratio = SequenceMatcher(None, target["name_norm"], item.name_norm).ratio()
    token_overlap = 0.0
    if target["name_tokens"] and item.name_tokens:
        token_overlap = len(target["name_tokens"] & item.name_tokens) / len(
            target["name_tokens"] | item.name_tokens
        )
    type_quality = 0.0
    if target["type_compact"] and item.type_compact:
        if target["type_compact"] == item.type_compact:
            type_quality = 1.0
        elif target["type_compact"] in item.type_compact or item.type_compact in target["type_compact"]:
            short = min(len(target["type_compact"]), len(item.type_compact))
            long = max(len(target["type_compact"]), len(item.type_compact))
            type_quality = max(0.55, short / long)
            if short >= 6 and any(ch.isdigit() for ch in target["type_compact"] + item.type_compact):
                type_quality = max(type_quality, 0.86)

    target_numbers = number_tokens(target["name"])
    item_numbers = number_tokens(item.name)
    if not target["type_compact"] and target_numbers and item_numbers and target_numbers != item_numbers:
        return "review"

    target_type_norm = target["type_norm"]
    item_type_norm = item.type_norm
    if "исп" in target_type_norm and "исп" in item_type_norm:
        target_exec = set(re.findall(r"исп\s*(\d+)", target_type_norm))
        item_exec = set(re.findall(r"исп\s*(\d+)", item_type_norm))
        if target_exec and item_exec and not (target_exec & item_exec):
            return "review"

    target_combined_numbers = number_tokens(f"{target['name']} {target['type'] or ''}")
    item_combined_numbers = number_tokens(f"{item.name} {item.type_code_article or ''}")
    if target_type_norm.startswith("гост") and target_numbers:
        target_dimensions = target_combined_numbers - number_tokens(target["type"])
        item_dimensions = item_combined_numbers - number_tokens(item.type_code_article)
        if target_dimensions and item_dimensions and not (target_dimensions & item_dimensions):
            return "review"

    target_model_tokens = model_tokens(target["type"])
    item_model_tokens = model_tokens(item.type_code_article)
    if (
        target_model_tokens
        and item_model_tokens
        and not (target_model_tokens & item_model_tokens)
        and type_quality == 0
    ):
        return "review"

    # Strong code/type match is enough even when the wording differs substantially.
    exact_type = bool(target["type_compact"] and item.type_compact and target["type_compact"] == item.type_compact)
    target_type_inside_db = bool(
        target["type_compact"] and item.type_compact and target["type_compact"] in item.type_compact
    )
    db_type_inside_target = bool(
        target["type_compact"] and item.type_compact and item.type_compact in target["type_compact"]
    )

    if exact_type:
        return "fill"
    if type_quality >= 0.86 and target_type_inside_db:
        return "fill"
    if type_quality >= 0.86 and db_type_inside_target and (name_ratio >= 0.35 or token_overlap >= 0.10):
        return "fill"
    if type_quality >= 0.78 and (name_ratio >= 0.28 or token_overlap >= 0.08):
        return "fill"
    if type_quality >= 0.55 and name_ratio >= 0.58 and token_overlap >= 0.18:
        return "fill"
    if name_ratio >= 0.84 and token_overlap >= 0.45 and score - second_score >= 0.02:
        return "fill"
    return "review"


def target_rows():
    workbook = openpyxl.load_workbook(WORKBOOK_PATH, data_only=True, read_only=True)
    for sheet in workbook.worksheets:
        for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            if row_idx == 1:
                continue
            position = row[0] if len(row) > 0 else None
            name = row[1] if len(row) > 1 else None
            if not isinstance(position, (int, float)) or not isinstance(name, str) or not name.strip():
                continue
            if str(name).strip().lower().startswith(("итого", "всего", "сумма ндс")):
                continue
            target = {
                "sheet": sheet.title,
                "row": row_idx,
                "position": position,
                "name": name,
                "type": row[2] if len(row) > 2 else None,
                "manufacturer": row[3] if len(row) > 3 else None,
                "unit": row[4] if len(row) > 4 else None,
                "quantity": row[5] if len(row) > 5 else None,
                "name_norm": normalize_text(name),
                "type_norm": normalize_text(row[2] if len(row) > 2 else None),
                "manufacturer_norm": normalize_text(row[3] if len(row) > 3 else None),
                "unit_norm": normalize_text(row[4] if len(row) > 4 else None),
                "name_tokens": tokens(name),
                "type_tokens": tokens(row[2] if len(row) > 2 else None),
                "type_compact": compact(row[2] if len(row) > 2 else None),
            }
            yield target


def best_match(target: dict, db_items: list[DbItem]):
    scored = sorted(((score_item(target, item), item) for item in db_items), key=lambda x: x[0], reverse=True)
    best_score, best_item = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0
    return best_score, second_score, best_item


def main():
    db_items = load_db_items()
    report = []
    for target in target_rows():
        best_score, second_score, item = best_match(target, db_items)
        action = match_action(target, item, best_score, second_score)
        report.append(
            {
                "sheet": target["sheet"],
                "row": target["row"],
                "position": target["position"],
                "target_name": target["name"],
                "target_type": target["type"],
                "target_qty": target["quantity"],
                "action": action,
                "score": best_score,
                "margin": round(best_score - second_score, 4),
                "match_system": item.system,
                "match_name": item.name,
                "match_type": item.type_code_article,
                "material_unit_with_vat": item.material_unit,
                "installation_unit_with_vat": item.installation_unit,
                "material_vat_note": item.material_vat_note,
                "installation_vat_note": item.installation_vat_note,
            }
        )

    summary = {
        "targets": len(report),
        "auto_fill": sum(1 for row in report if row["action"] == "fill"),
        "review": sum(1 for row in report if row["action"] == "review"),
        "by_sheet": {},
        "review_rows": [row for row in report if row["action"] == "review"][:25],
    }
    for row in report:
        summary["by_sheet"].setdefault(row["sheet"], {"targets": 0, "auto_fill": 0, "review": 0})
        summary["by_sheet"][row["sheet"]]["targets"] += 1
        summary["by_sheet"][row["sheet"]][row["action"] == "fill" and "auto_fill" or "review"] += 1

    out = Path("work/sintek_match_report.json")
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
