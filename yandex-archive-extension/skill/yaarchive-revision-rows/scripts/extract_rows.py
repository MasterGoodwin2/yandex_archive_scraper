#!/usr/bin/env python3
"""
Яндекс Архив, ревизские сказки: привязка соседних столбцов к Основному блоку.

Вход : данные страницы (__NEXT_DATA__) в .json, либо HTML-страница с этим <script>,
       либо папка с такими файлами.
Выход: JSON (по странице: блоки -> строки -> ячейки) и плоский CSV (одна строка = одна
       строка таблицы внутри Основного блока).

Распознавание НЕ выполняется: используются только готовые тексты и координаты.
Подробности и допущения: см. SKILL.md.
"""
import argparse
import csv
import json
import re
import statistics
import sys
from pathlib import Path

NEXT_RE = re.compile(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', re.S)
_MAP = str.maketrans({"ѣ": "е", "і": "и", "i": "и", "ѳ": "ф", "ѵ": "и", "ё": "е"})

CSV_COLS = ["page", "node_id", "sex", "block_no", "block_id", "family", "family_no",
            "row_no", "rows_in_block", "names_text",
            "age_prev", "age_mid", "departed_when", "absent_since", "age_now", "flags"]


# ----------------------------------------------------------------------------- вход
def load_node(path):
    raw = Path(path).read_text(encoding="utf-8")
    m = NEXT_RE.search(raw)
    data = json.loads(m.group(1) if m else raw)
    if "props" in data:
        data = data["props"]["pageProps"]
    if "currentNode" in data:
        data = data["currentNode"]
    return data


def norm(t):
    t = t.lower().translate(_MAP)
    t = re.sub(r"[^\w\s/]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def header_kind(text):
    n = norm(text)
    if n == "крестьяне":
        return "main", "M"
    if n == "крестьянки":
        return "main", "F"
    if n == "семьи":
        return "family", None
    if n == "лета":
        return "age", None
    if n.startswith("когда имен"):
        return "departed_when", None
    if re.search(r"котор\w* времени", n):
        return "absent_since", None
    return None


class Blk:
    def __init__(self, i, j, t):
        pts = j["points"]
        xs, ys = [p["x"] for p in pts], [p["y"] for p in pts]
        self.i, self.id, self.tid = i, j["id"], t["id"]
        self.text = t["text"].strip()
        self.x0, self.y0, self.x1, self.y1 = min(xs), min(ys), max(xs), max(ys)
        self.w, self.h = self.x1 - self.x0, self.y1 - self.y0
        self.cx, self.cy = (self.x0 + self.x1) / 2, (self.y0 + self.y1) / 2

    def bb(self):
        return [self.x0, self.y0, self.x1, self.y1]


def cell_json(blocks):
    return {"text": " ".join(b.text for b in blocks),
            "ids": [b.id for b in blocks],
            "text_block_ids": [b.tid for b in blocks],
            "bbox": [min(b.x0 for b in blocks), min(b.y0 for b in blocks),
                     max(b.x1 for b in blocks), max(b.y1 for b in blocks)]}


def to_int(text):
    d = re.sub(r"\D", "", text or "")
    return int(d) if d else None


# ----------------------------------------------------------------------------- страница
def process(node, a):
    size = node.get("originalImageSize") or {}
    W, H = size.get("width"), size.get("height")
    jb = node.get("joinedTextBlocks") or []
    tb = node.get("textBlocks") or []
    rep = {"page": node.get("sheetPageNumber"), "node_id": node.get("id"), "image": [W, H],
           "flags": [], "blocks": [], "orphans": [], "totals": {}, "titles": []}
    if not (W and H) or not jb or len(jb) != len(tb):
        rep["flags"].append(f"count_mismatch:joined={len(jb)},text={len(tb)}")
        return rep

    blks = [Blk(i, j, t) for i, (j, t) in enumerate(zip(jb, tb))]

    # 1. заголовки столбцов (по тексту), они же задают x-границы столбцов
    heads = []
    for b in blks:
        if b.cy < 0.4 * H:
            k = header_kind(b.text)
            if k:
                heads.append((k, b))
    mains = sorted([h for h in heads if h[0][0] == "main"], key=lambda h: h[1].cx)
    if not mains:
        rep["flags"].append("no_main_header")
        return rep
    hdr_bottom = max(b.y1 for _, b in heads)
    # тексты шапки страницы (название НП, дата и т. п.) - для определения населённого пункта
    _hid = {b.id for _, b in heads}
    rep["titles"] = [b.text for b in sorted((b for b in blks if b.id not in _hid and b.cy <= hdr_bottom),
                                            key=lambda b: (b.cy, b.cx))]
    split_x = a.split_x or W / 2

    def group_of(b):
        if len(mains) == 1:
            return mains[0][0][1]
        return mains[0][0][1] if b.cx < split_x else mains[-1][0][1]

    cols, by_group = [], {}
    for k, b in heads:
        by_group.setdefault(group_of(b), []).append((k[0], b))
    for g, items in by_group.items():
        items.sort(key=lambda t: t[1].cx)
        ages = [b for k, b in items if k == "age"]
        for k, b in items:
            role = k
            if k == "age":
                role = "age_now" if b is ages[-1] else ("age_prev" if b is ages[0] else "age_mid")
            cols.append({"role": role, "x0": b.x0, "x1": b.x1, "group": g})

    head_ids = {b.id for _, b in heads}
    body = [b for b in blks if b.id not in head_ids and b.cy > hdr_bottom]

    # 2. итоговая строка внизу страницы («Итого мужескаго пола налицо» + число справа)
    foot_ids = set()
    for f in [b for b in body if norm(b.text).startswith("итого")]:
        foot_ids.add(f.id)
        sex = "M" if "мужеск" in norm(f.text) else ("F" if "женск" in norm(f.text) else None)
        right = [b for b in body if b.id not in foot_ids and b.x0 >= f.x1 - 0.01 * W
                 and abs(b.cy - f.cy) < 0.5 * max(b.h, f.h)]
        right.sort(key=lambda b: b.x0)
        if sex and right:
            foot_ids.add(right[0].id)
            rep["totals"][sex] = to_int(right[0].text)
    rest = [b for b in body if b.id not in foot_ids]

    # 3. каждому блоку - столбец по x
    def col_of(b):
        best = None
        for c in cols:
            dx = max(c["x0"] - b.cx, 0, b.cx - c["x1"])
            if best is None or dx < best[0]:
                best = (dx, c)
        return best

    mainb, fam, cells = {}, {}, []
    for b in rest:
        dx, c = col_of(b)
        if dx > a.col_tol * W:
            rep["orphans"].append({"text": b.text, "id": b.id, "bbox": b.bb(), "why": "no_column"})
            continue
        if c["role"] == "main":
            mainb.setdefault(c["group"], []).append(b)
        elif c["role"] == "family":
            fam.setdefault(c["group"], []).append(b)
        else:
            cells.append((b, c))
    for g in mainb:
        mainb[g].sort(key=lambda b: b.y0)
    for g in fam:
        fam[g].sort(key=lambda b: b.cy)

    # 4. ячейки -> Основной блок своей половины (самый маленький из содержащих по y)
    assigned = {m.id: [] for g in mainb for m in mainb[g]}
    nested = set()
    for b, c in cells:
        cand = mainb.get(c["group"], [])
        inside = [m for m in cand if m.y0 <= b.cy <= m.y1]
        flag = None
        if inside:
            m = min(inside, key=lambda m: m.h)
            if len(inside) > 1:
                nested.add(m.id)
        else:
            near = sorted(cand, key=lambda m: min(abs(b.cy - m.y0), abs(b.cy - m.y1)))
            if near and min(abs(b.cy - near[0].y0), abs(b.cy - near[0].y1)) <= a.near_tol * H:
                m, flag = near[0], "y_near"
            else:
                rep["orphans"].append({"text": b.text, "id": b.id, "bbox": b.bb(),
                                       "why": "no_main_block", "role": c["role"]})
                continue
        ov = (min(b.x1, c["x1"]) - max(b.x0, c["x0"])) / b.w if b.w else 1
        assigned[m.id].append((b, c, [f for f in (flag, "weak_col" if ov < 0.5 else None) if f]))

    med = statistics.median([b.h for b, _ in cells]) if cells else 0
    used_fam = set()

    # 5. сборка блоков и строк
    for g in sorted(mainb, key=lambda g: mainb[g][0].cx):
        for no, m in enumerate(mainb[g], 1):
            bl = {"block_no": no, "sex": g, "block_id": m.id, "text_block_id": m.tid,
                  "bbox": m.bb(), "names_text": m.text, "flags": [], "rows": []}
            if m.id in nested:
                bl["flags"].append("nested_block")

            # номер семьи: последний номер выше (или чуть ниже верха) блока
            fl = [f for f in fam.get(g, []) if f.cy <= m.y0 + a.fam_tol * H]
            if fl:
                f = fl[-1]
                used_fam.add(f.id)
                bl.update(family=f.text, family_no=to_int(f.text), family_cell_id=f.id)
            else:
                bl.update(family=None, family_no=None)
                bl["flags"].append("family_missing")
            if any(m.y0 + a.fam_tol * H < f.cy <= m.y1 for f in fam.get(g, [])):
                bl["flags"].append("family_changes_inside_block")

            items = assigned[m.id]
            normal = sorted([t for t in items if t[0].h <= a.tall * med], key=lambda t: t[0].cy)
            tall = [t for t in items if t[0].h > a.tall * med]
            rows = []
            for b, c, fls in normal:
                last = rows[-1] if rows else None
                if last:
                    ov = min(b.y1, last["y1"]) - max(b.y0, last["y0"])
                    if ov > 0 and ov / min(b.h, last["y1"] - last["y0"]) >= a.row_overlap:
                        r = last
                    else:
                        r = None
                else:
                    r = None
                if r is None:
                    r = {"y0": b.y0, "y1": b.y1, "cells": {}, "flags": []}
                    rows.append(r)
                r["y0"], r["y1"] = min(r["y0"], b.y0), max(r["y1"], b.y1)
                r["cells"].setdefault(c["role"], []).append(b)
                r["flags"] += [f"{x}:{c['role']}" for x in fls]
            for b, c, fls in tall:  # высокая ячейка («новорожденные.») покрывает несколько строк
                hit = [r for r in rows
                       if (min(b.y1, r["y1"]) - max(b.y0, r["y0"])) / max(r["y1"] - r["y0"], 1) >= 0.4]
                if not hit:
                    hit = [{"y0": b.y0, "y1": b.y1, "cells": {}, "flags": []}]
                    rows.append(hit[0])
                for r in hit:
                    if c["role"] in r["cells"]:
                        r["flags"].append(f"span_conflict:{c['role']}")
                    else:
                        r["cells"][c["role"]] = [b]
                        r["flags"].append(f"span:{c['role']}")
            rows.sort(key=lambda r: r["y0"])
            for k, r in enumerate(rows, 1):
                out = {"row_no": k, "y": [r["y0"], r["y1"]], "flags": list(r["flags"]), "cells": {}}
                for role, bs in r["cells"].items():
                    out["cells"][role] = cell_json(bs)
                    if len(bs) > 1:
                        out["flags"].append(f"multi_cell:{role}")
                bl["rows"].append(out)
            bl["rows_in_block"] = len(rows)
            rep["blocks"].append(bl)

    for g, fs in fam.items():
        for f in fs:
            if f.id not in used_fam:
                rep["flags"].append(f"family_cell_unused:{f.text}")

    # 6. проверка по итогам внизу страницы: строки без «Когда именно» (выбыл) == «Итого налицо»
    chk = {}
    for sex, total in rep["totals"].items():
        rows_all = [r for bl in rep["blocks"] if bl["sex"] == sex for r in bl["rows"]]
        present = [r for r in rows_all if "departed_when" not in r["cells"]]
        chk[sex] = {"expected": total, "rows_total": len(rows_all),
                    "rows_present": len(present), "ok": total == len(present)}
        if total != len(present):
            rep["flags"].append(f"totals_mismatch:{sex}:expected={total},rows={len(present)}")
    rep["totals_check"] = chk
    if rep["orphans"]:
        rep["flags"].append(f"orphans:{len(rep['orphans'])}")
    return rep


# ----------------------------------------------------------------------------- выход
def csv_rows(rep):
    for bl in rep["blocks"]:
        for r in bl["rows"]:
            row = {"page": rep["page"], "node_id": rep["node_id"], "sex": bl["sex"],
                   "block_no": bl["block_no"], "block_id": bl["block_id"],
                   "family": bl["family"], "family_no": bl["family_no"],
                   "row_no": r["row_no"], "rows_in_block": bl["rows_in_block"],
                   "names_text": bl["names_text"]}
            for role in ("age_prev", "age_mid", "departed_when", "absent_since", "age_now"):
                row[role] = r["cells"].get(role, {}).get("text", "")
            row["flags"] = ";".join(bl["flags"] + r["flags"])
            yield row


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+", help="файлы .json/.html или папки")
    ap.add_argument("--out-json", default="pages.json")
    ap.add_argument("--out-csv", default="rows.csv")
    ap.add_argument("--split-x", type=float, default=None, help="x границы левой/правой страницы разворота (по умолчанию W/2)")
    ap.add_argument("--fam-tol", type=float, default=0.025, help="допуск по y для номера семьи, доля высоты")
    ap.add_argument("--near-tol", type=float, default=0.01, help="допуск по y для ячейки вне блока, доля высоты")
    ap.add_argument("--col-tol", type=float, default=0.05, help="макс. расстояние до столбца по x, доля ширины")
    ap.add_argument("--tall", type=float, default=1.9, help="ячейка выше медианной в N раз считается «высокой»")
    ap.add_argument("--row-overlap", type=float, default=0.5, help="мин. перекрытие по y для одной строки")
    a = ap.parse_args()

    files = []
    for p in a.inputs:
        p = Path(p)
        files += sorted(f for f in p.rglob("*") if f.suffix in (".json", ".html")) if p.is_dir() else [p]

    pages, n_rows = [], 0
    with open(a.out_csv, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=CSV_COLS)
        w.writeheader()
        for f in files:
            try:
                rep = process(load_node(f), a)
            except Exception as e:  # одна битая страница не должна останавливать пакет
                rep = {"page": None, "source": str(f), "flags": [f"error:{type(e).__name__}:{e}"], "blocks": []}
            rep["source"] = str(f)
            pages.append(rep)
            for row in csv_rows(rep):
                w.writerow(row)
                n_rows += 1
            rows = sum(b["rows_in_block"] for b in rep["blocks"])
            print(f"стр. {rep.get('page')}: блоков {len(rep['blocks'])}, строк {rows}, "
                  f"флаги: {rep['flags'] or '-'}")
    Path(a.out_json).write_text(json.dumps(pages, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"готово: {len(pages)} стр., {n_rows} строк -> {a.out_csv}, {a.out_json}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
