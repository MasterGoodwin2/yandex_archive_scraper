#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сбор данных о людях "моей округи" из исповедных ведомостей Яндекс Архива (версия 1.3).

Принцип: ничего не менять, не достраивать, не удалять. Текст записей сохраняется как распознан;
имена, отчества, связи НЕ разбираются.

Порядок работы:
  1. Запрос HTML-файла со ссылками (list_links.html).
  2. Запрос места сохранения Excel.
  3. Вход в Яндекс (логин/пароль вводите вы сами в окне браузера).
  4. Для каждой ссылки:
     - год = текст ссылки; ya.ru и yandex.ru считаются одним и тем же сайтом;
     - диапазон = sheet_page_from .. sheet_page_to из адреса;
     - проверка диапазона: просматриваются страницы диапазона +-TOLERANCE;
       если найдено меньше 4 из 6 населённых пунктов округи — ссылка пропускается;
     - иначе собираются дворы: № двора ↔ № мужчины/женщины ↔ запись ↔ возраст (по координатам блоков).
  5. Excel: лист "Люди" (одна строка = один двор, все люди двора в одной ячейке), "Страницы", "Ссылки".

Установка:
  python3 -m pip install playwright beautifulsoup4 openpyxl
  (браузер: запущенный Chrome с портом 9222 или установленный Google Chrome)
"""

import json
import os
import random
import re
import sys
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from urllib.parse import parse_qs, urlparse

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from playwright.sync_api import sync_playwright

# ----------------------------------------------------------------------------
# НАСТРОЙКИ
# ----------------------------------------------------------------------------
LOGIN_URL = (
    "https://passport.yandex.ru/pwl-yandex/auth/password"
    "?retpath=https%3A%2F%2Fya.ru%2Farchive%2F&backpath=https%3A%2F%2Fya.ru%2Farchive%2F"
)
ARCHIVE_HOSTS = ("ya.ru", "yandex.ru")      # равнозначные домены
WORK_HOST = "https://ya.ru"
PAGE_URL_TEMPLATE = WORK_HOST + "/archive/catalog/{file_id}/{page}"

TOLERANCE = 3          # диапазон неточный: +-2-3 страницы с каждого края
MIN_SETTLEMENTS = 4    # минимум населённых пунктов округи в диапазоне
DEFAULT_SPAN = 15      # если в ссылке нет sheet_page_to: from .. from+DEFAULT_SPAN

_PREFIX = r"(?:\bсел\w*|\bдер\w*|\bд\.|\bс\.)\s*[\[\(«\"]?\s*"
SETTLEMENTS = {
    "село Троицкое":       _PREFIX + r"троицк",
    "деревня Острогина":   _PREFIX + r"острогин",
    "деревня Черепкова":   _PREFIX + r"черепков",
    "деревня Рублева":     _PREFIX + r"рублев",
    "деревня Луки":        _PREFIX + r"лук(?:и)?\b",
    "деревня Мякинина":    _PREFIX + r"(?:м[яе]кинин|мемин)",
}
OTHER_HEADER = re.compile(
    r"(?:\bсел\w*|\bдер\w*|\bд\.|\bс\.)\s+[\[\(«\"]?([А-ЯЁІѢ][а-яёіѣъь]{3,})"
)
PREFIX_ONLY = re.compile(r"^\s*(?:сел\w*|дер\w*|[дс])\.?\s*$", re.I)
YARD_A = re.compile(r"^\s*(?:двор\w*|дв)\.?\s*(?:№|N)?\s*(\d{1,3})(?!\w)", re.I)
YARD_B = re.compile(r"^\s*(?:№|N)\s*(\d{1,3})\s*[.:]?\s*$")
HEADER_MAX_LEN = 70

DELAY_RANGE = (1.0, 2.5)
PAGE_TIMEOUT_MS = 45_000
MAX_RETRIES = 2
RENDER_WAIT_MS = 15_000
LOGIN_WAIT_SEC = 600
DEBUG_DIR = "debug_yandex_archive"   # сюда сохранятся HTML, блоки и ответы запросов первой страницы
# ----------------------------------------------------------------------------


def norm(s: str) -> str:
    s = s.lower()
    for a, b in (("ё", "е"), ("ѣ", "е"), ("і", "и"), ("ѵ", "и"), ("ѳ", "ф")):
        s = s.replace(a, b)
    return s.replace("ъ", "").replace("ь", "")


# ============================================================================
# JS, выполняемый внутри страницы: текст окна + блоки с координатами
# ============================================================================
EXTRACT_JS = r"""
async ({maxWait, wantDebug}) => {
  const rectOf = (e) => e.getBoundingClientRect();
  const visible = (e) => {
    const cs = getComputedStyle(e);
    if (cs.display === 'none' || cs.visibility === 'hidden') return false;
    const r = rectOf(e);
    return r.width > 0 && r.height > 0;
  };
  const findViewer = () => {
    const list = [...document.querySelectorAll('canvas, img, svg')].filter((e) => visible(e))
      .map((e) => ({ e, r: rectOf(e) })).filter((x) => x.r.width > 200 && x.r.height > 200)
      .sort((a, b) => b.r.width * b.r.height - a.r.width * a.r.height);
    return list.length ? list[0].r : null;
  };
  const autoText = () => {
    const vr = findViewer();
    const minLeft = vr ? vr.right - 20 : window.innerWidth * 0.4;
    const skip = new Set(['SCRIPT','STYLE','CANVAS','IMG','SVG','NOSCRIPT','HEAD','HTML','BODY']);
    const cands = [];
    for (const e of document.body.querySelectorAll('*')) {
      if (skip.has(e.tagName.toUpperCase())) continue;
      const r = rectOf(e);
      if (r.width < 100 || r.height < 60 || r.left < minLeft) continue;
      if (!visible(e)) continue;
      const t = e.innerText || '';
      if (t.trim().length < 30) continue;
      cands.push({ e, len: t.length, area: r.width * r.height });
    }
    if (!cands.length) return null;
    const maxLen = Math.max(...cands.map((c) => c.len));
    const good = cands.filter((c) => c.len >= maxLen * 0.85).sort((a, b) => a.area - b.area);
    return good[0].e;
  };
  const getText = () => {
    const el = autoText();
    return el ? { found: true, text: el.innerText || '', el } : { found: false, text: '', el: null };
  };
  const t0 = Date.now();
  let last = -1, stable = 0, res = getText();
  while (Date.now() - t0 < maxWait) {
    res = getText();
    const len = res.text.trim().length;
    if (res.found && len > 0 && len === last) { if (++stable >= 3) break; } else stable = 0;
    if (res.found && len === 0 && Date.now() - t0 > 6000) break;
    last = len;
    await new Promise((r) => setTimeout(r, 600));
  }
  const collectBlocks = (roots, region) => {
    const out = [], seen = new Set();
    const sx = window.scrollX, sy = window.scrollY;
    const push = (t, r) => {
      t = t.replace(/\s+/g, ' ').trim();
      if (!t || r.width <= 0 || r.height <= 0) return;
      if (region && (r.right < region.left - 4 || r.left > region.right + 4 || r.bottom < region.top - 4 || r.top > region.bottom + 4)) return;
      const b = { t, x: Math.round((r.left + sx) * 10) / 10, y: Math.round((r.top + sy) * 10) / 10,
                  w: Math.round(r.width * 10) / 10, h: Math.round(r.height * 10) / 10 };
      const key = t + '|' + Math.round(b.x) + '|' + Math.round(b.y);
      if (seen.has(key)) return;
      seen.add(key); out.push(b);
    };
    for (const root of roots) {
      if (!root) continue;
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      let n;
      while ((n = walker.nextNode()) && out.length < 4000) {
        const val = n.nodeValue || '';
        if (!val.trim()) continue;
        const pe = n.parentElement;
        if (!pe) continue;
        const tag = pe.tagName.toUpperCase();
        if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'NOSCRIPT') continue;
        if (!visible(pe)) continue;
        const rg = document.createRange();
        rg.selectNodeContents(n);
        const rects = [...rg.getClientRects()].filter((r) => r.width > 0 && r.height > 0);
        if (!rects.length) continue;
        if (rects.length === 1 || !/\s/.test(val.trim())) { push(val, rg.getBoundingClientRect()); continue; }
        const lines = [];
        const rx = /\S+/g;
        let m;
        while ((m = rx.exec(val))) {
          const wr = document.createRange();
          wr.setStart(n, m.index); wr.setEnd(n, m.index + m[0].length);
          const r = wr.getBoundingClientRect();
          if (r.width <= 0) continue;
          const l = lines[lines.length - 1];
          if (l && Math.abs(r.top - l.top) < r.height * 0.5) { l.t += ' ' + m[0]; l.left = Math.min(l.left, r.left); l.right = Math.max(l.right, r.right); }
          else lines.push({ t: m[0], left: r.left, right: r.right, top: r.top, height: r.height });
        }
        for (const l of lines) push(l.t, { left: l.left, right: l.right, top: l.top, bottom: l.top + l.height, width: l.right - l.left, height: l.height });
      }
    }
    return out;
  };
  const vr = findViewer();
  const textBlocks = res.found ? collectBlocks([res.el], null) : [];
  const overlayBlocks = vr ? collectBlocks([document.body], vr) : [];
  let debug = null;
  if (wantDebug) {
    const resources = performance.getEntriesByType('resource')
      .filter((e) => ['fetch', 'xmlhttprequest'].includes(e.initiatorType)).map((e) => e.name);
    const json = [];
    let budget = 600000;
    for (const u of [...new Set(resources)].filter((x) => /^https:\/\/(?:ya\.ru|yandex\.ru)\//.test(x)).slice(0, 20)) {
      try {
        const r = await fetch(u, { credentials: 'include' });
        const t = await r.text();
        if (!/^\s*[\[{]/.test(t)) continue;
        const part = t.slice(0, Math.min(t.length, 200000, budget));
        budget -= part.length;
        json.push({ url: u, status: r.status, truncated: part.length < t.length, body: part });
        if (budget <= 0) break;
      } catch (e) {}
    }
    debug = { viewer: vr ? { left: vr.left, top: vr.top, width: vr.width, height: vr.height } : null,
              textBlocks, overlayBlocks, resources, json };
  }
  return { found: res.found, text: res.text, textBlocks, overlayBlocks, debug };
}
"""


# ============================================================================
# Сопоставление блоков: № двора ↔ № мужчины/женщины ↔ запись ↔ возраст
# (то же, что layout.js в расширении)
# ============================================================================
NUM_RX = re.compile(r"^(?:№|N|n)?\s*(\d{1,3})\s*[.,)]?$")
LEAD_NUM_RX = re.compile(r"^(\d{1,3})\s*[.)]\s*(\D.*)$")
TRAIL_AGE_RX = re.compile(r"^(.*[^\d\s№])\s+(\d{1,3})$")
MAX_AGE = 110

ARTIFACT = re.compile("^(?:" + "|".join([
    "лета", "семьи", "семья", "ныне", "нине", "на лицо", "налицо", "крестьян", "мужеск", "женск", "мужск", "итого",
    "по последне", "по после", "ис того", "из того", "во временно", "с которого", "скоторого", "скоторо", "когда",
    "кто были", "которые", "кто же", "ктож", "показани", "числа", "состояли", "состояло", "выбыли", "подано",
    "ревизск", "ревизия", "сказка", "оборот", "попов$", "исповедь", "у исповеди", "и святого", "исповедались",
    "токмо", "причастились", "домов", "двор$", "число$", "людеи", "людей", "мужеска", "женска", "всего", "прибыло",
    "убыло", "выбыло", "в том числе", "с прибылыми", "по ревизии", "во время", "родилось", "умерло", "имена",
    "и отчества", "пол$", "номер", "возраст", "лет$",
]) + ")")
ALWAYS_ARTIFACT = re.compile(r"^(?:итого|всего|по ревизии|в том числе)")
TITLE_NOISE = re.compile(r"(?:ревизск|исповедн|ведомост|сказк|уезд|губерн|волост|приход)")


def kline(s):
    return re.sub(r"\s+", " ", re.sub(r"[^а-яa-z0-9 ]", " ", norm(s))).strip()


def match_header(line, nxt=""):
    """('own', имя) / ('other', имя) / None."""
    # длинная строка-«титул» («…Московской Губерніи и уѣзда Села троицкаго Здеревнями…»): берём первое село/деревню
    if len(line) > HEADER_MAX_LEN and re.search(r"(?:уезд|губерн)", norm(line)):
        n = norm(line)
        if re.search(r"(?:уезда|губернии)\s+(?:и\s+уезда\s+)?(?:сел[а-я]*|дер[а-я]*)\.?\s*[\[\(«\"]?\s*(?:[а-я]+\s+){0,1}", n):
            best, best_pos = None, 10 ** 9
            for name, rx in SETTLEMENTS.items():
                mm = re.search(rx, n)
                if mm and mm.start() < best_pos:
                    best, best_pos = name, mm.start()
            if best:
                return "own", best
    cand = [line]
    # заголовок мог разорваться на две строки: "Деревня" / "Острогина" — склеиваем ТОЛЬКО если текущая строка — одно слово-префикс
    if nxt and PREFIX_ONLY.match(line):
        cand.append(line + " " + nxt)
    for c in cand:
        if len(c) > HEADER_MAX_LEN:
            continue
        n = norm(c)
        for name, rx in SETTLEMENTS.items():
            if re.search(rx, n):
                return "own", name
        m = OTHER_HEADER.search(c)
        if m:
            return "other", m.group(1)
    return None


def match_yard(text):
    m = YARD_A.match(text)
    if m:
        rest = re.sub(r"^[\s.:,;-]+", "", text[m.end():]).strip()
        return int(m.group(1)), rest
    m = YARD_B.match(text)
    if m:
        return int(m.group(1)), ""
    return None


def find_settlements(text):
    t = norm(text)
    return {name for name, rx in SETTLEMENTS.items() if re.search(rx, t)}


def _median(vals, dflt):
    if not vals:
        return dflt
    v = sorted(vals)
    return v[len(v) // 2]


class State:
    def __init__(self):
        self.current = None
        self.other = None
        self.fam = None
        self.cols = {"m": None, "f": None, "yard": None}
        self.dropped = 0
        self.block_pages = 0
        self.line_pages = 0
        self.src = None


def _flush(st, out):
    if st.fam and st.fam["persons"]:
        out.append(st.fam)
    st.fam = None


def _start_family(st, out, yard, ctx):
    _flush(st, out)
    st.fam = {"settlement": ctx["settlement"], "yard": "" if yard is None else yard,
              "persons": [], "notes": [], "url": ctx["url"], "page": ctx["page"]}


def _ensure_family(st, out, ctx):
    if st.fam is None or st.fam["settlement"] != ctx["settlement"]:
        _start_family(st, out, None, ctx)
    return st.fam


def _note(fam, msg):
    if msg not in fam["notes"]:
        fam["notes"].append(msg)


def _apply_header(st, out, h):
    _flush(st, out)
    kind, name = h
    st.current = name if kind == "own" else None
    st.other = name if kind == "other" else None


def _settlement_of(st, inside):
    if st.current:
        return st.current
    if st.other:
        return None
    return "не определён (начало диапазона)" if inside else None


def atomize(b):
    t = re.sub(r"\s+", " ", str(b.get("t", ""))).strip()
    if not t:
        return []

    def mk(kind, text, x, w, v=None):
        return {"kind": kind, "t": text, "x": x, "w": w, "cx": x + w / 2, "v": v,
                "y": b["y"], "h": b["h"], "used": False, "dead": False}

    m = NUM_RX.match(t)
    if m:
        return [mk("num", t, b["x"], b["w"], int(m.group(1)))]
    m = LEAD_NUM_RX.match(t)
    if m:
        wn = b["w"] * len(m.group(1)) / len(t)
        return [mk("num", m.group(1), b["x"], wn, int(m.group(1))),
                mk("text", m.group(2).strip(), b["x"] + wn, b["w"] - wn)]
    m = TRAIL_AGE_RX.match(t)
    if m and 0 < int(m.group(2)) <= MAX_AGE:
        wn = b["w"] * len(m.group(2)) / len(t)
        return [mk("text", m.group(1).strip(), b["x"], b["w"] - wn),
                mk("num", m.group(2), b["x"] + b["w"] - wn, wn, int(m.group(2)))]
    return [mk("text", t, b["x"], b["w"])]


def cluster_rows(atoms):
    bs = sorted(atoms, key=lambda a: a["y"] + a["h"] / 2)
    tol = max(3, _median([a["h"] for a in bs], 10) * 0.5)
    rows = []
    for a in bs:
        cy = a["y"] + a["h"] / 2
        if rows and abs(cy - rows[-1]["cy"]) <= tol:
            last = rows[-1]
            last["atoms"].append(a)
            last["cy"] = (last["cy"] * (len(last["atoms"]) - 1) + cy) / len(last["atoms"])
        else:
            rows.append({"cy": cy, "atoms": [a]})
    for r in rows:
        r["atoms"].sort(key=lambda a: a["x"])
    return rows


def process_blocks(blocks, st, ctx, inside, out):
    atoms = []
    for b in blocks:
        atoms.extend(atomize(b))
    if not atoms:
        return
    rows = cluster_rows(atoms)
    min_x = min(a["x"] for a in atoms)
    max_x = max(a["x"] + a["w"] for a in atoms)
    col_tol = max(20, (max_x - min_x) * 0.04)

    for r in rows:
        for a in r["atoms"]:
            if a["kind"] != "text":
                continue
            k = kline(a["t"])
            if re.match(r"^(?:мужеск|мужск)", k):
                st.cols["m"] = a["cx"]
            elif re.match(r"^женск", k):
                st.cols["f"] = a["cx"]
            elif re.match(r"^(?:номер\s+)?(?:двор\w*|дом\w*)$", k):
                st.cols["yard"] = a["cx"]
    text_xs = sorted(a["x"] for a in atoms if a["kind"] == "text" and not re.match(r"^(?:мужеск|женск|мужск)", kline(a["t"])))
    text_left = text_xs[int(len(text_xs) * 0.1)] if text_xs else float("inf")

    def gender_for(cx):
        m, f = st.cols["m"], st.cols["f"]
        if m is not None and f is not None:
            return "м" if abs(cx - m) <= abs(cx - f) else "ж"
        return ""

    ri = 0
    while ri < len(rows):
        row = rows[ri]
        ri += 1
        ats = list(row["atoms"])
        joined = " ".join(a["t"] for a in ats)
        next_joined = " ".join(a["t"] for a in rows[ri]["atoms"]) if ri < len(rows) else ""

        h = match_header(joined, next_joined)
        if h:
            _apply_header(st, out, h)
            if PREFIX_ONLY.match(joined):
                ri += 1
            continue
        has_num = any(a["kind"] == "num" for a in ats)
        k = kline(joined)
        all_heads = all(ARTIFACT.match(kline(a["t"])) or TITLE_NOISE.search(kline(a["t"])) for a in ats)
        if not has_num and (ARTIFACT.match(k) or TITLE_NOISE.search(k) or all_heads):
            st.dropped += 1
            continue
        if ALWAYS_ARTIFACT.match(k):
            st.dropped += 1
            continue
        cur = _settlement_of(st, inside)
        if cur is None:
            continue
        ctx["settlement"] = cur

        yard_no = None
        for a in ats:
            if a["kind"] != "text":
                continue
            y = match_yard(a["t"])
            if y:
                yard_no = y[0]
                if y[1]:
                    a["t"] = y[1]
                else:
                    a["dead"] = True
                break
        ats = [a for a in ats if not a["dead"]]

        def free_nums():
            return [a for a in ats if a["kind"] == "num" and not a["used"]]

        if yard_no is None:
            cand = None
            if st.cols["yard"] is not None:
                cand = next((a for a in free_nums() if abs(a["cx"] - st.cols["yard"]) <= col_tol), None)
            else:
                texts_ = [a for a in ats if a["kind"] == "text"]
                first_text = texts_[0]["x"] if texts_ else float("inf")
                left = [a for a in free_nums() if a["cx"] < first_text]
                if len(left) >= 2:
                    cand = left[0]
                elif not texts_ and len(left) == 1 and left[0]["x"] < text_left - col_tol:
                    cand = left[0]
            if cand:
                yard_no = cand["v"]
                cand["used"] = True
        if yard_no is not None:
            _start_family(st, out, yard_no, ctx)

        texts = [a for a in ats if a["kind"] == "text"]
        if not texts:
            lone = free_nums()
            if lone and yard_no is None:
                fam = _ensure_family(st, out, ctx)
                _note(fam, "число без записи рядом (стр. %s): %s" % (ctx["page"], ", ".join(str(a["v"]) for a in lone)))
            continue
        fam = _ensure_family(st, out, ctx)
        people = [{"gender": gender_for(a["cx"]), "no": "", "text": a["t"], "age": "", "atom": a} for a in texts]

        gaps = [[] for _ in people]
        tail = []
        for n in free_nums():
            idx = next((i for i, p in enumerate(people) if n["cx"] < p["atom"]["cx"]), -1)
            (tail if idx == -1 else gaps[idx]).append(n)
        unmatched = []

        def set_age(p, n):
            if 0 < n["v"] <= MAX_AGE:
                p["age"] = n["v"]
                return True
            return False

        for i, p in enumerate(people):
            g = gaps[i]
            prev = people[i - 1] if i > 0 else None
            own, prev_age = [], None
            if prev is None:
                own = g[-1:]
                unmatched.extend(g[:-1])
            elif len(g) == 1:
                n = g[0]
                d_prev = n["cx"] - (prev["atom"]["x"] + prev["atom"]["w"])
                d_next = p["atom"]["x"] - n["cx"]
                if prev["age"] == "" and d_prev <= d_next:
                    prev_age = n
                else:
                    own = [n]
            elif len(g) >= 2:
                prev_age = g[0]
                own = [g[-1]]
                unmatched.extend(g[1:-1])
            if prev_age is not None and not set_age(prev, prev_age):
                unmatched.append(prev_age)
            if own:
                p["no"] = own[0]["v"]
        if tail:
            last = people[-1]
            if last["age"] == "" and set_age(last, tail[0]):
                unmatched.extend(tail[1:])
            else:
                unmatched.extend(tail)
        for p in people:
            fam["persons"].append({"gender": p["gender"], "no": p["no"], "text": p["text"], "age": p["age"]})
            if not p["gender"]:
                _note(fam, "пол не определён: нет шапок «Мужеск.»/«Женск.» (стр. %s)" % ctx["page"])
        if unmatched:
            _note(fam, "несопоставленные числа (стр. %s): %s" % (ctx["page"], ", ".join(str(a["v"]) for a in unmatched)))


def process_lines(text, st, ctx, inside, out):
    """Запасной режим: координат нет — записи строками как есть, № двора по маркерам; № и возраст не сопоставляются."""
    lines = [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]
    i = 0
    while i < len(lines):
        ln = lines[i]
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        i += 1
        h = match_header(ln, nxt)
        if h:
            _apply_header(st, out, h)
            if PREFIX_ONLY.match(ln):
                i += 1
            continue
        y = match_yard(ln)
        cur = _settlement_of(st, inside)
        if cur is None:
            continue
        ctx["settlement"] = cur
        if y:
            _start_family(st, out, y[0], ctx)
            if not y[1]:
                continue
        body = y[1] if y else ln
        if len(body) < 3:
            continue
        k = kline(body)
        has_num = bool(re.search(r"\d", body))
        if (not has_num and (ARTIFACT.match(k) or TITLE_NOISE.search(k))) or ALWAYS_ARTIFACT.match(k):
            st.dropped += 1
            continue
        fam = _ensure_family(st, out, ctx)
        _note(fam, "координаты блоков не получены: № и возраст не сопоставлены, записи даны строками как есть")
        fam["persons"].append({"gender": "", "no": "", "text": body, "age": ""})


def is_table(blocks, relaxed=False):
    if relaxed:
        return len(blocks) >= 4 and len({round((b["y"] + b["h"] / 2) / 8) for b in blocks}) >= 2
    if len(blocks) < 8:
        return False
    if len({round((b["y"] + b["h"] / 2) / 8) for b in blocks}) < 3:
        return False
    xs = sorted(b["x"] for b in blocks)
    gap = max(20, (xs[-1] - xs[0] + 1) * 0.04)
    cols, run = 0, 1
    for i in range(1, len(xs) + 1):
        if i < len(xs) and xs[i] - xs[i - 1] <= gap:
            run += 1
        else:
            if run >= 2:
                cols += 1
            run = 1
    return cols >= 3


def process_page(page, st, inside, out):
    """page = {candidates:[(name, blocks)], text, url, page}. Возвращает, как собрана страница."""
    ctx = {"url": page["url"], "page": page["page"], "settlement": None}
    for name, blocks in page["candidates"]:
        blocks = [b for b in blocks if str(b.get("t", "")).strip()]
        if is_table(blocks, st.src == name):
            st.src = name
            st.block_pages += 1
            process_blocks(blocks, st, ctx, inside, out)
            return "блоки: " + name
    st.line_pages += 1
    process_lines(page["text"], st, ctx, inside, out)
    return "строки (координат нет)"


def format_person(p):
    label = (p["gender"] or "") + ("№%s" % p["no"] if p["no"] != "" else "")
    return (label + " " if label else "") + p["text"] + (", возр. %s" % p["age"] if p["age"] != "" else "")


def family_cell(fam):
    return "\n".join(format_person(p) for p in fam["persons"])


# ============================================================================
# Ввод/вывод
# ============================================================================
def ask_paths():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    html_path = filedialog.askopenfilename(
        title="Выберите HTML-файл со списком ссылок (list_links.html)",
        filetypes=[("HTML", "*.html *.htm"), ("Все файлы", "*.*")])
    if not html_path:
        sys.exit("Файл со ссылками не выбран.")
    xlsx_path = filedialog.asksaveasfilename(
        title="Куда сохранить результат (Excel)", defaultextension=".xlsx",
        initialfile="result.xlsx", filetypes=[("Excel", "*.xlsx")])
    if not xlsx_path:
        sys.exit("Путь сохранения не выбран.")
    root.destroy()
    return html_path, xlsx_path


def _int(s):
    try:
        return int(s)
    except (TypeError, ValueError):
        return None


def parse_links(html_path):
    """Возвращает список словарей: url, year, file_id, p_from, p_to, note."""
    with open(html_path, "r", encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    items, seen = [], set()
    for a in soup.find_all("a", href=True):
        url = a["href"].strip()
        u = urlparse(url)
        if u.hostname not in ARCHIVE_HOSTS or not u.path.startswith("/archive"):
            continue
        if url in seen:
            continue
        seen.add(url)

        # Год: текст всего абзаца — в списке бывают "разорванные" ссылки вида <a>186</a>4.
        holder = a.find_parent("p") or a
        m = re.search(r"(?<!\d)(1[5-9]\d\d)(?!\d)", holder.get_text("", strip=True))
        year = int(m.group(1)) if m else None

        q = parse_qs(u.query)
        note = ""
        file_id, p_from, p_to = None, None, None

        if u.path.startswith("/archive/search"):
            file_id = (q.get("fileId") or [None])[0]
            p_from = _int((q.get("sheet_page_from") or [""])[0])
            p_to = _int((q.get("sheet_page_to") or [""])[0])
        elif u.path.startswith("/archive/catalog/"):
            parts = [x for x in u.path.split("/") if x]
            if len(parts) >= 4:
                file_id = parts[2]
                p_from = p_to = _int(parts[3])
                note = "ссылка на одну страницу каталога, диапазона в ссылке нет"

        if p_from is not None and p_to is None:
            p_to = p_from + DEFAULT_SPAN
            note = f"нет sheet_page_to — взято from+{DEFAULT_SPAN}"
        if p_from is not None and p_to is not None and p_from > p_to:
            p_from, p_to = p_to, p_from

        items.append(dict(url=url, year=year, file_id=file_id, p_from=p_from, p_to=p_to, note=note))
    return items


def wait_for_login(page):
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    print("\n>>> В окне браузера введите логин и пароль Яндекса.")
    print(f">>> Жду входа до {LOGIN_WAIT_SEC // 60} мин...\n")
    deadline = time.time() + LOGIN_WAIT_SEC
    while time.time() < deadline:
        if "passport.yandex" not in page.url and "ya.ru" in page.url:
            break
        time.sleep(1)
    else:
        sys.exit("Время ожидания входа истекло.")
    time.sleep(1.5)
    try:
        page.goto("https://yandex.ru/archive", wait_until="domcontentloaded", timeout=30_000)
        page.goto(WORK_HOST + "/archive", wait_until="domcontentloaded", timeout=30_000)
    except Exception:
        pass
    print("Вход выполнен.")


_debug_saved = False


def fetch_sheet(page, file_id, n):
    """Открывает страницу n документа, возвращает dict: url, text, candidates, err."""
    global _debug_saved
    url = PAGE_URL_TEMPLATE.format(file_id=file_id, page=n)
    last_err = None
    for _ in range(MAX_RETRIES + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
            if "captcha" in page.url:
                input("\nЯндекс просит капчу. Решите её в браузере и нажмите Enter...")
            if "passport.yandex" in page.url:
                input("\nСессия сбросилась. Войдите заново в браузере и нажмите Enter...")
                page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
            want_debug = not _debug_saved
            res = page.evaluate(EXTRACT_JS, {"maxWait": RENDER_WAIT_MS, "wantDebug": want_debug})
            if want_debug:
                os.makedirs(DEBUG_DIR, exist_ok=True)
                with open(os.path.join(DEBUG_DIR, "first_page.html"), "w", encoding="utf-8") as f:
                    f.write(page.content())
                with open(os.path.join(DEBUG_DIR, "first_page_debug.json"), "w", encoding="utf-8") as f:
                    json.dump(res.get("debug"), f, ensure_ascii=False)
                page.screenshot(path=os.path.join(DEBUG_DIR, "first_page.png"), full_page=True)
                _debug_saved = True
            return {
                "url": url, "text": res.get("text", ""),
                "candidates": [("окно текста", res.get("textBlocks") or []), ("скан", res.get("overlayBlocks") or [])],
                "err": "" if res.get("found") else "окно распознанного текста не найдено",
            }
        except Exception as e:  # noqa: BLE001
            last_err = type(e).__name__
            time.sleep(2)
    return {"url": url, "text": "", "candidates": [], "err": last_err or "ошибка"}


def save_workbook(path, families, sheets, links):
    wb = Workbook()
    bold = Font(bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")

    ws = wb.active
    ws.title = "Люди"
    ws.append(["Год", "Населённый пункт", "Двор №", "Люди двора (№ м/ж, как распознано, возраст)",
               "Примечание", "Ссылка на страницу"])
    for f in families:
        ws.append([f["year"], f["settlement"], f["yard"], _c(family_cell(f)), "; ".join(f["notes"]), f["url"]])
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        row[0].alignment = wrap

    ws2 = wb.create_sheet("Страницы")
    ws2.append(["Год", "Найденные НП округи", "Ссылка", "Распознанный текст страницы"])
    for r in sheets:
        ws2.append([r["year"], r["found"], r["url"], _c(r["text"])])

    ws3 = wb.create_sheet("Ссылки")
    ws3.append(["Год", "Исходная ссылка", "Диапазон из ссылки", "Просмотрено страниц", "Найдено НП (шт.)",
                "Какие НП", "Решение", "Дворов (семей)", "Людей", "Как собрано", "Примечание"])
    for r in links:
        ws3.append([_c(x) for x in r])

    for sh in (ws, ws2, ws3):
        sh.freeze_panes = "A2"
        for cell in sh[1]:
            cell.font = bold
        for col in sh.columns:
            w = max((len(str(x.value).split("\n")[0]) for x in col[:200] if x.value), default=8)
            sh.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 70)
    ws.column_dimensions["D"].width = 60
    wb.save(path)


def _c(v):
    v = "" if v is None else v
    return v[:32000] if isinstance(v, str) else v


def main():
    html_path, xlsx_path = ask_paths()
    items = parse_links(html_path)
    if not items:
        sys.exit("В файле не найдено ссылок Яндекс Архива.")
    print(f"Ссылок Яндекс Архива: {len(items)}")

    families, sheet_rows, link_rows = [], [], []
    page_cache = {}        # (file_id, page) -> результат fetch_sheet

    with sync_playwright() as pw:
        # 1) пробуем подключиться к уже запущенному Chrome (порт 9222), 2) иначе запускаем установленный Chrome
        try:
            browser = pw.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            print("Подключился к запущенному Chrome (порт 9222).")
        except Exception:
            print("Chrome с портом 9222 не найден — запускаю установленный Google Chrome...")
            browser = pw.chromium.launch(headless=False, channel="chrome")
            context = browser.new_context(locale="ru-RU")
        page = context.new_page()
        wait_for_login(page)

        for idx, it in enumerate(items, 1):
            rng = f"{it['p_from']}–{it['p_to']}" if it["p_from"] is not None else "—"
            tag = f"[{idx}/{len(items)}] {it['year']}"

            if not it["file_id"] or it["p_from"] is None:
                link_rows.append([it["year"], it["url"], rng, 0, 0, "", "пропущена: нет fileId или диапазона", 0, 0, "", it["note"]])
                print(f"{tag}: пропущена (нет fileId/диапазона)")
                continue

            lo = max(1, it["p_from"] - TOLERANCE)
            hi = it["p_to"] + TOLERANCE

            # --- 1. загрузка страниц окна и проверка диапазона ---
            sheets = {}
            for n in range(lo, hi + 1):
                key = (it["file_id"], n)
                if key not in page_cache:
                    page_cache[key] = fetch_sheet(page, it["file_id"], n)
                    time.sleep(random.uniform(*DELAY_RANGE))
                sheets[n] = page_cache[key]

            found_all, per_page = set(), {}
            for n, r in sheets.items():
                f = find_settlements(r["text"])
                per_page[n] = f
                found_all |= f

            if len(found_all) < MIN_SETTLEMENTS:
                link_rows.append([it["year"], it["url"], rng, len(sheets), len(found_all), "; ".join(sorted(found_all)),
                                  f"ПРОПУЩЕНА: найдено < {MIN_SETTLEMENTS} НП округи", 0, 0, "", it["note"]])
                print(f"{tag}: диапазон {rng} — найдено НП {len(found_all)} < {MIN_SETTLEMENTS}, пропуск")
                save_workbook(xlsx_path, families, sheet_rows, link_rows)
                continue

            # --- 2. сбор дворов ---
            st = State()
            fams, hows = [], set()
            for n in sorted(sheets):
                r = sheets[n]
                inside = it["p_from"] <= n <= it["p_to"]
                sheet_rows.append({"year": it["year"], "found": "; ".join(sorted(per_page[n])),
                                   "url": r["url"], "text": f"ОШИБКА: {r['err']}" if r["err"] else r["text"]})
                if r["err"]:
                    continue
                hows.add(process_page({"candidates": r["candidates"], "text": r["text"], "url": r["url"], "page": n},
                                      st, inside, fams))
            _flush(st, fams)
            for fam in fams:
                fam["year"] = it["year"] or ""
            families.extend(fams)
            persons = sum(len(f["persons"]) for f in fams)

            link_rows.append([it["year"], it["url"], rng, len(sheets), len(found_all), "; ".join(sorted(found_all)),
                              "ОБРАБОТАНА", len(fams), persons, "; ".join(sorted(hows)),
                              "; ".join(x for x in [it["note"], f"отброшено служебных строк: {st.dropped}" if st.dropped else ""] if x)])
            print(f"{tag}: диапазон {rng} — НП {len(found_all)}, дворов: {len(fams)}, людей: {persons}")
            save_workbook(xlsx_path, families, sheet_rows, link_rows)

        browser.close()

    save_workbook(xlsx_path, families, sheet_rows, link_rows)
    print(f"\nГотово. Дворов: {len(families)}. Файл: {xlsx_path}")
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showinfo("Готово", f"Сохранено: {xlsx_path}\nДворов: {len(families)}")
    except Exception:
        pass


if __name__ == "__main__":
    main()
