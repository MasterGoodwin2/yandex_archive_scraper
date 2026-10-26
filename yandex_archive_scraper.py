#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сбор данных о людях "моей округи" из исповедных ведомостей Яндекс Архива (версия 1.0).

РЕКОНСТРУКЦИЯ. Оригинальный файл версии 1.0 не сохранился. Скрипт восстановлен обратно
из расширения yandex-archive-extension 1.0.0 (logic.js + app.js): логика разбора перенесена
из JavaScript в Python один в один. Оболочка (вход в Яндекс, запуск браузера, окна выбора файлов,
запись Excel) в расширении другая, поэтому взята из версии 1.3.

Что делает версия 1.0:
  1. Запрос HTML-файла со ссылками (list_links.html).
  2. Запрос места сохранения Excel.
  3. Вход в Яндекс (логин/пароль вводите вы сами в окне браузера).
  4. Для каждой ссылки:
     - год = текст ссылки; ya.ru и yandex.ru считаются одним и тем же сайтом;
     - диапазон страниц берётся РОВНО из ссылки (sheet_page_from .. sheet_page_to);
     - загружаются страницы диапазона, текст берётся из блока распознанного текста;
     - строки под заголовком населённого пункта округи сохраняются как есть;
       числа 1..110 в строке попадают в колонку «Числа-возрасты (черновик)».
  5. Excel: лист "Люди" (одна строка = одна распознанная строка), "Страницы", "Ссылки".
     Файл обновляется после каждой ссылки.

Остановка: Ctrl+C — работа прерывается, собранное сохраняется.

Установка:
  python3 -m pip install playwright beautifulsoup4 openpyxl
  (браузер: запущенный Chrome с портом 9222 или установленный Google Chrome)
"""

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
from openpyxl.styles import Font
from playwright.sync_api import sync_playwright

# ----------------------------------------------------------------------------
# НАСТРОЙКИ
# ----------------------------------------------------------------------------
ARCHIVE_HOSTS = ("ya.ru", "yandex.ru")      # равнозначные домены
WORK_HOST = "https://ya.ru"
PAGE_URL_TEMPLATE = WORK_HOST + "/archive/catalog/{file_id}/{page}"

DEFAULT_SPAN = 15      # если в ссылке нет sheet_page_to: from .. from+DEFAULT_SPAN
HEADER_MAX_LEN = 70    # строка короче этого может быть заголовком населённого пункта
BIG_RANGE = 300        # предупреждение о слишком большом диапазоне

DELAY_RANGE = (1.0, 2.5)
PAGE_TIMEOUT_MS = 45_000
MAX_RETRIES = 2
RENDER_WAIT_MS = 15_000
DEBUG_HTML = True                      # сохранить HTML первой страницы (first_page.html)
DEBUG_DIR = "debug_yandex_archive"

# CSS-селекторы блока распознанного текста (по очереди; запасной — весь body)
TEXT_SELECTORS = [
    "[class*='ranscript']", "[class*='ecognized']", "[class*='ocr']",
    "[class*='structured']", "[class*='Structured']",
]
# ----------------------------------------------------------------------------

# Границы слов (в Python \w понимает кириллицу)
WB = r"(?<!\w)"
WE = r"(?!\w)"

# Обязателен префикс "сел/дер" (в т.ч. "дер." / "д." / "с.")
_PREFIX = rf"(?:{WB}(?:сел|дер)\w*\.?|{WB}[дс]\.)\s*[\[\(«\"]?\s*"
SETTLEMENTS = {
    "село Троицкое":     re.compile(_PREFIX + r"троицк"),
    "деревня Острогина": re.compile(_PREFIX + r"острогин"),
    "деревня Черепкова": re.compile(_PREFIX + r"черепков"),
    "деревня Рублева":   re.compile(_PREFIX + r"рублев"),
    "деревня Луки":      re.compile(_PREFIX + r"лук(?:и)?" + WE),
    "деревня Мякинина":  re.compile(_PREFIX + r"м[яе]кинин"),
}
# "Деревня"/"Дер." с любой буквы; сокращения "д."/"с." только строчные,
# чтобы не принять за заголовок инициалы ("Д. Иванов").
OTHER_HEADER = re.compile(
    rf"(?:{WB}(?:[Сс][Ее][Лл]|[Дд][Ее][Рр])\w*\.?|{WB}[дс]\.)\s*[\[\(«\"]?\s*([А-ЯЁІѢ][а-яёіѣъь]{{3,}})"
)
# строка, состоящая только из слова-префикса ("Деревня" / "Село" / "д.")
PREFIX_ONLY = re.compile(r"^\s*(?:(?:[Сс][Ее][Лл]|[Дд][Ее][Рр])\w*|[дсДС])\.?\s*$")


def norm(s: str) -> str:
    s = s.lower()
    for a, b in (("ё", "е"), ("ѣ", "е"), ("і", "и"), ("ѵ", "и"), ("ѳ", "ф")):
        s = s.replace(a, b)
    return s.replace("ъ", "").replace("ь", "")


# ============================================================================
# JS, выполняемый внутри страницы архива: достаёт текст блока распознанного текста
# ============================================================================
EXTRACT_JS = r"""
async ({selectors, maxWait, wantHtml}) => {
  const get = () => {
    for (const sel of selectors) {
      let els;
      try { els = [...document.querySelectorAll(sel)]; } catch (e) { continue; }
      // убираем вложенные дубли: если элемент лежит внутри другого найденного — пропускаем
      els = els.filter((e) => !els.some((o) => o !== e && o.contains(e)));
      const txt = els.map((e) => e.innerText || '').filter((s) => s.trim()).join('\n');
      if (txt.length > 200) return txt;
    }
    return document.body ? document.body.innerText : '';
  };
  const t0 = Date.now();
  let last = -1, stable = 0, text = '';
  while (Date.now() - t0 < maxWait) {
    text = get();
    if (text.length > 200 && text.length === last) {
      stable += 1;
      if (stable >= 2) break;
    } else {
      stable = 0;
    }
    last = text.length;
    await new Promise((r) => setTimeout(r, 600));
  }
  return { text, html: wantHtml ? document.documentElement.outerHTML : null };
}
"""


# ============================================================================
# Разбор
# ============================================================================
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


def find_settlements(text):
    t = norm(text)
    return {name for name, rx in SETTLEMENTS.items() if rx.search(t)}


def match_header(line, nxt=""):
    """('own', имя) / ('other', имя) / None."""
    cand = [line]
    # заголовок мог разорваться на две строки: "Деревня" / "Острогина".
    # Склеиваем ТОЛЬКО если текущая строка — одно слово-префикс.
    if nxt and PREFIX_ONLY.match(line):
        cand.append(line + " " + nxt)
    for c in cand:
        if len(c) > HEADER_MAX_LEN:
            continue
        n = norm(c)
        for name, rx in SETTLEMENTS.items():
            if rx.search(n):
                return "own", name
        m = OTHER_HEADER.search(c)
        if m:
            return "other", m.group(1)
    return None


def extract_people(sheet_text, state, strictly_inside=True):
    """state: {'current','other','started'} — переносится между страницами."""
    lines = [s.strip() for s in re.split(r"\r?\n", sheet_text) if s.strip()]
    people = []
    skip_next = False
    for i, ln in enumerate(lines):
        if skip_next:                       # вторая половина разорванного заголовка
            skip_next = False
            continue
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        h = match_header(ln, nxt)
        if h:
            kind, name = h
            state["current"] = name if kind == "own" else None
            state["other"] = name if kind == "other" else None
            skip_next = bool(PREFIX_ONLY.match(ln))
            continue
        if len(ln) < 6 or not re.search(r"[А-ЯЁІѢ][а-яёіѣъь]+", ln):
            continue
        cur = state.get("current")
        if cur is None:
            # заголовка своего НП ещё не было: берём только внутри диапазона
            # и только если не встретился чужой НП
            if state.get("other") or not strictly_inside or not state.get("started"):
                continue
            cur = "не определён (начало диапазона)"
        ages = [int(x) for x in re.findall(r"(?<!\w)(\d{1,3})(?!\w)", ln)]
        ages = [x for x in ages if 0 < x < 111]
        people.append({"settlement": cur, "line_no": i + 1,
                       "ages": ", ".join(str(x) for x in ages), "text": ln})
    return people


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


NEEDS_USER = re.compile(r"captcha|passport\.yandex|/auth", re.I)


def ensure_login(page):
    print("Проверяю вход в Яндекс...")
    try:
        page.goto(WORK_HOST + "/archive", wait_until="domcontentloaded", timeout=60_000)
    except Exception as e:  # noqa: BLE001
        print("Страница ya.ru/archive открылась с ошибкой:", e)
    on_archive = re.match(r"^https://(ya\.ru|yandex\.ru)/archive", page.url or "")
    if not on_archive or NEEDS_USER.search(page.url or ""):
        input("\n>>> Войдите в Яндекс в окне браузера (должен быть открыт ya.ru/archive) "
              "и нажмите Enter здесь...")
    # прогреваем сессию на втором домене (yandex.ru) — на случай ссылок с ним
    try:
        page.goto("https://yandex.ru/archive", wait_until="domcontentloaded", timeout=30_000)
        page.goto(WORK_HOST + "/archive", wait_until="domcontentloaded", timeout=30_000)
    except Exception:  # noqa: BLE001
        pass
    print("Вход проверен.")


_debug_saved = False


def fetch_sheet(page, file_id, n):
    """Открывает страницу n документа, возвращает dict: url, text, err."""
    global _debug_saved
    url = PAGE_URL_TEMPLATE.format(file_id=file_id, page=n)
    last_err = ""
    for attempt in range(MAX_RETRIES + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
            guard = 0
            while NEEDS_USER.search(page.url or "") and guard < 3:
                guard += 1
                if "captcha" in page.url:
                    input("\nЯндекс просит капчу. Решите её в браузере и нажмите Enter...")
                else:
                    input("\nСессия Яндекса сбросилась. Войдите заново в браузере и нажмите Enter...")
                page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
            if (page.url or "").startswith("chrome-error"):
                raise RuntimeError("страница не загрузилась (ошибка сети)")

            want_html = DEBUG_HTML and not _debug_saved
            res = page.evaluate(EXTRACT_JS, {"selectors": TEXT_SELECTORS, "maxWait": RENDER_WAIT_MS,
                                             "wantHtml": want_html}) or {}
            if want_html and res.get("html"):
                os.makedirs(DEBUG_DIR, exist_ok=True)
                with open(os.path.join(DEBUG_DIR, "first_page.html"), "w", encoding="utf-8") as f:
                    f.write(res["html"])
                _debug_saved = True
                print(f"Сохранён отладочный файл {DEBUG_DIR}/first_page.html")
            return {"url": url, "text": res.get("text") or "", "err": ""}
        except Exception as e:  # noqa: BLE001
            last_err = str(e) or type(e).__name__
            print(f"  страница {n}: ошибка «{last_err}», попытка {attempt + 1}/{MAX_RETRIES + 1}")
            time.sleep(2)
    return {"url": url, "text": "", "err": last_err or "ошибка"}


def _c(v):
    v = "" if v is None else v
    return v[:32000] if isinstance(v, str) else v


def _widths(sheet):
    for col in sheet.columns:
        w = 8
        for x in col[:200]:
            if x.value not in (None, ""):
                w = max(w, len(str(x.value)))
        sheet.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 70)


def save_workbook(path, people, sheets, links):
    wb = Workbook()
    bold = Font(bold=True)

    ws = wb.active
    ws.title = "Люди"
    ws.append(["Год", "Населённый пункт", "Страница", "Строка на странице",
               "Числа-возрасты (черновик)", "Распознанная строка", "Ссылка на страницу"])
    for r in people:
        ws.append([_c(x) for x in (r["year"], r["settlement"], r["page"], r["line_no"],
                                   r["ages"], r["text"], r["url"])])

    ws2 = wb.create_sheet("Страницы")
    ws2.append(["Год", "Страница", "Найденные НП округи", "Ссылка", "Распознанный текст страницы"])
    for r in sheets:
        ws2.append([_c(x) for x in (r["year"], r["page"], r["found"], r["url"], r["text"])])

    ws3 = wb.create_sheet("Ссылки")
    ws3.append(["Год", "Исходная ссылка", "Диапазон из ссылки", "Просмотрено страниц",
                "Найдено НП (шт.)", "Какие НП", "Решение", "Людей собрано", "Примечание"])
    for r in links:
        ws3.append([_c(x) for x in r])

    for sh in (ws, ws2, ws3):
        sh.freeze_panes = "A2"
        for cell in sh[1]:
            cell.font = bold
        _widths(sh)
    wb.save(path)


def main():
    html_path, xlsx_path = ask_paths()
    items = parse_links(html_path)
    if not items:
        sys.exit("В файле не найдено ссылок Яндекс Архива.")
    print(f"Ссылок Яндекс Архива: {len(items)}")

    people, sheet_rows, link_rows = [], [], []
    page_cache = {}          # (file_id, page) -> результат fetch_sheet
    seen_people = set()
    stop = False

    with sync_playwright() as pw:
        # 1) пробуем подключиться к уже запущенному Chrome (порт 9222), 2) иначе запускаем установленный Chrome
        try:
            browser = pw.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            print("Подключился к запущенному Chrome (порт 9222).")
        except Exception:  # noqa: BLE001
            print("Chrome с портом 9222 не найден — запускаю установленный Google Chrome...")
            browser = pw.chromium.launch(headless=False, channel="chrome")
            context = browser.new_context(locale="ru-RU")
        page = context.new_page()

        try:
            ensure_login(page)

            for idx, it in enumerate(items, 1):
                if stop:
                    break
                year = it["year"]
                tag = f"[{idx}/{len(items)}] {year if year is not None else '—'}"
                rng = f"{it['p_from']}–{it['p_to']}" if it["p_from"] is not None else "—"

                if not it["file_id"] or it["p_from"] is None:
                    link_rows.append([year, it["url"], rng, 0, 0, "", "пропущена: нет fileId или диапазона", 0, it["note"]])
                    print(f"{tag}: пропущена (нет fileId/диапазона)")
                    continue

                lo, hi = it["p_from"], it["p_to"]
                if hi - lo + 1 > BIG_RANGE:
                    print(f"{tag}: внимание, диапазон очень большой ({hi - lo + 1} страниц)")

                # 1. загрузка страниц диапазона из ссылки (диапазон — ровно из ссылки)
                texts = {}
                try:
                    for n in range(lo, hi + 1):
                        key = (it["file_id"], n)
                        if key not in page_cache:
                            print(f"{tag}: загружаю страницу {n} (диапазон {lo}–{hi})...")
                            page_cache[key] = fetch_sheet(page, it["file_id"], n)
                            if n < hi:
                                time.sleep(random.uniform(*DELAY_RANGE))
                        texts[n] = page_cache[key]
                except KeyboardInterrupt:
                    print("\nОстановка запрошена — сохраняю собранное...")
                    stop = True

                # 2. сбор людей
                found_all, per_page = set(), {}
                for n, r in texts.items():
                    f = find_settlements(r["text"])
                    per_page[n] = f
                    found_all |= f

                state = {"current": None, "other": None, "started": True}
                collected = 0
                for n in sorted(texts):
                    r = texts[n]
                    sheet_rows.append({"year": year, "page": n, "found": "; ".join(sorted(per_page[n])),
                                       "url": r["url"], "text": f"ОШИБКА: {r['err']}" if r["err"] else r["text"]})
                    for p in extract_people(r["text"], state, True):
                        key = (it["file_id"], n, p["line_no"], p["text"])
                        if key in seen_people:
                            continue
                        seen_people.add(key)
                        p.update(year=year, page=n, url=r["url"])
                        people.append(p)
                        collected += 1

                total = hi - lo + 1
                verdict = f"ПРЕРВАНА (страниц {len(texts)} из {total})" if len(texts) < total else "ОБРАБОТАНА"
                link_rows.append([year, it["url"], rng, len(texts), len(found_all), "; ".join(sorted(found_all)),
                                  verdict, collected, it["note"]])
                print(f"{tag}: диапазон {rng} — НП {len(found_all)}, людей: {collected}")
                save_workbook(xlsx_path, people, sheet_rows, link_rows)
        except KeyboardInterrupt:
            print("\nОстановлено пользователем.")
        finally:
            try:
                save_workbook(xlsx_path, people, sheet_rows, link_rows)
            finally:
                try:
                    browser.close()
                except Exception:  # noqa: BLE001
                    pass

    print(f"\nГотово. Людей: {len(people)}. Файл: {xlsx_path}")
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showinfo("Готово", f"Сохранено: {xlsx_path}\nЛюдей: {len(people)}")
    except Exception:  # noqa: BLE001
        pass


if __name__ == "__main__":
    main()
