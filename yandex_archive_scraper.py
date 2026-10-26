#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сбор данных о людях "моей округи" из исповедных ведомостей и ревизских сказок Яндекс Архива (версия 1.2).

РЕКОНСТРУКЦИЯ. Оригинальный файл версии 1.2 не сохранился. Скрипт восстановлен обратно
из расширения yandex-archive-extension 1.2.0 (logic.js + names.js + parser.js + page-scripts.js + app.js):
логика разбора перенесена из JavaScript в Python один в один. Оболочка (вход в Яндекс, запуск браузера,
окна выбора файлов, запись Excel) в расширении другая, поэтому взята из версии 1.3.

Что нового по сравнению с 1.1:
  * служебные строки (шапки колонок, «Итого…», «Семьи», «Лѣта», «Нынѣ на лицо», «Показанія дѣйства»...) отбрасываются;
  * каждый человек — отдельная строка: Имя, Отчество, Фамилия, Пол, Связь («жена», «сын», «брат»...), «С кем связан».
    Имена приводятся к современной орфографии и именительному падежу; обрезанные («дмитр», «косм») достраиваются;
    отчество детей достраивается от имени отца; пол определяется по имени (словарь — раздел «Имена»);
  * статусы (приёмыш, пасынок, бобыль, вдова, казак, увечен...) — колонка «Статус»;
  * события (умер, бежал, куплен, переведён, в рекруты, вновь рождённый) — колонка «Событие»;
  * год рождения = год документа (из ссылки, иначе из даты на странице, иначе из заголовка) − возраст;
  * «Ссылка на страницу» — только у главы семьи (хозяина двора);
  * заголовок-«титул» страницы («…уезда села Троицкого…») определяет населённый пункт;
    колонка «Дата в тексте страницы»; колонка «Примечание разбора» — всё, что стоит проверить глазами.

Известные ограничения (как и в расширении): возраст берётся только из числа в самой строке; словарь имён
(~170 имён) можно дополнять в списках MALE / FEMALE / ALIASES.

Диапазон страниц берётся РОВНО из ссылки. Остановка: Ctrl+C — собранное сохраняется.

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

HEADER_MAX_LEN = 70    # строка короче этого может быть заголовком населённого пункта
BIG_RANGE = 300        # предупреждение о слишком большом диапазоне
MAX_MISSES = 3         # подряд страниц без окна текста — остановка

DELAY_RANGE = (1.0, 2.5)
PAGE_TIMEOUT_MS = 45_000
MAX_RETRIES = 2
RENDER_WAIT_MS = 15_000
DEBUG_HTML = True                      # сохранить HTML первой страницы (first_page.html)
DEBUG_DIR = "debug_yandex_archive"

# Ручная настройка окон (аналог кнопок «Указать окно текста / строку даты вручную» в расширении).
# CSS-селектор; пустая строка — автоопределение.
TEXT_SELECTOR = ""
DATE_SELECTOR = ""

# ------------------------------------------------------------------
# РЕЖИМЫ ОБРАБОТКИ. Здесь меняются параметры сбора под тип документа.
# ------------------------------------------------------------------
MODES = {
    "confession": {
        "id": "confession",
        "label": "Исповедные ведомости",
        "extra_prefixes": [],       # слова-префиксы заголовка НП (кроме "сел…" и "дер…")
        "min_line_len": 6,          # минимальная длина строки с человеком
        "two_ages": False,          # в строке один возраст
        "household": False,         # номера дворов не отслеживаем
        "default_span": 15,         # если в ссылке нет sheet_page_to: from .. from+default_span
        "description": ("Заголовки: село / деревня. Собираются: двор (по порядку), № мужчины / женщины, возраст "
                        "(один), вычисляется год рождения = год документа − возраст. Колонки «Показания действия» "
                        "(у исповеди были / исповедались, но не причастились / не были) не собираются. "
                        "Ссылка на страницу — только у хозяина двора."),
    },
    "revision": {
        "id": "revision",
        "label": "Ревизские сказки",
        "extra_prefixes": ["починок", "починк", "выселок", "выселк", "слобод", "хутор", "пустош", "погост"],
        "min_line_len": 5,          # записи короче: «сын Иван 5»
        "two_ages": True,           # обычно два возраста: по прошлой ревизии и ныне
        "household": True,          # отслеживаем «Двор № 5» / «№ 12»
        "default_span": 15,
        "description": ("Заголовки: также починок, выселок, слобода, хутор и т. п. Каждый человек — в своей строке: "
                        "имя, отчество (достраивается от отца), фамилия, пол (по имени), связь (жена, сын, брат...), "
                        "статус (приёмыш, бобыль, вдова...) сохраняется; события (умер, бежал, куплен, в рекруты) — "
                        "в колонку «Событие». Ссылка — только главе семьи. Старая орфография заменяется на "
                        "современную, служебные строки шапки отбрасываются."),
    },
}
DEFAULT_MODE = "confession"
# ----------------------------------------------------------------------------


def get_mode(mode_id):
    return MODES.get(mode_id) or MODES[DEFAULT_MODE]


# ============================================================================
# ИМЕНА (names.js): современная орфография, падежи, пол, отчества «в стиле ревизских сказок»
# ============================================================================
MALE = ("Адриан Акакий Александр Алексей Аким Анисим Андрей Андреян Антон Аникей Арефий Артемий Архип Афанасий "
        "Богдан Борис Василий Варфоломей Гаврила Герасим Гордей Григорий Гурий Данила Дементий Демид Денис Дмитрий "
        "Евграф Евдоким Евсей Евстафий Егор Елисей Емельян Епифан Ермолай Ефим Ефрем Захар Зиновий Иван Игнатий Игнат "
        "Илларион Илья Осип Исай Кирилл Клим Кондратий Конон Константин Корней Кузьма Лаврентий Лазарь Лев Леонтий Лука "
        "Лукьян Макар Максим Мартын Матвей Мина Михаил Мокей Моисей Назар Нестор Никанор Никита Никифор Николай Никон "
        "Нефед Астафий Павел Панкрат Пантелей Парфен Петр Пимен Платон Поликарп Прокофий Прохор Роман Савва Савелий "
        "Самсон Семен Сергей Сидор Сила Силантий Спиридон Степан Тарас Терентий Тимофей Тит Трофим Тихон Фаддей Федор "
        "Федот Федосей Филат Филипп Фока Фома Харитон Христофор Яков Яким Анофрий Олимпий Полуехт Лестарх Ларион Мирон "
        "Мосей Нифонт Нил Онисим Остафий Пров Прокл Софрон Тарасий Устин Фирс Ферапонт Фаддей Юрий").split(" ")
FEMALE = ("Авдотья Агафья Аграфена Агриппина Акулина Александра Алена Анастасия Анна Анисья Анфиса Арина "
          "Афимья Варвара Василиса Вера Дарья Домна Домника Евдокия Евпраксия Екатерина Елена Елизавета Ефросинья Ирина "
          "Ксения Лукерья Любовь Малания Марина Марфа Мавра Мария Матрена Меланья Наталья Настасья Неонила Ольга Пелагея "
          "Прасковья Параскева Софья Степанида Стефанида Татьяна Ульяна Устинья Улита Фекла Федосья Федора Фетинья "
          "Христина Таисия Харитина Аксинья Акилина Марья Катерина Олена Соломонида Аннушка Дуня Ненила Лидия Надежда"
          ).split(" ")

# церковные / сокращённые / иначе написанные формы → крестьянская форма
ALIASES = {
    "иоанн": "Иван", "ианн": "Иван", "иаков": "Яков", "георгий": "Егор", "юрий": "Егор", "феодор": "Федор",
    "иосиф": "Осип", "йосиф": "Осип", "алексий": "Алексей", "евфимий": "Ефим", "сергий": "Сергей",
    "симеон": "Семен", "симон": "Семен", "фома": "Фома", "феодосий": "Федосей", "гавриил": "Гаврила",
    "даниил": "Данила", "данил": "Данила", "кирилл": "Кирила", "кондрат": "Кондратий", "кандратий": "Кондратий",
    "кандратей": "Кондратий", "кондратей": "Кондратий", "ефрем": "Ефрем", "евфросиния": "Ефросинья",
    "мария": "Марья", "наталия": "Наталья", "феврония": "Феврония", "екатерина": "Екатерина",
    "катерина": "Екатерина", "татиана": "Татьяна", "ирина": "Арина", "иулиания": "Ульяна", "юлиания": "Ульяна",
    "ульяна": "Ульяна", "улиана": "Ульяна", "пелагия": "Пелагея", "феодора": "Федора", "феодосия": "Федосья",
    "евдокия": "Авдотья", "авдотия": "Авдотья", "анисия": "Анисья", "акулина": "Акулина", "аксиния": "Аксинья",
    "ксения": "Аксинья", "афанасей": "Афанасий", "астафий": "Остафий", "василей": "Василий", "григорей": "Григорий",
    "дмитрей": "Дмитрий", "игнатей": "Игнатий", "леонтей": "Леонтий", "прокофей": "Прокофий", "терентей": "Терентий",
    "андреи": "Андрей", "алексеи": "Алексей", "матвеи": "Матвей", "тимофеи": "Тимофей", "сергеи": "Сергей",
    "евсеи": "Евсей", "елисеи": "Елисей", "корнеи": "Корней", "василеи": "Василий", "григореи": "Григорий",
    "дмитреи": "Дмитрий", "николаи": "Николай", "федосеи": "Федосей", "ермолаи": "Ермолай", "пантелеи": "Пантелей",
    "мокеи": "Мокей", "аникеи": "Аникей", "арефеи": "Арефий", "афонасеи": "Афанасий", "афанасеи": "Афанасий",
    "казма": "Кузьма", "косма": "Кузьма", "косьма": "Кузьма", "козьма": "Кузьма", "иванъ": "Иван",
    "стефан": "Степан", "стефанъ": "Степан", "ярофей": "Тимофей", "афимья": "Афимья", "нестер": "Нестор",
    "нестор": "Нестор", "афонасий": "Афанасий", "нифантий": "Нифонт", "фетинья": "Фетинья", "оксинья": "Аксинья",
    "парасковья": "Прасковья", "прасковия": "Прасковья",
}

# особые формы отчества (основа отчества от имени отца, мужская)
PATR_SPECIAL = {
    "Павел": "Павлов", "Яков": "Яковлев", "Михаил": "Михайлов", "Лев": "Львов", "Петр": "Петров",
    "Гаврила": "Гаврилов", "Нефед": "Нефедов", "Любим": "Любимов", "Данила": "Данилов", "Сила": "Силин",
    "Илья": "Ильин", "Никита": "Никитин", "Кузьма": "Кузьмин", "Фома": "Фомин", "Лука": "Лукин", "Мина": "Минин",
    "Савва": "Саввин", "Осип": "Осипов", "Афанасий": "Афанасьев", "Василий": "Васильев", "Григорий": "Григорьев",
    "Дмитрий": "Дмитриев",
}

# ---------------------------------------------------------------- орфография
_MARKS = re.compile("[̀-ͯ҃-҉]")
_MODERN_MAP = {"ѣ": "е", "Ѣ": "Е", "і": "и", "І": "И", "ї": "и", "Ѳ": "Ф", "ѳ": "ф", "Ѵ": "И", "ѵ": "и",
               "i": "и", "ѡ": "о", "Ѡ": "О", "ѧ": "я", "ѫ": "у"}
_MODERN_RX = re.compile("[ѣѢіІїѲѳѴѵiѡѠѧѫ]")
_HARD_END = re.compile("([А-Яа-яЁё])[Ъъ](?![А-Яа-яЁё])")


def modern(s):
    s = _MARKS.sub("", s or "")
    s = _MODERN_RX.sub(lambda m: _MODERN_MAP.get(m.group(0), m.group(0)), s)
    return _HARD_END.sub(r"\1", s)       # твёрдый знак в конце слова


def key(s):
    """Ключ для поиска: современная орфография, строчные, е вместо ё, без ь, й → и."""
    s = modern(s).lower().replace("ё", "е")
    s = re.sub("[^а-я]", "", s)
    return s.replace("ь", "").replace("й", "и")


def cap(s):
    return s[:1].upper() + s[1:].lower() if s else s


# ---------------------------------------------------------------- склонение
def forms(nom, gender):
    out = {"nom": [nom], "gen": [], "other": []}
    base, last, l2 = nom[:-1], nom[-1:], nom[-2:]
    soft = bool(re.search("[гкхжчшщ]$", base))
    if gender == "м":
        if l2 == "ий":
            out["gen"] += [nom[:-2] + "ия", nom[:-2] + "ья"]
            out["other"] += [nom[:-2] + "ию", nom[:-2] + "ием"]
        elif last == "й":
            out["gen"].append(base + "я")
            out["other"] += [base + "ю", base + "ем"]
        elif last == "ь":
            out["gen"].append(base + "я")
            out["other"] += [base + "ю", base + "ем"]
        elif last == "а":
            out["gen"].append(base + ("и" if soft else "ы"))
            out["other"] += [base + "е", base + "у", base + "ой"]
        elif last == "я":
            out["gen"].append(base + "и")
            out["other"] += [base + "е", base + "ю"]
        else:
            out["gen"].append(nom + "а")
            out["other"] += [nom + "у", nom + "ом"]
            if re.search("[рлнм]$", nom) and len(nom) > 4:
                out["gen"].append(re.sub("е([^е]*)$", r"\1", nom, count=1) + "а")
        if nom == "Павел":
            out["gen"].append("Павла")
            out["other"] += ["Павлу", "Павлом"]
        if nom == "Петр":
            out["gen"].append("Петра")
            out["other"].append("Петру")
        if nom == "Лев":
            out["gen"].append("Льва")
    else:
        if l2 == "ия":
            out["gen"].append(base + "и")
            out["other"] += [base + "ю", nom[:-1] + "и"]
        elif l2 == "ья":
            out["gen"].append(nom[:-1] + "и")
            out["other"] += [nom[:-1] + "ю", nom[:-1] + "ей"]
        elif last == "я":
            out["gen"].append(base + "и")
            out["other"] += [base + "ю", base + "ей"]
        elif last == "а":
            out["gen"].append(base + ("и" if soft else "ы"))
            out["other"] += [base + "е", base + "у", base + "ой"]
        elif last == "ь":
            out["gen"].append(base + "и")
            out["other"] += [base + "и", base + "ью"]
        else:
            out["gen"].append(nom + "и")
    return out


# индекс: ключ -> [{nom, g, cs}]
INDEX = {}
NOMKEYS = []
FZ = {}


def _fz(k):
    return k.replace("о", "а").replace("е", "и")


def _add(k, rec):
    if not k:
        return
    arr = INDEX.setdefault(k, [])
    if not any(r["nom"] == rec["nom"] and r["cs"] == rec["cs"] for r in arr):
        arr.append(rec)
    fa = FZ.setdefault(_fz(k), [])
    if not any(r["nom"] == rec["nom"] and r["cs"] == rec["cs"] for r in fa):
        fa.append(rec)


def _build_index():
    allnames = [(n, "м") for n in MALE] + [(n, "ж") for n in FEMALE]
    for n, g in allnames:
        _add(key(n), {"nom": n, "g": g, "cs": "nom"})
        NOMKEYS.append((key(n), n, g))
    for n, g in allnames:
        f = forms(n, g)
        for x in f["gen"]:
            _add(key(x), {"nom": n, "g": g, "cs": "gen"})
        for x in f["other"]:
            _add(key(x), {"nom": n, "g": g, "cs": "obl"})
    for a, to in ALIASES.items():
        t0 = INDEX.get(key(to))
        if not t0:
            continue
        t = t0[0]
        _add(key(a), {"nom": t["nom"], "g": t["g"], "cs": "nom"})
        NOMKEYS.append((key(a), t["nom"], t["g"]))
        f = forms(a[:1].upper() + a[1:], t["g"])
        for x in f["gen"]:
            _add(key(x), {"nom": t["nom"], "g": t["g"], "cs": "gen"})


_build_index()


def _pick(arr, prefer):
    if prefer:
        x = next((r for r in arr if r["cs"] in ("gen", "obl")), None)
        if x:
            return x
    return next((r for r in arr if r["cs"] == "nom"), arr[0])


def lookup(word, prefer=False):
    """Поиск имени: {nom, g, cs, trunc} или None. Обрезанные «дмитр», «косм» — по единственному совпадению."""
    k = key(word)
    if len(k) < 3:
        return None
    hit = INDEX.get(k)
    if hit:
        return {**_pick(hit, prefer), "trunc": False, "alt": len(hit) > 1}
    fh = FZ.get(_fz(k))
    if fh:
        return {**_pick(fh, prefer), "trunc": False, "fuzzy": True}
    if len(k) >= 4:
        c = [x for x in NOMKEYS if x[0].startswith(k)]
        uniq = {x[1] for x in c}
        if len(uniq) == 1:
            _, nom, g = c[0]
            return {"nom": nom, "g": g, "cs": "nom", "trunc": True}
    return None


# ---------------------------------------------------------------- отчества
def patr_base(nom):                      # Иван → Иванов
    if nom in PATR_SPECIAL:
        return PATR_SPECIAL[nom]
    if re.search("ий$", nom):
        return nom[:-2] + "иев"
    if re.search("ей$", nom):
        return nom[:-2] + "еев"
    if re.search("ай$", nom):
        return nom[:-2] + "аев"
    if re.search("й$", nom):
        return nom[:-1] + "ев"
    if re.search("ь$", nom):
        return nom[:-1] + "ев"
    if re.search("[ая]$", nom):
        return nom[:-1] + "ин"
    if re.search("[жчшщц]$", nom):
        return nom + "ев"
    return nom + "ов"


def patr(father_nom, gender):
    return patr_base(father_nom) + ("а" if gender == "ж" else "") if father_nom else ""


PATR_SET = {key(patr_base(n)) for n in MALE}
PATR_RX = re.compile(r"^(.{2,}?(?:ов|ев|ин|ын))(а|у|ы|ым|ом|ых|ой|ою|ыми|ами|е)?$")


def patr_parse(word):
    """основа: «Иванова»/«Ивановы»/«Иванову» → «Иванов»; возвращает dict или None."""
    w = modern(word)
    k = w.lower().replace("ё", "е")
    m = PATR_RX.match(k)
    if not m:
        return None
    suf = m.group(2)
    return {"base": cap(m.group(1)), "plural": suf in ("ы", "ых"), "fem": suf in ("а", "ой", "ою"),
            "known": key(m.group(1)) in PATR_SET}


def patr_to_father(base):
    """отчество → имя отца (для привязки «Ивановы дети» к Ивану)."""
    k = key(base)
    for n in MALE:
        if key(patr_base(n)) == k:
            return n
    return None


# ---------------------------------------------------------------- слова-связи, статусы, события
REL = {}


def _rel_add(rel, g, plural, words):
    for w in words.split(" "):
        REL[key(w)] = {"rel": rel, "g": g, "plural": plural}


_rel_add("жена", "ж", False, "жена жѣна")
_rel_add("вдова", "ж", False, "вдова")
_rel_add("дочь", "ж", False, "дочь дочъ дочерь дочери дочка")
_rel_add("дочери", "ж", True, "дочерей дочеры")
_rel_add("сын", "м", False, "сын снъ сынъ сынь")
_rel_add("сыновья", "м", True, "сыновья сыновья сыновей сыновья")
_rel_add("дети", "", True, "дети дѣти детей дѣтей дети детки")
_rel_add("брат", "м", False, "брат братъ")
_rel_add("братья", "м", True, "братья братьи")
_rel_add("сестра", "ж", False, "сестра")
_rel_add("племянник", "м", False, "племянник племянникъ")
_rel_add("племянники", "м", True, "племянники племянникы племянницы")
_rel_add("зять", "м", False, "зять")
_rel_add("мать", "ж", False, "мать матъ")
_rel_add("невестка", "ж", False, "невестка снохa сноха")
_rel_add("внук", "м", False, "внук внукъ")
_rel_add("внуки", "м", True, "внуки внуки внуков")
_rel_add("отец", "м", False, "отец отецъ")

# статусы сохраняем (в «Статус» и в строке)
STATUS_WORDS = ["приемыш", "примышь", "приимыш", "пасынок", "пономарь", "салдатка", "солдатка", "казак", "бобыль",
                "вдов", "холост", "женат", "девица", "женка", "глуха", "глух", "увечен", "без ног", "крив", "слеп",
                "сирота"]
STATUS_KEYS = [key(w) for w in STATUS_WORDS if " " not in w]
NUMWORD_KEYS = {"один", "одна", "одно", "два", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять",
                "десять", "одиннадцать", "двенадцать", "тринадцать", "двадцать", "тридцать", "сорок", "пятьдесят",
                "шестьдесят", "семьдесят", "восемьдесят", "девяносто", "сто", "полтора", "полугода"}

# события (выбыл / отлучка / прибыл)
EVENT_RX = re.compile(
    r"(?:^|\s)(умер|умре|умерла|умерли|бежал|бежала|бежали|купл[её]н[а-я]*|куплен[а-я]*|причисл[а-я]*|"
    r"переведен[а-я]*|переведе[а-я]*|отдан[а-я]*|продан[а-я]*|отпущен[а-я]*|убыл[а-я]*|выбыл[а-я]*|"
    r"вновь\s+рожд[а-я]*|новорожд[а-я]*|взят[а-я]*|женил[а-я]*|вышла\s+замуж|пропущен[а-я]*|"
    r"в\s+рекрут[а-я]*|в\s+отлучк[а-я]*|в\s+солдат[а-я]*)", re.I)
EVENT_SPLIT_RX = re.compile(
    r"(?=(?:^|\s)(?:умер|умре|бежал|купл[её]н|причисл|переведен|отдан|продан|отпущен|убыл|выбыл|"
    r"вновь\s+рожд|новорожд|взят|женил|пропущен)[а-я]*)", re.I)


# ============================================================================
# РАЗБОР СТРОКИ С ЛЮДЬМИ (parser.js): артефакты, имена/отчества, связи, статусы, события
# ============================================================================
def kline(s):
    """нормализованная строка с пробелами (современная орфография, без ь, только русские буквы)."""
    s = modern(s).lower().replace("ё", "е").replace("ь", "")
    return re.sub(r"\s+", " ", re.sub("[^а-я ]", " ", s)).strip()


# служебные строки шапки/итогов и колонок, которые не собираются
ARTIFACT = re.compile("^(?:" + "|".join([
    "лета", "семьи", "семья", "ныне", "нине", "на лицо", "налицо", "крестьян", "мужеск", "женск", "итого",
    "по последне", "по после", "ис того", "из того", "во временно", "с которого", "скоторого", "скоторо", "когда",
    "кто были", "которые", "кто же", "ктож", "показани", "числа", "состояли", "состояло", "выбыли", "подано",
    "ревизск", "ревизия", "сказка", "оборот", "попов$", "исповедь", "у исповеди", "и святого", "исповедались",
    "токмо", "причастились", "домов", "двор$", "число$", "людеи", "мужеска", "женска", "всего", "прибыло", "убыло",
    "выбыло", "в том числе", "с прибылыми", "по ревизии", "во время", "родилось", "умерло", "имена", "и отчества",
]) + ")")

LINK_SKIP = {"же", "ж", "того", "тож", "тоже", "также", "такожде", "другои", "второи", "третии", "меншои", "малои",
             "большои", "балшои", "старшии", "меньшии", "из", "от", "на", "во", "по", "при", "при ним", "году",
             "года", "лет", "лета", "ныне", "нине", "отъ", "его", "ея", "ее", "их", "ими", "но", "как", "был",
             "была", "были", "которои", "котораи", "первои", "вновь", "они", "он", "она", "им", "ему", "ей",
             "оного", "сих", "сеи", "вдовы"}
SMALL = {"и", "в", "с", "к", "о", "у", "а", "я", "во", "по", "на", "за", "из", "от", "до", "же", "ли", "не", "ни",
         "ж", "им", "ей"}


def _is_capital(w):
    return bool(re.match("^[А-ЯЁІѢ]", w))


def tokenize(line):
    out = []
    raw = [x for x in re.split(r"\s+", re.sub('[“”«»"()\\[\\]{}]', " ", modern(line))) if x]
    for r0 in raw:
        colon = bool(re.search("[:;]$", r0))
        r = re.sub(r"^[\W_]+|[\W_]+$", "", r0)
        if not r:
            continue
        if re.fullmatch("[0-9]+", r):
            out.append({"type": "num", "v": int(r), "w": r})
            continue
        if re.search("[0-9]", r):                      # «2-й», «1810-го», «9½»
            out.append({"type": "skip", "w": r})
            continue
        w = "".join(ch for ch in r if ch.isalpha() or ch == "-").strip("-")
        if not w:
            continue
        out.append({"type": "word", "w": w, "colon": colon})
    return out


def classify(t):
    if t["type"] != "word":
        return t
    w = t["w"]
    k = key(w)
    t["k"] = k
    if not k:
        t["type"] = "skip"
        return t
    if k in NUMWORD_KEYS:
        t["type"] = "skip"
        return t
    if k == "его":
        t["type"], t["who"] = "link", "м"
        return t
    if k in ("ея", "ее", "еи"):
        t["type"], t["who"] = "link", "ж"
        return t
    if k in REL:
        t["type"], t["rel"] = "rel", REL[k]
        return t
    if len(k) >= 3 and any(k.startswith(s) for s in STATUS_KEYS):
        t["type"] = "status"
        return t
    if k in LINK_SKIP or k in SMALL or len(w) <= 1:
        t["type"] = "skip"
        return t
    nm = lookup(w)
    if nm and len(w) >= 3:
        t["type"], t["nm"], t["nmGen"] = "name", nm, lookup(w, True)
        return t
    p = patr_parse(w)
    if p:
        t["type"], t["p"] = "pat", p
        return t
    t["type"] = "unk" if (_is_capital(w) or len(w) >= 4) else "skip"
    return t


def _pc(w):
    return "-".join(cap(x) for x in (w or "").split("-"))


def modern_word(w):
    return _pc("".join(ch for ch in modern(w) if ch.isalpha() or ch == "-"))


def new_person(ctx, **o):
    p = {"name": "", "patr": "", "surname": "", "gender": "", "rel": "", "relTo": "", "status": "", "event": "",
         "head": False, "raw": [], "flags": [], "famNo": ctx["fam"]["n"], "anchor": None}
    p.update(o)
    ctx["fam"]["members"].append(p)
    return p


def find_anchor(ctx, name_tok, pat_tok):
    fam = ctx["fam"]["members"]
    want = name_tok["nm"]["nom"]
    pk = key(pat_tok["p"]["base"]) if pat_tok else None
    for m in reversed(fam):
        if m["name"] == want and (not pk or key(re.sub("а$", "", m["patr"])) == pk):
            return m
    for m in reversed(fam):
        if m["name"] == want:
            return m
    return None


def parse_line(line, ctx):
    toks = [classify(t) for t in tokenize(line)]
    res = {"persons": [], "events": [], "ages": [], "nums": []}

    # события: от слова-события до конца строки (или до следующего имени)
    ml = modern(line)
    em = EVENT_RX.search(ml)
    names_before = toks
    if em:
        idx = em.start()
        head = ml[:idx].strip()
        tail = ml[idx:].strip()
        evs = []
        for s in EVENT_SPLIT_RX.split(tail):
            s = re.sub(r"^[,.;]+|[,.;]+$", "", re.sub(r"\s+", " ", modern(s)).strip())
            if s:
                evs.append(s[:1].upper() + s[1:].lower())
        res["events"] = evs
        names_before = [classify(t) for t in tokenize(head)]
    T = names_before
    for t in T:                                       # числа строки
        if t["type"] == "num" and 0 < t["v"] < 111:
            res["ages"].append(t["v"])
    W = [t for t in T if t["type"] not in ("num", "skip")]
    res["nums"] = res["ages"]

    listing = None
    cur = None                # последний созданный в строке
    prev_type = ""
    pending_anchor = None

    def build(name_tok, rel):
        nm = name_tok["nm"]
        p = new_person(ctx, name=nm["nom"], gender=nm["g"], rel=rel["rel"] if rel else "", relTo="")
        p["raw"].append(modern_word(name_tok["w"]))
        if nm.get("fuzzy"):
            p["flags"].append("имя по созвучию: «" + name_tok["w"] + "» → " + nm["nom"])
        if nm.get("trunc"):
            p["flags"].append("имя обрезано: «" + name_tok["w"] + "» → " + nm["nom"])
        if rel:
            if rel.get("g") and nm["g"] and rel["g"] != nm["g"]:
                p["flags"].append("пол по имени, а не по связи «" + rel["rel"] + "»")
            if not p["gender"]:
                p["gender"] = rel.get("g")
        return p

    i = 0
    while i < len(W):
        t = W[i]
        nxt = W[i + 1] if i + 1 < len(W) else None
        if t["type"] == "name":
            # якорь: имя в косвенном падеже + (отчество/фамилия) + связь
            j, pat_tok = i + 1, None
            while j < len(W) and W[j]["type"] in ("pat", "unk") and j - i <= 3:
                if W[j]["type"] == "pat" and not pat_tok:
                    pat_tok = W[j]
                j += 1
            if (j < len(W) and W[j]["type"] == "rel" and t["nmGen"] and t["nmGen"]["cs"] in ("gen", "obl")
                    and (t["nm"]["cs"] != "nom" or j > i + 1 or i == 0 or not listing)):
                t["nm"] = t["nmGen"]
                pending_anchor = find_anchor(ctx, t, pat_tok) or new_person(
                    ctx, name=t["nm"]["nom"], gender=t["nm"]["g"],
                    patr=((patr(None, t["nm"]["g"]) or pat_tok["p"]["base"]) if pat_tok else ""), virtual=True)
                if pending_anchor.get("virtual") and not pending_anchor["raw"]:
                    pending_anchor["raw"].append(modern_word(t["w"]))
                    pending_anchor["flags"].append("человек из ссылки, отдельной строки нет")
                ctx["fam"]["anchor"] = pending_anchor
                listing = None
                i = j                      # (JS: i = j - 1; continue → i++)
                continue
            if listing:
                p = build(t, listing)
                a = listing.get("anchor")
                p["relTo"] = " ".join(x for x in (a["name"], a["patr"], a["surname"]) if x) if a else ""
                p["anchor"] = a
            elif cur and prev_type == "name" and not cur["patr"] and not cur["surname"] and cur["name"] != t["nm"]["nom"]:
                # «многоименный»: следующее имя подряд — отдельный ребёнок первого
                p = build(t, {"rel": "дочь" if t["nm"]["g"] == "ж" else "сын", "g": t["nm"]["g"]})
                p["anchor"] = cur
                p["relTo"] = cur["name"]
                p["flags"].append("многоименная запись: имя отнесено к ребёнку")
            else:
                p = build(t, None)
                p["head"] = True
                ctx["fam"]["n"] += 1
                ctx["fam"]["members"] = [p]
                p["famNo"] = ctx["fam"]["n"]
                ctx["fam"]["anchor"] = p
                pending_anchor = None
            res["persons"].append(p)
            cur = p
        elif t["type"] == "pat" and nxt and nxt["type"] == "rel":
            # «Ивановы дети», «Матвеевы сыновья»: якорь — человек с таким отчеством отца
            father = patr_to_father(t["p"]["base"])
            a = None
            if father:
                for m in reversed(ctx["fam"]["members"]):
                    if m["name"] == father:
                        a = m
                        break
            pending_anchor = a or cur or ctx["fam"]["anchor"]
            ctx["fam"]["anchor"] = pending_anchor
        elif t["type"] == "rel":
            prev_link = W[i - 1] if i > 0 and W[i - 1]["type"] == "link" else None
            a = pending_anchor or ctx["fam"]["anchor"]
            if prev_link and prev_link["who"] == "ж":
                f = next((m for m in reversed(ctx["fam"]["members"]) if m["gender"] == "ж"), None)
                if f:
                    a = f
            listing = {"rel": t["rel"]["rel"], "g": t["rel"]["g"], "plural": t["rel"]["plural"], "anchor": a}
            pending_anchor = None
        elif t["type"] == "link":
            pass                           # «Его»/«Ея» ссылаются на текущий якорь; ничего не создаём
        elif t["type"] == "status":
            target = cur or ctx["fam"]["anchor"]
            if target:
                target["status"] = target["status"] + "; " + t["k"] if target["status"] else t["k"]
                target["raw"].append(modern_word(t["w"]))
        elif t["type"] == "pat":
            if cur and not cur["patr"]:
                cur["patr"] = t["p"]["base"] + ("а" if cur["gender"] == "ж" else "")
                cur["raw"].append(modern_word(t["w"]))
            elif cur and not cur["surname"]:
                cur["surname"] = modern_word(t["w"])
                cur["raw"].append(modern_word(t["w"]))
        elif t["type"] == "unk":
            if cur and not cur["surname"] and cur["raw"]:
                cur["surname"] = modern_word(t["w"])
                cur["raw"].append(modern_word(t["w"]))
        prev_type = t["type"]
        i += 1

    # отчества и фамилии от якоря
    for p in res["persons"]:
        a = p["anchor"]
        if not a or p["head"]:
            continue
        if not p["patr"]:
            if p["rel"] in ("сын", "дочь", "дети", "сыновья", "дочери") and a["name"]:
                p["patr"] = patr(a["name"], p["gender"])
            elif p["rel"] in ("брат", "братья", "сестра") and a["patr"]:
                p["patr"] = re.sub("а$", "", a["patr"]) + ("а" if p["gender"] == "ж" else "")
            if p["patr"]:
                p["flags"].append("отчество достроено от " +
                                  ("брата/сестры" if p["rel"] in ("брат", "братья", "сестра") else "отца"))
        if not p["surname"] and a["surname"] and p["rel"] != "жена":
            p["surname"] = a["surname"]
    for p in res["persons"]:
        p["full"] = " ".join(x for x in (p["name"], p["patr"], p["surname"]) if x)
    return res


def is_artifact(line):
    k = kline(line)
    if not k:
        return True
    return bool(ARTIFACT.match(k))


_MONTH_RX = re.compile(r"январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр|"
                       r"іюн|iюн|iюл|мар|фев|янв|ноя|ок", re.I)


def text_date(text):
    """дата из заголовка страницы: «1850 года Іюля 10 дня», «1811 Года сентября 30 дня»."""
    lines = re.split(r"\r?\n", str(text or ""))
    for ln in lines:
        s = modern(ln).strip()
        if len(s) < 60 and re.search(r"(?<![0-9])1[5-9][0-9][0-9](?:-го)?\s+года?\s+[а-я]+\s+[0-9]{1,2}\s+дня", s, re.I):
            return s
    for ln in lines:
        m = re.search(r"(?<![0-9])(1[5-9][0-9][0-9])(?:-го)?\s+года?\s+([а-яА-Я]+)\s+(?:дня|числа)?", modern(ln), re.I)
        if m and _MONTH_RX.search(m.group(2)):
            return ln.strip()[:80]
    return ""


# ============================================================================
# РАЗБОР ЛИНКОВ / ЗАГОЛОВКОВ / ЛЮДЕЙ (logic.js)
# ============================================================================
# строки-заголовки документа, которые не являются людьми
TITLE_NOISE = re.compile(r"(?:ревизск|исповедн|ведомост|сказк|уезд|губерн|волост|приход)", re.I)
TITLE_NOISE_EXC = re.compile(r"[а-яё]+\s+[а-яё]+\s+[а-яё]+\s+(?:сын|жена|дочь)", re.I)

# Границы слов (в Python \w понимает кириллицу)
WB = r"(?<!\w)"
WE = r"(?!\w)"


def ci(w: str) -> str:
    """Слово -> регулярка, нечувствительная к регистру (без флага re.I, чтобы [А-Я] оставалось строгим)."""
    return "".join(ch if ch.upper() == ch else f"[{ch.upper()}{ch}]" for ch in w)


# Населённые пункты округи: название -> основа (по нормализованному тексту)
SETTLEMENT_TAILS = {
    "село Троицкое":     r"троицк",
    "деревня Острогина": r"острогин",
    "деревня Черепкова": r"черепков",
    "деревня Рублева":   r"рублев",
    "деревня Луки":      r"лук(?:и)?" + WE,
    "деревня Мякинина":  r"(?:м[яе]кинин|мемин)",
}

_compiled = {}


def compile_mode(mode_id):
    m = get_mode(mode_id)
    if m["id"] in _compiled:
        return _compiled[m["id"]]
    words = "|".join(["сел", "дер"] + m["extra_prefixes"])
    words_ci = "|".join(ci(w) for w in ["сел", "дер"] + m["extra_prefixes"])

    # Обязателен префикс (в т.ч. "дер." / "д." / "с."). Применяется к нормализованному тексту.
    prefix = rf"(?:{WB}(?:{words})\w*\.?|{WB}[дс]\.)\s*[\[\(«\"]?\s*"
    settlements = {name: re.compile(prefix + tail) for name, tail in SETTLEMENT_TAILS.items()}

    # "Деревня"/"Дер." с любой буквы; сокращения "д."/"с." только строчные,
    # чтобы не принять за заголовок инициалы ("Д. Иванов").
    other_header = re.compile(
        rf"(?:{WB}(?:{words_ci})\w*\.?|{WB}[дс]\.)\s*[\[\(«\"]?\s*([А-ЯЁІѢ][а-яёіѣъь]{{3,}})")
    # строка, состоящая только из слова-префикса ("Деревня" / "Село" / "д.")
    prefix_only = re.compile(rf"^\s*(?:(?:{words_ci})\w*|[дсДС])\.?\s*$")

    # маркеры двора: "Двор № 5 ..." или строка "№ 12"
    dvor = ci("двор")
    hh_a = re.compile(rf"^\s*(?:{dvor}\w*|[Дд][Вв])\.?\s*(?:№|N|n)?\s*(\d{{1,3}})(?!\w)")
    hh_b = re.compile(r"^\s*(?:№|N)\s*(\d{1,3})\s*[.:]?\s*$")

    _compiled[m["id"]] = dict(settlements=settlements, other_header=other_header,
                              prefix_only=prefix_only, hh_a=hh_a, hh_b=hh_b)
    return _compiled[m["id"]]


def norm(s: str) -> str:
    s = s.lower()
    for a, b in (("ё", "е"), ("ѣ", "е"), ("і", "и"), ("ѵ", "и"), ("ѳ", "ф")):
        s = s.replace(a, b)
    return s.replace("ъ", "").replace("ь", "")


def _int(s):
    try:
        return int(s)
    except (TypeError, ValueError):
        return None


def parse_links(html_path, mode_id=DEFAULT_MODE):
    """Возвращает список словарей: url, year, file_id, p_from, p_to, note."""
    span = get_mode(mode_id)["default_span"]
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
            p_to = p_from + span
            note = f"нет sheet_page_to — взято from+{span}"
        if p_from is not None and p_to is not None and p_from > p_to:
            p_from, p_to = p_to, p_from

        items.append(dict(url=url, year=year, file_id=file_id, p_from=p_from, p_to=p_to, note=note))
    return items


def find_settlements(text, mode_id=DEFAULT_MODE):
    settlements = compile_mode(mode_id)["settlements"]
    t = norm(text)
    return {name for name, rx in settlements.items() if rx.search(t)}


def match_header(line, nxt="", mode_id=DEFAULT_MODE):
    """('own', имя) / ('other', имя) / None."""
    c_ = compile_mode(mode_id)
    # длинная строка-«титул» («…Московской Губерніи и уѣзда Села троицкаго Здеревнями…»): берём первое село/деревню
    if len(line) > HEADER_MAX_LEN and re.search("(?:уезд|губерн)", norm(line)):
        n = norm(line)
        m = re.search(r"(?:уезда|губернии)\s+(?:и\s+уезда\s+)?(?:сел[а-я]*|дер[а-я]*)\.?\s*[\[\(«\"]?\s*(?:[а-я]+\s+){0,1}", n)
        best, best_pos = None, 1e9
        if m:
            for name, rx in c_["settlements"].items():
                mm = rx.search(n)
                if mm and mm.start() < best_pos:
                    best, best_pos = name, mm.start()
            if best:
                return "own", best
    cand = [line]
    # заголовок мог разорваться на две строки: "Деревня" / "Острогина".
    # Склеиваем ТОЛЬКО если текущая строка — одно слово-префикс.
    if nxt and c_["prefix_only"].match(line):
        cand.append(line + " " + nxt)
    for c in cand:
        if len(c) > HEADER_MAX_LEN:
            continue
        n = norm(c)
        for name, rx in c_["settlements"].items():
            if rx.search(n):
                return "own", name
        m = c_["other_header"].search(c)
        if m:
            return "other", m.group(1)
    return None


def extract_people(sheet_text, state, strictly_inside, mode_id=DEFAULT_MODE):
    """state — переносится между страницами: current, other, started, household (+ служебные поля разбора)."""
    mode = get_mode(mode_id)
    c_ = compile_mode(mode_id)
    lines = [s.strip() for s in re.split(r"\r?\n", sheet_text) if s.strip()]
    people = []
    skip_next = False
    for i, ln in enumerate(lines):
        if skip_next:                       # вторая половина разорванного заголовка
            skip_next = False
            continue
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        h = match_header(ln, nxt, mode_id)
        if h:
            kind, name = h
            state["current"] = name if kind == "own" else None
            state["other"] = name if kind == "other" else None
            state["household"] = None
            skip_next = bool(c_["prefix_only"].match(ln))
            continue
        if mode["household"]:
            hm = c_["hh_a"].match(ln) or c_["hh_b"].match(ln)
            if hm:
                state["household"] = int(hm.group(1))
                continue
        if len(ln) < mode["min_line_len"]:
            continue
        if is_artifact(ln):                 # шапка, итоги, колонки «Показания»
            state["dropped"] = state.get("dropped", 0) + 1
            continue
        if TITLE_NOISE.search(ln) and not TITLE_NOISE_EXC.search(ln):
            continue
        cur = state.get("current")
        if cur is None:
            # заголовка своего НП ещё не было: берём только внутри диапазона
            # и только если не встретился чужой НП
            if state.get("other") or not strictly_inside or not state.get("started"):
                continue
            cur = "не определён (начало диапазона)"
        if state.get("famFor") != cur:       # новый населённый пункт — счётчики семей и людей с нуля
            state["famFor"] = cur
            state["ctx"] = {"fam": {"n": 0, "members": [], "anchor": None}}
            state["seqM"] = 0
            state["seqF"] = 0
            state["lastLine"] = []
        res = parse_line(ln, state["ctx"])
        # события («умер», «бежал», «куплен», «в рекруты»...) относятся к людям предыдущей строки
        if res["events"]:
            prev = res["persons"] if res["persons"] else (state.get("lastLine") or [])
            if prev:
                if len(res["events"]) == len(prev):
                    for k, p in enumerate(prev):
                        p["event"] = p["event"] + "; " + res["events"][k] if p["event"] else res["events"][k]
                else:
                    p0 = prev[0]
                    p0["event"] = (p0["event"] + "; " if p0["event"] else "") + "; ".join(res["events"])
                    if len(prev) > 1:
                        p0["flags"].append("событие отнесено к первому из " + str(len(prev)) + " (проверить)")
        if not res["persons"]:
            continue
        state["lastLine"] = res["persons"]
        single = len(res["persons"]) == 1
        for p in res["persons"]:
            if p["gender"] == "м":
                state["seqM"] += 1
            elif p["gender"] == "ж":
                state["seqF"] += 1
            ages = res["ages"]
            age1 = ages[0] if single and len(ages) > 0 else ""
            age2 = ages[1] if single and len(ages) > 1 else ""
            people.append({
                "settlement": cur,
                "household": state.get("household") or p["famNo"] or "",
                "line_no": i + 1,
                "ages": ", ".join(str(x) for x in ages),
                "age1": age1 if mode["two_ages"] else "",
                "age2": age2 if mode["two_ages"] else "",
                "age": (age2 if age2 != "" else age1) if single else "",
                "text": " ".join(p["raw"]),
                "source": ln,
                "name": p["name"], "patr": p["patr"], "surname": p["surname"], "gender": p["gender"],
                "rel": p["rel"], "relTo": p["relTo"], "status": p["status"],
                "head": bool(p["head"]),
                "seq_no": state["seqF"] if p["gender"] == "ж" else state["seqM"],
                "_p": p,       # событие и примечание дописываются позже — читаются при записи в Excel
            })
    return people


def person_event(r):
    return r["_p"]["event"]


def person_note(r):
    return "; ".join(r["_p"]["flags"])


# ============================================================================
# JS, выполняемый внутри страницы архива (page-scripts.js): окно распознанного текста
# справа от окна со сканом + строка даты документа выше окна со сканом
# ============================================================================
EXTRACT_JS = r"""
async (opts) => {
  const textSel = opts.textSel || '';
  const dateSel = opts.dateSel || '';
  const maxWait = opts.maxWait || 15000;

  const rectOf = (e) => e.getBoundingClientRect();
  const visible = (e) => {
    const cs = getComputedStyle(e);
    if (cs.display === 'none' || cs.visibility === 'hidden') return false;
    const r = rectOf(e);
    return r.width > 0 && r.height > 0;
  };
  const clean = (s) => (s || '').replace(/[ \t ]+/g, ' ').replace(/\s*\n\s*/g, '\n').trim();
  const MONTHS = /январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр/i;

  // окно со сканом: самый большой canvas/img/svg
  const findViewer = () => {
    const list = [...document.querySelectorAll('canvas, img, svg')]
      .filter((e) => visible(e))
      .map((e) => ({ e, r: rectOf(e) }))
      .filter((x) => x.r.width > 200 && x.r.height > 200)
      .sort((a, b) => b.r.width * b.r.height - a.r.width * a.r.height);
    if (list.length) return list[0].r;
    return null;
  };

  const bySelector = (sel) => {
    if (!sel) return null;
    let els;
    try { els = [...document.querySelectorAll(sel)]; } catch (e) { return null; }
    return els.length ? els : null;
  };

  // окно распознанного текста: элемент справа от окна со сканом с самым большим текстом;
  // из равнозначных берётся самый вложенный (самый маленький по площади)
  const autoText = () => {
    const vr = findViewer();
    const minLeft = vr ? vr.right - 20 : window.innerWidth * 0.4;
    const skip = new Set(['SCRIPT', 'STYLE', 'CANVAS', 'IMG', 'SVG', 'NOSCRIPT', 'HEAD', 'HTML', 'BODY']);
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
    return { el: good[0].e, viewer: !!vr };
  };

  const getText = () => {
    const sel = bySelector(textSel);
    if (sel) return { text: sel.map((e) => e.innerText || '').filter((s) => s.trim()).join('\n'), how: 'вручную', found: true };
    const a = autoText();
    if (a) return { text: a.el.innerText || '', how: a.viewer ? 'авто' : 'авто (окно скана не найдено, взята правая часть)', found: true };
    return { text: '', how: '', found: false };
  };

  // дата документа: строка выше окна со сканом (над панелью инструментов), содержащая год
  const autoDate = () => {
    const vr = findViewer();
    if (!vr) return '';
    const cands = [];
    for (const e of document.body.querySelectorAll('*')) {
      const r = rectOf(e);
      if (r.width < 20 || r.height < 8 || r.bottom > vr.top + 4) continue;
      if (r.right < vr.left || r.left > vr.right) continue;
      if (vr.top - r.bottom > 300) continue;
      if (!visible(e)) continue;
      const t = clean(e.innerText);
      if (t.length < 4 || t.length > 200 || t.includes('\n')) continue;
      if (!/(?<!\d)1[5-9]\d\d(?!\d)/.test(t)) continue;
      cands.push({ e, t, bottom: r.bottom, score: (MONTHS.test(t) ? 2 : 0) + (/г\.|года|год/i.test(t) ? 1 : 0) });
    }
    // оставляем самые вложенные (без вложенных кандидатов)
    const leaf = cands.filter((c) => !cands.some((o) => o !== c && c.e.contains(o.e)));
    leaf.sort((a, b) => b.score - a.score || b.bottom - a.bottom);
    return leaf.length ? leaf[0].t : '';
  };

  const getDate = () => {
    const sel = bySelector(dateSel);
    if (sel) return { date: clean(sel.map((e) => e.innerText || '').join(' ')), how: 'вручную' };
    const d = autoDate();
    return { date: d, how: d ? 'авто' : '' };
  };

  // ждём, пока текст распознанной части прогрузится и перестанет меняться
  const t0 = Date.now();
  let last = -1, stable = 0, res = getText();
  while (Date.now() - t0 < maxWait) {
    res = getText();
    const len = res.text.trim().length;
    if (res.found && len > 0 && len === last) {
      if (++stable >= 3) break;
    } else {
      stable = 0;
    }
    if (res.found && len === 0 && Date.now() - t0 > 6000) break; // пустая страница (чистый лист)
    last = len;
    await new Promise((r) => setTimeout(r, 600));
  }
  const d = getDate();
  return {
    found: res.found,
    text: res.found ? res.text : '',
    textHow: res.how,
    date: d.date,
    dateHow: d.how,
    html: opts.wantHtml ? document.documentElement.outerHTML : null,
  };
}
"""


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


def ask_mode():
    """Выбор типа обработки (в расширении — переключатель в разделе 1)."""
    ids = list(MODES)
    print("Тип обработки:")
    for i, mid in enumerate(ids, 1):
        print(f"  {i}. {MODES[mid]['label']} — {MODES[mid]['description']}")
    while True:
        s = input(f"Выберите номер [1-{len(ids)}] (Enter = {MODES[DEFAULT_MODE]['label']}): ").strip()
        if not s:
            return DEFAULT_MODE
        if s.isdigit() and 1 <= int(s) <= len(ids):
            return ids[int(s) - 1]


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
    """Открывает страницу n документа, возвращает dict: url, text, date, found, err."""
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
            res = page.evaluate(EXTRACT_JS, {"textSel": TEXT_SELECTOR, "dateSel": DATE_SELECTOR,
                                             "maxWait": RENDER_WAIT_MS, "wantHtml": want_html}) or {}
            if want_html and res.get("html"):
                os.makedirs(DEBUG_DIR, exist_ok=True)
                with open(os.path.join(DEBUG_DIR, "first_page.html"), "w", encoding="utf-8") as f:
                    f.write(res["html"])
                _debug_saved = True
                print(f"Сохранён отладочный файл {DEBUG_DIR}/first_page.html")
            if res.get("found"):
                print(f"  стр. {n}: окно текста — {res.get('textHow')}, знаков: {len(res.get('text') or '')}"
                      + (f"; дата: «{res['date']}» ({res.get('dateHow')})" if res.get("date") else "; дата не найдена"))
            return {"url": url, "text": res.get("text") or "", "date": res.get("date") or "",
                    "found": bool(res.get("found")),
                    "err": "" if res.get("found") else "окно распознанного текста не найдено"}
        except Exception as e:  # noqa: BLE001
            last_err = str(e) or type(e).__name__
            print(f"  страница {n}: ошибка «{last_err}», попытка {attempt + 1}/{MAX_RETRIES + 1}")
            time.sleep(2)
    return {"url": url, "text": "", "date": "", "found": False, "err": last_err or "ошибка"}


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


def save_workbook(path, mode_id, people, sheets, links):
    revision = get_mode(mode_id)["id"] == "revision"
    wb = Workbook()
    bold = Font(bold=True)

    ws = wb.active
    ws.title = "Люди"
    ws.append(["Год", "Тип документа", "Дата документа (со страницы)", "Дата в тексте страницы", "Населённый пункт",
               "Семья № (по порядку)" if revision else "Двор № (по порядку)", "Страница", "Строка на странице",
               "Имя", "Отчество", "Фамилия", "Пол", "Связь", "С кем связан", "Статус", "Событие",
               "№ мужчины" if revision else "№ мужчины (по порядку)", "№ женщины",
               "Возраст", "Год рождения", "Числа-возрасты (черновик)",
               "Возраст: 1-е число" if revision else "", "Возраст: 2-е число" if revision else "",
               "Распознанная строка", "Исходная строка (целиком)", "Примечание разбора", "Ссылка на страницу"])
    for r in people:
        gender = "мужчина" if r["gender"] == "м" else "женщина" if r["gender"] == "ж" else ""
        ws.append([_c(x) for x in (
            r["year"], r["doc_type"], r["doc_date"], r["text_date"], r["settlement"], r["household"],
            r["page"], r["line_no"], r["name"], r["patr"], r["surname"], gender,
            r["rel"], r["relTo"], r["status"], person_event(r),
            r["seq_no"] if r["gender"] == "м" else "", r["seq_no"] if r["gender"] == "ж" else "",
            r["age"], r["birth_year"], r["ages"], r["age1"], r["age2"],
            r["text"], r["source"], person_note(r), r["url"] if r["head"] else "")])

    ws2 = wb.create_sheet("Страницы")
    ws2.append(["Год", "Тип документа", "Страница", "Дата документа (со страницы)",
                "Найденные НП округи", "Ссылка", "Распознанный текст страницы"])
    for r in sheets:
        ws2.append([_c(x) for x in (r["year"], r["doc_type"], r["page"], r["doc_date"],
                                    r["found"], r["url"], r["text"])])

    ws3 = wb.create_sheet("Ссылки")
    ws3.append(["Год", "Тип документа", "Исходная ссылка", "Диапазон из ссылки", "Просмотрено страниц",
                "Найдено НП (шт.)", "Какие НП", "Решение", "Людей собрано",
                "Дата документа (1-я страница)", "Примечание"])
    for r in links:
        ws3.append([_c(x) for x in r])

    for sh in (ws, ws2, ws3):
        sh.freeze_panes = "A2"
        for cell in sh[1]:
            cell.font = bold
        _widths(sh)
    wb.save(path)


def _year_of(s):
    mm = re.search(r"(?<!\d)(1[5-9]\d\d)(?!\d)", str(s or ""))
    return int(mm.group(1)) if mm else None


def main():
    html_path, xlsx_path = ask_paths()
    mode_id = ask_mode()
    m = get_mode(mode_id)
    items = parse_links(html_path, mode_id)
    if not items:
        sys.exit("В файле не найдено ссылок Яндекс Архива.")
    print(f"Ссылок Яндекс Архива: {len(items)}. Тип обработки: {m['label']}.")

    people, sheet_rows, link_rows = [], [], []
    page_cache = {}          # (file_id, page) -> результат fetch_sheet
    seen_people = set()
    stop = False
    misses = 0

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
            print(f"Окно текста: {TEXT_SELECTOR or 'авто'}; дата: {DATE_SELECTOR or 'авто'}.")

            for idx, it in enumerate(items, 1):
                if stop:
                    break
                year = it["year"]
                tag = f"[{idx}/{len(items)}] {year if year is not None else '—'}"
                rng = f"{it['p_from']}–{it['p_to']}" if it["p_from"] is not None else "—"

                if not it["file_id"] or it["p_from"] is None:
                    link_rows.append([year, m["label"], it["url"], rng, 0, 0, "",
                                      "пропущена: нет fileId или диапазона", 0, "", it["note"]])
                    print(f"{tag}: пропущена (нет fileId/диапазона)")
                    continue

                lo, hi = it["p_from"], it["p_to"]
                if hi - lo + 1 > BIG_RANGE:
                    print(f"{tag}: внимание, диапазон очень большой ({hi - lo + 1} страниц)")

                # 1. загрузка страниц диапазона из ссылки (диапазон — ровно из ссылки)
                texts = {}
                try:
                    for n in range(lo, hi + 1):
                        if stop:
                            break
                        k = (it["file_id"], n)
                        if k not in page_cache:
                            print(f"{tag}: загружаю страницу {n} (диапазон {lo}–{hi})...")
                            r = fetch_sheet(page, it["file_id"], n)
                            page_cache[k] = r
                            if not r["found"]:
                                misses += 1
                                if misses >= MAX_MISSES:
                                    stop = True
                                    print(f"Окно распознанного текста не найдено на {MAX_MISSES} страницах подряд — "
                                          "останавливаюсь, чтобы не собирать лишнее. Задайте TEXT_SELECTOR "
                                          "в начале скрипта и запустите снова.")
                            else:
                                misses = 0
                            if n < hi:
                                time.sleep(random.uniform(*DELAY_RANGE))
                        texts[n] = page_cache[k]
                except KeyboardInterrupt:
                    print("\nОстановка запрошена — сохраняю собранное...")
                    stop = True

                # 2. сбор людей
                found_all, per_page = set(), {}
                for n, r in texts.items():
                    f = find_settlements(r["text"], mode_id)
                    per_page[n] = f
                    found_all |= f

                state = {"current": None, "other": None, "started": True, "household": None}
                collected = 0
                first_date = ""
                for n in sorted(texts):
                    r = texts[n]
                    if not first_date and r["date"]:
                        first_date = r["date"]
                    sheet_rows.append({"year": year, "doc_type": m["label"], "page": n, "doc_date": r["date"],
                                       "found": "; ".join(sorted(per_page[n])), "url": r["url"],
                                       "text": f"ОШИБКА: {r['err']}" if r["err"] else r["text"]})
                    for p in extract_people(r["text"], state, True, mode_id):
                        pk = (it["file_id"], n, p["line_no"], p["name"], p["rel"], p["relTo"], p["text"])
                        if pk in seen_people:
                            continue
                        seen_people.add(pk)
                        tdate = text_date(r["text"])
                        doc_year = year or _year_of(r["date"]) or _year_of(tdate) or _year_of(it["url"])
                        p.update(year=doc_year or "", doc_type=m["label"], doc_date=r["date"], text_date=tdate,
                                 page=n, url=r["url"],
                                 birth_year=(doc_year - p["age"]) if doc_year and p["age"] != "" else "")
                        people.append(p)
                        collected += 1

                total = hi - lo + 1
                verdict = f"ПРЕРВАНА (страниц {len(texts)} из {total})" if len(texts) < total else "ОБРАБОТАНА"
                link_rows.append([year, m["label"], it["url"], rng, len(texts), len(found_all),
                                  "; ".join(sorted(found_all)), verdict, collected, first_date, it["note"]])
                print(f"{tag}: диапазон {rng} — НП {len(found_all)}, людей: {collected}")
                save_workbook(xlsx_path, mode_id, people, sheet_rows, link_rows)
        except KeyboardInterrupt:
            print("\nОстановлено пользователем.")
        finally:
            try:
                save_workbook(xlsx_path, mode_id, people, sheet_rows, link_rows)
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
