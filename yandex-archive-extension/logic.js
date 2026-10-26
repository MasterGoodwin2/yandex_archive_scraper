// Логика разбора (перенос из yandex_archive_scraper.py) + режимы обработки документов.
// В JS \b и \w не понимают кириллицу, поэтому границы слов заданы через Unicode-классы.
(function (root) {
  const ARCHIVE_HOSTS = ['ya.ru', 'yandex.ru'];
  const HEADER_MAX_LEN = 70;   // строка короче этого может быть заголовком населённого пункта

  // ------------------------------------------------------------------
  // РЕЖИМЫ ОБРАБОТКИ. Здесь меняются параметры сбора под тип документа.
  // ------------------------------------------------------------------
  const MODES = {
    confession: {
      id: 'confession',
      label: 'Исповедные ведомости',
      // слова-префиксы заголовка населённого пункта (кроме "сел…" и "дер…")
      extraPrefixes: [],
      minLineLen: 6,        // минимальная длина строки с человеком
      twoAges: false,       // в строке один возраст
      household: false,     // номера дворов не отслеживаем
      defaultSpan: 15,      // если в ссылке нет sheet_page_to: from .. from+defaultSpan
      description:
        'Заголовки населённых пунктов: село / деревня. В строке один возраст. ' +
        'Люди собираются по строкам под заголовком вашего населённого пункта.',
    },
    revision: {
      id: 'revision',
      label: 'Ревизские сказки',
      extraPrefixes: ['починок', 'починк', 'выселок', 'выселк', 'слобод', 'хутор', 'пустош', 'погост'],
      minLineLen: 5,        // записи короче: «сын Иван 5»
      twoAges: true,        // обычно два возраста: по прошлой ревизии и ныне
      household: true,      // отслеживаем «Двор № 5» / «№ 12»
      defaultSpan: 15,
      description:
        'Заголовки: также починок, выселок, слобода, хутор и т. п. Отслеживается номер двора ' +
        '(«Двор № 5», «№ 12»). В строке обычно два возраста — они выводятся в отдельные колонки. ' +
        'Допускаются более короткие строки.',
    },
  };
  const DEFAULT_MODE = 'confession';
  const getMode = (id) => MODES[id] || MODES[DEFAULT_MODE];

  // строки-заголовки документа, которые не являются людьми
  const TITLE_NOISE = /(?:ревизск|исповедн|ведомост|сказк|уезд|губерн|волост|приход)/i;

  const L = '\\p{L}\\p{N}_';
  const WB = `(?<![${L}])`;    // начало слова
  const WE = `(?![${L}])`;     // конец слова

  // слово → регулярка, нечувствительная к регистру (без флага i, чтобы [А-Я] оставалось строгим)
  const ci = (w) => [...w].map((ch) => (ch.toUpperCase() === ch ? ch : `[${ch.toUpperCase()}${ch}]`)).join('');

  // Населённые пункты округи: название -> основа (по нормализованному тексту)
  const SETTLEMENT_TAILS = {
    'село Троицкое':     'троицк',
    'деревня Острогина': 'острогин',
    'деревня Черепкова': 'черепков',
    'деревня Рублева':   'рублев',
    'деревня Луки':      'лук(?:и)?' + WE,
    'деревня Мякинина':  'м[яе]кинин',
  };

  const compiled = {};
  function compile(modeId) {
    const m = getMode(modeId);
    if (compiled[m.id]) return compiled[m.id];
    const words = ['сел', 'дер', ...m.extraPrefixes].join('|');
    const wordsCI = ['сел', 'дер', ...m.extraPrefixes].map(ci).join('|');

    // Обязателен префикс (в т.ч. "дер." / "д." / "с."). Применяется к нормализованному тексту.
    const prefix = `(?:${WB}(?:${words})[${L}]*\\.?|${WB}[дс]\\.)\\s*[\\[\\(«"]?\\s*`;
    const settlements = {};
    for (const [name, tail] of Object.entries(SETTLEMENT_TAILS)) settlements[name] = new RegExp(prefix + tail, 'u');

    // "Деревня"/"Дер." с любой буквы; сокращения "д."/"с." только строчные,
    // чтобы не принять за заголовок инициалы ("Д. Иванов").
    const otherHeader = new RegExp(
      `(?:${WB}(?:${wordsCI})[${L}]*\\.?|${WB}[дс]\\.)\\s*[\\[\\(«"]?\\s*([А-ЯЁІѢ][а-яёіѣъь]{3,})`, 'u');
    // строка, состоящая только из слова-префикса ("Деревня" / "Село" / "д.")
    const prefixOnly = new RegExp(`^\\s*(?:(?:${wordsCI})[${L}]*|[дсДС])\\.?\\s*$`, 'u');

    // маркеры двора: "Двор № 5 ..." или строка "№ 12"
    const dvor = ci('двор');
    const hhA = new RegExp(`^\\s*(?:${dvor}[${L}]*|[Дд][Вв])\\.?\\s*(?:№|N|n)?\\s*(\\d{1,3})(?![${L}])`, 'u');
    const hhB = /^\s*(?:№|N)\s*(\d{1,3})\s*[.:]?\s*$/u;

    compiled[m.id] = { settlements, otherHeader, prefixOnly, hhA, hhB };
    return compiled[m.id];
  }

  function norm(s) {
    s = s.toLowerCase();
    for (const [a, b] of [['ё', 'е'], ['ѣ', 'е'], ['і', 'и'], ['ѵ', 'и'], ['ѳ', 'ф']]) {
      s = s.split(a).join(b);
    }
    return s.split('ъ').join('').split('ь').join('');
  }

  function toInt(s) {
    if (s === null || s === undefined || s === '') return null;
    const n = parseInt(s, 10);
    return Number.isNaN(n) ? null : n;
  }

  // Текст элемента как в BeautifulSoup get_text("", strip=True):
  // куски текста обрезаются и склеиваются без пробелов (для "разорванных" ссылок <a>186</a>4)
  function strippedText(node) {
    let out = '';
    const walk = (n) => {
      for (const ch of n.childNodes || []) {
        if (ch.nodeType === 3) out += (ch.nodeValue || '').trim();
        else if (ch.nodeType === 1) walk(ch);
      }
    };
    walk(node);
    return out;
  }

  // doc — Document (DOMParser) со списком ссылок
  function parseLinks(doc, modeId) {
    const span = getMode(modeId).defaultSpan;
    const items = [];
    const seen = new Set();
    for (const a of doc.querySelectorAll('a[href]')) {
      const url = (a.getAttribute('href') || '').trim();
      let u;
      try { u = new URL(url); } catch (e) { continue; }
      if (!ARCHIVE_HOSTS.includes(u.hostname) || !u.pathname.startsWith('/archive')) continue;
      if (seen.has(url)) continue;
      seen.add(url);

      const holder = a.closest('p') || a;
      const m = strippedText(holder).match(/(?<!\d)(1[5-9]\d\d)(?!\d)/);
      const year = m ? parseInt(m[1], 10) : null;

      let note = '';
      let fileId = null, pFrom = null, pTo = null;

      if (u.pathname.startsWith('/archive/search')) {
        fileId = u.searchParams.get('fileId');
        pFrom = toInt(u.searchParams.get('sheet_page_from'));
        pTo = toInt(u.searchParams.get('sheet_page_to'));
      } else if (u.pathname.startsWith('/archive/catalog/')) {
        const parts = u.pathname.split('/').filter(Boolean); // archive / catalog / <id> / <page>
        if (parts.length >= 4) {
          fileId = parts[2];
          pFrom = pTo = toInt(parts[3]);
          note = 'ссылка на одну страницу каталога, диапазона в ссылке нет';
        }
      }

      if (pFrom !== null && pTo === null) {
        pTo = pFrom + span;
        note = `нет sheet_page_to — взято from+${span}`;
      }
      if (pFrom !== null && pTo !== null && pFrom > pTo) [pFrom, pTo] = [pTo, pFrom];

      items.push({ url, year, fileId, pFrom, pTo, note });
    }
    return items;
  }

  function findSettlements(text, modeId) {
    const { settlements } = compile(modeId);
    const t = norm(text);
    return new Set(Object.keys(settlements).filter((name) => settlements[name].test(t)));
  }

  // Если строка — заголовок населённого пункта: ['own', имя] / ['other', имя] / null
  function matchHeader(line, nxt, modeId) {
    const { settlements, otherHeader, prefixOnly } = compile(modeId);
    const cand = [line];
    // заголовок мог разорваться на две строки: "Деревня" / "Острогина".
    // Склеиваем ТОЛЬКО если текущая строка — одно слово-префикс.
    if (nxt && prefixOnly.test(line)) cand.push(line + ' ' + nxt);
    for (const c of cand) {
      if (c.length > HEADER_MAX_LEN) continue;
      const n = norm(c);
      for (const [name, rx] of Object.entries(settlements)) {
        if (rx.test(n)) return ['own', name];
      }
      const m = c.match(otherHeader);
      if (m) return ['other', m[1]];
    }
    return null;
  }

  // state: {current, other, started, household} — переносится между страницами
  function extractPeople(sheetText, state, strictlyInside, modeId) {
    const mode = getMode(modeId);
    const { prefixOnly, hhA, hhB } = compile(modeId);
    const lines = sheetText.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    const people = [];
    let skipNext = false;
    for (let i = 0; i < lines.length; i++) {
      if (skipNext) { skipNext = false; continue; } // вторая половина разорванного заголовка
      const ln = lines[i];
      const nxt = i + 1 < lines.length ? lines[i + 1] : '';
      const h = matchHeader(ln, nxt, modeId);
      if (h) {
        const [kind, name] = h;
        state.current = kind === 'own' ? name : null;
        state.other = kind === 'other' ? name : null;
        state.household = null;
        skipNext = prefixOnly.test(ln);
        continue;
      }
      if (mode.household) {
        const hm = ln.match(hhA) || ln.match(hhB);
        if (hm) { state.household = parseInt(hm[1], 10); continue; }
      }
      if (ln.length < mode.minLineLen || !/[А-ЯЁІѢ][а-яёіѣъь]+/.test(ln)) continue;
      if (TITLE_NOISE.test(ln)) continue;   // служебные строки: название документа, уезд, приход...
      let cur = state.current;
      if (cur === null || cur === undefined) {
        // заголовка своего НП ещё не было: берём только внутри диапазона
        // и только если не встретился чужой НП
        if (state.other || !strictlyInside || !state.started) continue;
        cur = 'не определён (начало диапазона)';
      }
      const ages = [...ln.matchAll(/(?<![\p{L}\p{N}_])(\d{1,3})(?![\p{L}\p{N}_])/gu)]
        .map((x) => parseInt(x[1], 10))
        .filter((x) => x > 0 && x < 111);
      people.push({
        settlement: cur,
        household: mode.household && state.household ? state.household : '',
        lineNo: i + 1,
        ages: ages.join(', '),
        age1: mode.twoAges && ages.length > 0 ? ages[0] : '',
        age2: mode.twoAges && ages.length > 1 ? ages[1] : '',
        text: ln,
      });
    }
    return people;
  }

  const api = {
    ARCHIVE_HOSTS, MODES, DEFAULT_MODE, getMode,
    norm, parseLinks, findSettlements, matchHeader, extractPeople,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Logic = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
