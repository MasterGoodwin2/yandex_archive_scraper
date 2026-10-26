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
        'Заголовки: село / деревня. Собираются: двор (по порядку), № мужчины / женщины, возраст (один), ' +
        'вычисляется год рождения = год документа − возраст. Колонки «Показания действия» (у исповеди были / ' +
        'исповедались, но не причастились / не были) не собираются. Ссылка на страницу — только у хозяина двора.',
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
        'Заголовки: также починок, выселок, слобода, хутор и т. п. Каждый человек — в своей строке: имя, отчество ' +
        '(достраивается от отца), фамилия, пол (по имени), связь (жена, сын, брат...), статус (приёмыш, бобыль, вдова...) ' +
        'сохраняется; события (умер, бежал, куплен, в рекруты) — в колонку «Событие». Ссылка — только главе семьи. ' +
        'Старая орфография заменяется на современную, служебные строки шапки отбрасываются.',
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
    'деревня Мякинина':  '(?:м[яе]кинин|мемин)',
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
    // длинная строка-«титул» («…Московской Губерніи и уѣзда Села троицкаго Здеревнями…»): берём первое село/деревню
    if (line.length > HEADER_MAX_LEN && /(?:уезд|губерн)/.test(norm(line))) {
      const n = norm(line);
      const m = n.match(/(?:уезда|губернии)\s+(?:и\s+уезда\s+)?(?:сел[а-я]*|дер[а-я]*)\.?\s*[\[\(«"]?\s*(?:[а-я]+\s+){0,1}/);
      let best = null, bestPos = 1e9;
      if (m) {
        for (const [name, rx] of Object.entries(settlements)) {
          const mm = rx.exec(n);
          if (mm && mm.index < bestPos) { best = name; bestPos = mm.index; }
        }
        if (best) return ['own', best];
      }
    }
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
      if (ln.length < mode.minLineLen) continue;
      const Pz = root.Parser || (typeof require !== 'undefined' ? require('./parser.js') : null);
      if (Pz.isArtifact(ln)) { state.dropped = (state.dropped || 0) + 1; continue; }   // шапка, итоги, колонки «Показания»
      if (TITLE_NOISE.test(ln) && !/[а-яё]+\s+[а-яё]+\s+[а-яё]+\s+(?:сын|жена|дочь)/i.test(ln)) continue;
      let cur = state.current;
      if (cur === null || cur === undefined) {
        // заголовка своего НП ещё не было: берём только внутри диапазона
        // и только если не встретился чужой НП
        if (state.other || !strictlyInside || !state.started) continue;
        cur = 'не определён (начало диапазона)';
      }
      if (state.famFor !== cur) {                    // новый населённый пункт — счётчики семей и людей с нуля
        state.famFor = cur;
        state.ctx = { fam: { n: 0, members: [], anchor: null } };
        state.seqM = 0; state.seqF = 0; state.lastLine = [];
      }
      const res = Pz.parseLine(ln, state.ctx);
      // события («умер», «бежал», «куплен», «в рекруты»...) относятся к людям предыдущей строки
      if (res.events.length) {
        const prev = res.persons.length ? res.persons : (state.lastLine || []);
        if (prev.length) {
          if (res.events.length === prev.length) prev.forEach((p, k) => { p.event = p.event ? p.event + '; ' + res.events[k] : res.events[k]; });
          else {
            const p0 = prev[0];
            p0.event = (p0.event ? p0.event + '; ' : '') + res.events.join('; ');
            if (prev.length > 1) p0.flags.push('событие отнесено к первому из ' + prev.length + ' (проверить)');
          }
        }
      }
      if (!res.persons.length) continue;
      state.lastLine = res.persons;
      const single = res.persons.length === 1;
      for (const p of res.persons) {
        if (p.gender === 'м') state.seqM += 1; else if (p.gender === 'ж') state.seqF += 1;
        const ages = res.ages;
        const age1 = single && ages.length > 0 ? ages[0] : '';
        const age2 = single && ages.length > 1 ? ages[1] : '';
        people.push({
          settlement: cur,
          household: state.household || p.famNo || '',
          lineNo: i + 1,
          ages: ages.join(', '),
          age1: mode.twoAges ? age1 : '',
          age2: mode.twoAges ? age2 : '',
          age: single ? (age2 !== '' ? age2 : age1) : '',
          text: p.raw.join(' '),
          source: ln,
          name: p.name, patr: p.patr, surname: p.surname, gender: p.gender,
          rel: p.rel, relTo: p.relTo, status: p.status,
          get event() { return p.event; },
          head: !!p.head, seqNo: p.gender === 'ж' ? state.seqF : state.seqM,
          get note() { return p.flags.join('; '); },
        });
      }
    }
    return people;
  }

  const api = {
    ARCHIVE_HOSTS, textDate: (t) => (root.Parser || require('./parser.js')).textDate(t), MODES, DEFAULT_MODE, getMode,
    norm, parseLinks, findSettlements, matchHeader, extractPeople,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Logic = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
