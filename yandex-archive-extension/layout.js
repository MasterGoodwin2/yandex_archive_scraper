// Сопоставление распознанных блоков по их положению на странице:
//   № двора  ↔  № мужчины / женщины  ↔  текст записи  ↔  возраст.
// Текст записи НЕ разбирается, не исправляется и не достраивается — сохраняется ровно так, как распознан.
// Блок = {t: текст, x, y, w, h} — прямоугольник на странице (в пикселях).
(function (root) {
  const Logic = root.Logic || (typeof require !== 'undefined' ? require('./logic.js') : null);

  const NUM_RX = /^(?:№|N|n)?\s*(\d{1,3})\s*[.,)]?$/u;            // «5», «5.», «№5»
  const LEAD_NUM_RX = /^(\d{1,3})\s*[.)]\s*(\D.*)$/u;               // «5. Иван Петров»
  const TRAIL_AGE_RX = /^(.*[^\d\s№])\s+(\d{1,3})$/u;               // «Иван Петров 45»
  const MAX_AGE = 110;

  // служебные строки шапки и итогов (применяются только к строкам БЕЗ чисел-возрастов, кроме «итого»)
  const ARTIFACT = new RegExp('^(?:' + [
    'лета', 'семьи', 'семья', 'ныне', 'нине', 'на лицо', 'налицо', 'крестьян', 'мужеск', 'женск', 'мужск', 'итого', 'по последне',
    'по после', 'ис того', 'из того', 'во временно', 'с которого', 'скоторого', 'скоторо', 'когда', 'кто были', 'которые',
    'кто же', 'ктож', 'показани', 'числа', 'состояли', 'состояло', 'выбыли', 'подано', 'ревизск', 'ревизия', 'сказка',
    'оборот', 'попов$', 'исповедь', 'у исповеди', 'и святого', 'исповедались', 'токмо', 'причастились', 'домов', 'двор$',
    'число$', 'людеи', 'людей', 'мужеска', 'женска', 'всего', 'прибыло', 'убыло', 'выбыло', 'в том числе', 'с прибылыми',
    'по ревизии', 'во время', 'родилось', 'умерло', 'имена', 'и отчества', 'пол$', 'номер', 'возраст', 'лет$',
  ].join('|') + ')');
  const ALWAYS_ARTIFACT = /^(?:итого|всего|по ревизии|в том числе)/;
  const TITLE_NOISE = /(?:ревизск|исповедн|ведомост|сказк|уезд|губерн|волост|приход)/;
  const kline = (s) => Logic.norm(s).replace(/[^а-яa-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();

  const median = (arr, dflt) => {
    if (!arr.length) return dflt;
    const a = arr.slice().sort((x, y) => x - y);
    return a[Math.floor(a.length / 2)];
  };

  function newState() {
    return { current: null, other: null, fam: null, cols: { m: null, f: null, yard: null }, dropped: 0, pages: 0, blockPages: 0, linePages: 0 };
  }

  // ---------------------------------------------------------------- семьи
  function flush(st, out) {
    if (st.fam && st.fam.persons.length) out.push(st.fam);
    st.fam = null;
  }
  function startFamily(st, out, yard, ctx) {
    flush(st, out);
    st.fam = { settlement: ctx.settlement, yard: yard === null ? '' : yard, persons: [], notes: [], url: ctx.url, page: ctx.page };
  }
  function ensureFamily(st, out, ctx) {
    if (!st.fam || st.fam.settlement !== ctx.settlement) {
      startFamily(st, out, null, ctx);
      st.fam.noYard = true;
    }
    return st.fam;
  }
  const addNote = (fam, msg) => { if (!fam.notes.includes(msg)) fam.notes.push(msg); };

  // подходит ли строка под «населённый пункт округи / чужой / нет»; обновляет состояние
  function applyHeader(st, out, h) {
    flush(st, out);
    const [kind, name] = h;
    st.current = kind === 'own' ? name : null;
    st.other = kind === 'other' ? name : null;
  }
  function settlementOf(st) {
    if (st.current) return st.current;
    if (st.other) return null;                       // чужой населённый пункт — не собираем
    return 'не определён (начало диапазона)';
  }

  // ------------------------------------------------------------ блоки → атомы
  // Блок делится на «число» и «текст», если число стоит в начале/в конце одной распознанной строки.
  function atomize(b) {
    const t = String(b.t || '').replace(/\s+/g, ' ').trim();
    if (!t) return [];
    const base = { y: b.y, h: b.h };
    const mk = (kind, text, x, w, v) => Object.assign({ kind, t: text, x, w, cx: x + w / 2, v }, base);
    let m = t.match(NUM_RX);
    if (m) return [mk('num', t, b.x, b.w, parseInt(m[1], 10))];
    m = t.match(LEAD_NUM_RX);
    if (m) {
      const k = m[1].length / t.length, wn = b.w * k;
      return [mk('num', m[1], b.x, wn, parseInt(m[1], 10)), mk('text', m[2].trim(), b.x + wn, b.w - wn)];
    }
    m = t.match(TRAIL_AGE_RX);
    if (m && parseInt(m[2], 10) > 0 && parseInt(m[2], 10) <= MAX_AGE) {
      const k = m[2].length / t.length, wn = b.w * k;
      return [mk('text', m[1].trim(), b.x, b.w - wn), mk('num', m[2], b.x + b.w - wn, wn, parseInt(m[2], 10))];
    }
    return [mk('text', t, b.x, b.w)];
  }

  function clusterRows(atoms) {
    const bs = atoms.slice().sort((a, b) => (a.y + a.h / 2) - (b.y + b.h / 2));
    const tol = Math.max(3, median(bs.map((a) => a.h), 10) * 0.5);
    const rows = [];
    for (const a of bs) {
      const cy = a.y + a.h / 2;
      const last = rows[rows.length - 1];
      if (last && Math.abs(cy - last.cy) <= tol) {
        last.atoms.push(a);
        last.cy = (last.cy * (last.atoms.length - 1) + cy) / last.atoms.length;
      } else rows.push({ cy, atoms: [a] });
    }
    rows.forEach((r) => r.atoms.sort((a, b) => a.x - b.x));
    return rows;
  }

  // ----------------------------------------------------------- сбор по блокам
  function processBlocks(blocks, st, ctx, modeId, out) {
    const atoms = [];
    for (const b of blocks) atoms.push(...atomize(b));
    if (!atoms.length) return;
    const rows = clusterRows(atoms);

    const minX = Math.min(...atoms.map((a) => a.x));
    const maxX = Math.max(...atoms.map((a) => a.x + a.w));
    const colTol = Math.max(20, (maxX - minX) * 0.04);

    // 1. шапки колонок страницы → центры колонок (запоминаются для следующих страниц)
    for (const r of rows) for (const a of r.atoms) {
      if (a.kind !== 'text') continue;
      const k = kline(a.t);
      if (/^(?:мужеск|мужск)/.test(k)) st.cols.m = a.cx;
      else if (/^женск/.test(k)) st.cols.f = a.cx;
      else if (/^(?:номер\s+)?(?:двор\w*|дом\w*)$/.test(k)) st.cols.yard = a.cx;
    }
    const textXs = atoms.filter((a) => a.kind === 'text' && !/^(?:мужеск|женск|мужск)/.test(kline(a.t))).map((a) => a.x).sort((p, q) => p - q);
    const textLeft = textXs.length ? textXs[Math.floor(textXs.length * 0.1)] : Infinity;

    const genderFor = (cx) => {
      const { m, f } = st.cols;
      if (m !== null && f !== null) return Math.abs(cx - m) <= Math.abs(cx - f) ? 'м' : 'ж';
      return '';
    };

    for (let ri = 0; ri < rows.length; ri++) {
      const row = rows[ri];
      let ats = row.atoms.slice();
      const joined = ats.map((a) => a.t).join(' ');
      const nextJoined = ri + 1 < rows.length ? rows[ri + 1].atoms.map((a) => a.t).join(' ') : '';

      // населённый пункт
      const h = Logic.matchHeader(joined, nextJoined, modeId);
      if (h) {
        applyHeader(st, out, h);
        if (Logic.isPrefixOnly(joined, modeId)) ri++;      // вторая половина разорванного заголовка
        continue;
      }
      const hasNum = ats.some((a) => a.kind === 'num');
      const k = kline(joined);
      const allHeads = ats.every((a) => { const ka = kline(a.t); return ARTIFACT.test(ka) || TITLE_NOISE.test(ka); });
      if (!hasNum && (ARTIFACT.test(k) || TITLE_NOISE.test(k) || allHeads)) { st.dropped++; continue; }
      if (ALWAYS_ARTIFACT.test(k)) { st.dropped++; continue; }
      // одиночные слова-шапки колонок («Двор», «Мужеск. пол») внутри строки с данными не удаляем
      const cur = settlementOf(st);
      if (cur === null) continue;
      ctx.settlement = cur;

      // 2. номер двора
      let yardNo = null;
      for (const a of ats) {
        if (a.kind !== 'text') continue;
        const y = Logic.matchYard(a.t, modeId);
        if (y) { yardNo = y.no; if (y.rest) { a.t = y.rest; } else a.dead = true; break; }
      }
      ats = ats.filter((a) => !a.dead);
      const nums = () => ats.filter((a) => a.kind === 'num' && !a.used);
      if (yardNo === null) {
        let cand = null;
        if (st.cols.yard !== null) cand = nums().find((a) => Math.abs(a.cx - st.cols.yard) <= colTol) || null;
        else {
          const texts = ats.filter((a) => a.kind === 'text');
          const firstText = texts.length ? texts[0].x : Infinity;
          const left = nums().filter((a) => a.cx < firstText);
          if (left.length >= 2) cand = left[0];
          else if (!texts.length && left.length === 1 && left[0].x < textLeft - colTol) cand = left[0];
        }
        if (cand) { yardNo = cand.v; cand.used = true; }
      } else {
        // номер двора найден по тексту «Двор № N» — число рядом с ним в той же строке не трогаем
      }
      if (yardNo !== null) startFamily(st, out, yardNo, ctx);

      // 3. люди строки
      const texts = ats.filter((a) => a.kind === 'text');
      if (!texts.length) {
        const lone = nums();
        if (lone.length && yardNo === null) {
          const fam = ensureFamily(st, out, ctx);
          addNote(fam, 'число без записи рядом (стр. ' + ctx.page + '): ' + lone.map((a) => a.v).join(', '));
        }
        continue;
      }
      const fam = ensureFamily(st, out, ctx);
      const people = texts.map((a) => ({ gender: genderFor(a.cx), no: '', text: a.t, age: '', atom: a }));

      // числа, лежащие между соседними записями (и до первой / после последней)
      const free = ats.filter((a) => a.kind === 'num' && !a.used);
      const gaps = people.map(() => []);               // gaps[i] — числа между записью i-1 и i (индекс 0 — до первой)
      const tail = [];
      for (const n of free) {
        let i = people.findIndex((p) => n.cx < p.atom.cx);
        if (i === -1) tail.push(n); else gaps[i].push(n);
      }
      const unmatched = [];
      const setAge = (p, n) => { if (n.v > 0 && n.v <= MAX_AGE) { p.age = n.v; return true; } return false; };
      people.forEach((p, i) => {
        const g = gaps[i];
        // слева от записи i: числа между записью i-1 и i
        const prev = i > 0 ? people[i - 1] : null;
        let own = [];                                  // числа, которые берёт запись i как свой номер
        let prevAge = null;                            // число, которое берёт запись i-1 как возраст
        if (!prev) own = g.slice(-1), unmatched.push(...g.slice(0, -1));
        else if (g.length === 1) {
          const n = g[0];
          const dPrev = n.cx - (prev.atom.x + prev.atom.w), dNext = p.atom.x - n.cx;
          if (prev.age === '' && dPrev <= dNext) prevAge = n; else own = [n];
        } else if (g.length >= 2) {
          prevAge = g[0]; own = [g[g.length - 1]]; unmatched.push(...g.slice(1, -1));
        }
        if (prevAge && !setAge(prev, prevAge)) unmatched.push(prevAge);
        if (own.length) { p.no = own[0].v; }
      });
      if (tail.length) {
        const last = people[people.length - 1];
        if (last.age === '' && setAge(last, tail[0])) unmatched.push(...tail.slice(1));
        else unmatched.push(...tail);
      }
      for (const p of people) {
        fam.persons.push({ gender: p.gender, no: p.no, text: p.text, age: p.age });
        if (!p.gender) addNote(fam, 'пол не определён: нет шапок «Мужеск.»/«Женск.» (стр. ' + ctx.page + ')');
      }
      if (unmatched.length) addNote(fam, 'несопоставленные числа (стр. ' + ctx.page + '): ' + unmatched.map((a) => a.v).join(', '));
    }
  }

  // ------------------------------------------------- запасной режим: только строки
  // Применяется, когда координат блоков получить не удалось: записи берутся строками как есть,
  // номер двора — по маркерам «Двор № N» / «№ N». Номер мужчины/женщины и возраст НЕ сопоставляются.
  function processLines(text, st, ctx, modeId, out) {
    const mode = Logic.getMode(modeId);
    const lines = String(text || '').split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i];
      const nxt = i + 1 < lines.length ? lines[i + 1] : '';
      const h = Logic.matchHeader(ln, nxt, modeId);
      if (h) { applyHeader(st, out, h); if (Logic.isPrefixOnly(ln, modeId)) i++; continue; }
      const y = Logic.matchYard(ln, modeId);
      const cur = settlementOf(st);
      if (cur === null) continue;
      ctx.settlement = cur;
      if (y) {
        startFamily(st, out, y.no, ctx);
        if (!y.rest) continue;
      }
      const body = y ? y.rest : ln;
      if (body.length < mode.minLineLen) continue;
      const k = kline(body);
      const hasNum = /\d/.test(body);
      if ((!hasNum && (ARTIFACT.test(k) || TITLE_NOISE.test(k))) || ALWAYS_ARTIFACT.test(k)) { st.dropped++; continue; }
      const fam = ensureFamily(st, out, ctx);
      addNote(fam, 'координаты блоков не получены: № и возраст не сопоставлены, записи даны строками как есть');
      fam.persons.push({ gender: '', no: '', text: body, age: '' });
    }
  }

  // «табличность» набора блоков: достаточно строк и не меньше 3 колонок (по x), в каждой не меньше 2 блоков
  function isTable(blocks, relaxed) {
    if (relaxed) return blocks.length >= 4 && new Set(blocks.map((b) => Math.round((b.y + b.h / 2) / 8))).size >= 2;
    if (blocks.length < 8) return false;
    if (new Set(blocks.map((b) => Math.round((b.y + b.h / 2) / 8))).size < 3) return false;
    const xs = blocks.map((b) => b.x).sort((a, b) => a - b);
    const gap = Math.max(20, (xs[xs.length - 1] - xs[0] + 1) * 0.04);
    let cols = 0, run = 1;
    for (let i = 1; i <= xs.length; i++) {
      if (i < xs.length && xs[i] - xs[i - 1] <= gap) run++;
      else { if (run >= 2) cols++; run = 1; }
    }
    return cols >= 3;
  }

  // page = {candidates: [{name, blocks}], text, url, page}. Берётся первый источник блоков, похожий на таблицу;
  // если такого нет — запасной режим по строкам текста.
  function processPage(page, st, modeId, out) {
    const ctx = { url: page.url, page: page.page, settlement: null };
    st.pages++;
    const cands = page.candidates || [{ name: 'блоки', blocks: page.blocks || [] }];
    for (const c of cands) {
      const blocks = (c.blocks || []).filter((b) => b && String(b.t || '').trim());
      // короткая страница (конец раздела) допускается, если прошлые страницы уже разобраны по этому же источнику блоков
      if (isTable(blocks, st.src === c.name)) {
        st.src = c.name;
        st.blockPages++;
        processBlocks(blocks, st, ctx, modeId, out);
        return 'блоки: ' + c.name;
      }
    }
    st.linePages++;
    processLines(page.text, st, ctx, modeId, out);
    return 'строки (координат нет)';
  }

  function finish(st, out) { flush(st, out); }

  // ------------------------------------------------------------ вывод в ячейку
  function formatPerson(p) {
    const label = (p.gender || '') + (p.no !== '' && p.no !== undefined ? '№' + p.no : '');
    return (label ? label + ' ' : '') + p.text + (p.age !== '' && p.age !== undefined ? ', возр. ' + p.age : '');
  }
  const familyCell = (fam) => fam.persons.map(formatPerson).join('\n');

  const api = { newState, processPage, finish, formatPerson, familyCell, atomize, clusterRows, isTable };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Layout = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
