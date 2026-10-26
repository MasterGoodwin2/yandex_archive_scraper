// Разбор строки с людьми: артефакты, имена/отчества, связи (жена/сын/брат...), статусы, события.
(function (root) {
  const N = root.Names || (typeof require !== 'undefined' ? require('./names.js') : null);

  // нормализованная строка с пробелами (современная орфография, без ь, только русские буквы)
  const kline = (s) => N.modern(s).toLowerCase().replace(/ё/g, 'е').replace(/ь/g, '').replace(/[^а-я ]/g, ' ').replace(/\s+/g, ' ').trim();

  // служебные строки шапки/итогов и колонок, которые не собираются
  const ARTIFACT = new RegExp('^(?:' + [
    'лета', 'семьи', 'семья', 'ныне', 'нине', 'на лицо', 'налицо', 'крестьян', 'мужеск', 'женск', 'итого', 'по последне',
    'по после', 'ис того', 'из того', 'во временно', 'с которого', 'скоторого', 'скоторо', 'когда', 'кто были', 'которые',
    'кто же', 'ктож', 'показани', 'числа', 'состояли', 'состояло', 'выбыли', 'подано', 'ревизск', 'ревизия', 'сказка',
    'оборот', 'попов$', 'исповедь', 'у исповеди', 'и святого', 'исповедались', 'токмо', 'причастились', 'домов', 'двор$',
    'число$', 'людеи', 'мужеска', 'женска', 'всего', 'прибыло', 'убыло', 'выбыло', 'в том числе', 'с прибылыми',
    'по ревизии', 'во время', 'родилось', 'умерло', 'имена', 'и отчества',
  ].join('|') + ')');

  const LINK_SKIP = new Set(['же', 'ж', 'того', 'тож', 'тоже', 'также', 'такожде', 'другои', 'второи', 'третии', 'меншои', 'малои',
    'большои', 'балшои', 'старшии', 'меньшии', 'из', 'от', 'на', 'во', 'по', 'при', 'при ним', 'году', 'года', 'лет', 'лета', 'ныне',
    'нине', 'отъ', 'его', 'ея', 'ее', 'их', 'ими', 'но', 'как', 'был', 'была', 'были', 'которои', 'котораи', 'первои', 'первои',
    'вновь', 'они', 'он', 'она', 'им', 'ему', 'ей', 'оного', 'сих', 'сеи', 'вдовы']);
  const SMALL = new Set(['и', 'в', 'с', 'к', 'о', 'у', 'а', 'я', 'во', 'по', 'на', 'за', 'из', 'от', 'до', 'же', 'ли', 'не', 'ни', 'ж', 'им', 'ей']);

  const isCapital = (w) => /^[А-ЯЁІѢ]/.test(w);

  function tokenize(line) {
    const out = [];
    const raw = N.modern(line).replace(/[“”«»"()\[\]{}]/g, ' ').split(/\s+/).filter(Boolean);
    for (const r0 of raw) {
      const colon = /[:;]$/.test(r0);
      const r = r0.replace(/^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$/gu, '');
      if (!r) continue;
      if (/^\d+$/.test(r)) { out.push({ type: 'num', v: parseInt(r, 10), w: r }); continue; }
      if (/\d/.test(r)) { out.push({ type: 'skip', w: r }); continue; }       // «2-й», «1810-го», «9½»
      const w = r.replace(/[^\p{L}\-]/gu, '').replace(/^-+|-+$/g, '');
      if (!w) continue;
      out.push({ type: 'word', w, colon });
    }
    return out;
  }

  function classify(t) {
    if (t.type !== 'word') return t;
    const w = t.w;
    const k = N.key(w);
    t.k = k;
    if (!k || k.length < 1) { t.type = 'skip'; return t; }
    if (N.NUMWORD_KEYS.has(k)) { t.type = 'skip'; return t; }
    if (k === 'его') { t.type = 'link'; t.who = 'м'; return t; }
    if (k === 'ея' || k === 'ее' || k === 'еи') { t.type = 'link'; t.who = 'ж'; return t; }
    if (N.REL[k]) { t.type = 'rel'; t.rel = N.REL[k]; return t; }
    if (k.length >= 3 && N.STATUS_KEYS.some((s) => k.startsWith(s))) {
      t.type = 'status'; return t;
    }
    if (LINK_SKIP.has(k) || SMALL.has(k) || w.length <= 1) { t.type = 'skip'; return t; }
    const nm = N.lookup(w);
    if (nm && (w.length >= 3)) { t.type = 'name'; t.nm = nm; t.nmGen = N.lookup(w, true); return t; }
    const p = N.patrParse(w);
    if (p) { t.type = 'pat'; t.p = p; return t; }
    t.type = isCapital(w) || w.length >= 4 ? 'unk' : 'skip';
    return t;
  }

  const pc = (w) => (w || '').split('-').map(N.cap).join('-');
  const modernWord = (w) => pc(N.modern(w).replace(/[^\p{L}\-]/gu, ''));

  // ctx: {fam:{n, members[], anchor}, lastLine:[persons]}
  function newPerson(ctx, o) {
    const p = Object.assign({ name: '', patr: '', surname: '', gender: '', rel: '', relTo: '', status: '', event: '', head: false,
      raw: [], flags: [], famNo: ctx.fam.n, anchor: null }, o);
    ctx.fam.members.push(p);
    return p;
  }

  function findAnchor(ctx, nameTok, patTok) {
    const fam = ctx.fam.members;
    const want = nameTok.nm.nom;
    const pk = patTok ? N.key(patTok.p.base) : null;
    for (let i = fam.length - 1; i >= 0; i--) {
      const m = fam[i];
      if (m.name === want && (!pk || N.key(m.patr.replace(/а$/, '')) === pk)) return m;
    }
    for (let i = fam.length - 1; i >= 0; i--) if (fam[i].name === want) return fam[i];
    return null;
  }

  function parseLine(line, ctx) {
    const toks = tokenize(line).map(classify);
    const res = { persons: [], events: [], ages: [], nums: [] };

    // события: от слова-события до конца строки (или до следующего имени)
    const evMatch = N.modern(line).match(N.EVENT_RX);
    let namesBefore = toks;
    if (evMatch) {
      const idx = N.modern(line).search(N.EVENT_RX);
      const head = N.modern(line).slice(0, idx).trim();
      const tail = N.modern(line).slice(idx).trim();
      res.events = tail.split(/(?=(?:^|\s)(?:умер|умре|бежал|купл[её]н|причисл|переведен|отдан|продан|отпущен|убыл|выбыл|вновь\s+рожд|новорожд|взят|женил|пропущен)[а-я]*)/i)
        .map((s) => N.modern(s).replace(/\s+/g, ' ').trim().replace(/^[,.;]+|[,.;]+$/g, '')).filter(Boolean)
        .map((s) => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase());
      namesBefore = tokenize(head).map(classify);
    }
    const T = namesBefore;
    // числа строки
    for (const t of T) if (t.type === 'num' && t.v > 0 && t.v < 111) res.ages.push(t.v);
    const W = T.filter((t) => t.type !== 'num' && t.type !== 'skip');
    res.nums = res.ages;

    let listing = null;
    let cur = null;           // последний созданный в строке
    let prevType = '';
    let pendingAnchor = null;

    const build = (nameTok, rel) => {
      const nm = nameTok.nm;
      const p = newPerson(ctx, { name: nm.nom, gender: nm.g, rel: rel ? rel.rel : '', relTo: '' });
      p.raw.push(modernWord(nameTok.w));
      if (nm.fuzzy) p.flags.push('имя по созвучию: «' + nameTok.w + '» → ' + nm.nom);
      if (nm.trunc) p.flags.push('имя обрезано: «' + nameTok.w + '» → ' + nm.nom);
      if (rel) {
        if (rel.g && nm.g && rel.g !== nm.g) p.flags.push('пол по имени, а не по связи «' + rel.rel + '»');
        if (!p.gender) p.gender = rel.g;
      }
      return p;
    };

    for (let i = 0; i < W.length; i++) {
      const t = W[i];
      const nxt = W[i + 1];
      if (t.type === 'name') {
        // якорь: имя в косвенном падеже + (отчество/фамилия) + связь
        let j = i + 1, patTok = null;
        while (j < W.length && (W[j].type === 'pat' || W[j].type === 'unk') && j - i <= 3) { if (W[j].type === 'pat' && !patTok) patTok = W[j]; j++; }
        if (W[j] && W[j].type === 'rel' && t.nmGen && (t.nmGen.cs === 'gen' || t.nmGen.cs === 'obl') && (t.nm.cs !== 'nom' || j > i + 1 || i === 0 || !listing)) {
          t.nm = t.nmGen;
          pendingAnchor = findAnchor(ctx, t, patTok) || newPerson(ctx, {
            name: t.nm.nom, gender: t.nm.g, patr: patTok ? N.patr(null, t.nm.g) || patTok.p.base : '', virtual: true,
          });
          if (pendingAnchor.virtual && !pendingAnchor.raw.length) { pendingAnchor.raw.push(modernWord(t.w)); pendingAnchor.flags.push('человек из ссылки, отдельной строки нет'); }
          ctx.fam.anchor = pendingAnchor;
          listing = null;
          i = j - 1;
          continue;
        }
        let p;
        if (listing) {
          p = build(t, listing);
          p.relTo = listing.anchor ? [listing.anchor.name, listing.anchor.patr, listing.anchor.surname].filter(Boolean).join(' ') : '';
          p.anchor = listing.anchor;
        } else if (cur && prevType === 'name' && !cur.patr && !cur.surname && cur.name !== t.nm.nom) {
          // «многоименный»: следующее имя подряд — отдельный ребёнок первого
          p = build(t, { rel: t.nm.g === 'ж' ? 'дочь' : 'сын', g: t.nm.g });
          p.anchor = cur;
          p.relTo = cur.name;
          p.flags.push('многоименная запись: имя отнесено к ребёнку');
        } else {
          p = build(t, null);
          p.head = true;
          ctx.fam.n += 1;
          ctx.fam.members = [];
          ctx.fam.members.push(p);
          p.famNo = ctx.fam.n;
          ctx.fam.anchor = p;
          pendingAnchor = null;
        }
        res.persons.push(p);
        cur = p;
      } else if (t.type === 'pat' && nxt && nxt.type === 'rel') {
        // «Ивановы дети», «Матвеевы сыновья»: якорь — человек с таким отчеством отца
        const father = N.patrToFather(t.p.base);
        let a = null;
        if (father) for (let q = ctx.fam.members.length - 1; q >= 0; q--) if (ctx.fam.members[q].name === father) { a = ctx.fam.members[q]; break; }
        pendingAnchor = a || cur || ctx.fam.anchor;
        ctx.fam.anchor = pendingAnchor;
      } else if (t.type === 'rel') {
        const prevLink = W[i - 1] && W[i - 1].type === 'link' ? W[i - 1] : null;
        let a = pendingAnchor || ctx.fam.anchor;
        if (prevLink && prevLink.who === 'ж') {
          const f = [...ctx.fam.members].reverse().find((m) => m.gender === 'ж');
          if (f) a = f;
        }
        listing = { rel: t.rel.rel, g: t.rel.g, plural: t.rel.plural, anchor: a };
        pendingAnchor = null;
      } else if (t.type === 'link') {
        // «Его»/«Ея» ссылаются на текущий якорь; ничего не создаём
      } else if (t.type === 'status') {
        const target = cur || ctx.fam.anchor;
        if (target) { target.status = target.status ? target.status + '; ' + t.k : t.k; target.raw.push(modernWord(t.w)); }
      } else if (t.type === 'pat') {
        if (cur && !cur.patr) {
          cur.patr = t.p.base + (cur.gender === 'ж' ? 'а' : '');
          cur.raw.push(modernWord(t.w));
        } else if (cur && !cur.surname) { cur.surname = modernWord(t.w); cur.raw.push(modernWord(t.w)); }
      } else if (t.type === 'unk') {
        if (cur && !cur.surname && cur.raw.length) { cur.surname = modernWord(t.w); cur.raw.push(modernWord(t.w)); }
      }
      prevType = t.type;
    }

    // отчества и фамилии от якоря
    for (const p of res.persons) {
      const a = p.anchor;
      if (!a || p.head) continue;
      if (!p.patr) {
        if (['сын', 'дочь', 'дети', 'сыновья', 'дочери'].includes(p.rel) && a.name) p.patr = N.patr(a.name, p.gender);
        else if (['брат', 'братья', 'сестра'].includes(p.rel) && a.patr) p.patr = a.patr.replace(/а$/, '') + (p.gender === 'ж' ? 'а' : '');
        if (p.patr) p.flags.push('отчество достроено от ' + (['брат', 'братья', 'сестра'].includes(p.rel) ? 'брата/сестры' : 'отца'));
      }
      if (!p.surname && a.surname && p.rel !== 'жена') p.surname = a.surname;
    }
    for (const p of res.persons) {
      const parts = [p.name, p.patr, p.surname].filter(Boolean);
      p.full = parts.join(' ');
    }
    return res;
  }

  function isArtifact(line) {
    const k = kline(line);
    if (!k) return true;
    return ARTIFACT.test(k);
  }

  // дата из заголовка страницы: «1850 года Іюля 10 дня», «1811 Года сентября 30 дня»
  function textDate(text) {
    for (const ln of String(text || '').split(/\r?\n/)) {
      const s = N.modern(ln).trim();
      if (s.length < 60 && /(?<!\d)1[5-9]\d\d(?:-го)?\s+года?\s+[а-я]+\s+\d{1,2}\s+дня/i.test(s)) return s;
    }
    for (const ln of String(text || '').split(/\r?\n/)) {
      const m = N.modern(ln).match(/(?<!\d)(1[5-9]\d\d)(?:-го)?\s+года?\s+([а-яА-Я]+)\s+(?:дня|числа)?/i);
      if (m && /январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр|іюн|iюн|iюл|мар|фев|янв|ноя|ок/i.test(m[2])) return ln.trim().slice(0, 80);
    }
    return '';
  }

  const api = { parseLine, isArtifact, kline, textDate, tokenize, classify };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Parser = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
