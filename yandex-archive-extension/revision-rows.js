// Ревизские сказки: привязка соседних столбцов к Основному блоку («Крестьяне» / «Крестьянки»).
// Работает по готовым данным страницы Яндекс Архива (__NEXT_DATA__ -> currentNode):
//   joinedTextBlocks[i] = {id, points[4]}  (координаты на скане, пиксели)
//   textBlocks[i]       = {id, text}       (тот же блок; списки идут в одном порядке)
// Текст НЕ разбирается и НЕ меняется. Это перенос на JS скрипта навыка yaarchive-revision-rows
// (skill/yaarchive-revision-rows/scripts/extract_rows.py); результаты обеих версий должны совпадать.
(function (root) {
  'use strict';

  const DEFAULTS = { splitX: null, famTol: 0.025, nearTol: 0.01, colTol: 0.05, tall: 1.9, rowOverlap: 0.5 };
  const CH = { 'ѣ': 'е', 'і': 'и', 'i': 'и', 'ѳ': 'ф', 'ѵ': 'и', 'ё': 'е' };
  const ROLES = ['age_prev', 'age_mid', 'departed_when', 'absent_since', 'age_now'];

  const norm = (t) => String(t || '').toLowerCase().replace(/[ѣіiѳѵё]/g, (c) => CH[c])
    .replace(/[^\p{L}\p{N}_\s/]/gu, ' ').replace(/\s+/g, ' ').trim();

  function headerKind(text) {
    const n = norm(text);
    if (n === 'крестьяне') return ['main', 'M'];
    if (n === 'крестьянки') return ['main', 'F'];
    if (n === 'семьи') return ['family', null];
    if (n === 'лета') return ['age', null];
    if (n.startsWith('когда имен')) return ['departed_when', null];
    if (/котор[\p{L}\p{N}_]* времени/u.test(n)) return ['absent_since', null];
    return null;
  }

  function mkBlk(i, j, t) {
    const pts = j.points || [];
    if (!pts.length) throw new Error('пустой polygon у блока ' + j.id);
    const xs = pts.map((p) => p.x), ys = pts.map((p) => p.y);
    const x0 = Math.min(...xs), y0 = Math.min(...ys), x1 = Math.max(...xs), y1 = Math.max(...ys);
    return { i, id: j.id, tid: t.id, text: String(t.text || '').trim(), x0, y0, x1, y1,
      w: x1 - x0, h: y1 - y0, cx: (x0 + x1) / 2, cy: (y0 + y1) / 2 };
  }

  const cellJson = (bs) => ({
    text: bs.map((b) => b.text).join(' '),
    ids: bs.map((b) => b.id),
    text_block_ids: bs.map((b) => b.tid),
    bbox: [Math.min(...bs.map((b) => b.x0)), Math.min(...bs.map((b) => b.y0)),
      Math.max(...bs.map((b) => b.x1)), Math.max(...bs.map((b) => b.y1))],
  });

  const toInt = (t) => { const d = String(t || '').replace(/\D/g, ''); return d ? parseInt(d, 10) : null; };
  const median = (arr) => {
    if (!arr.length) return 0;
    const s = arr.slice().sort((x, y) => x - y), m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  };
  const push = (map, k, v) => { if (!map.has(k)) map.set(k, []); map.get(k).push(v); };

  function process(node, opts) {
    const a = Object.assign({}, DEFAULTS, opts || {});
    const size = node.originalImageSize || {};
    const W = size.width, H = size.height;
    const jb = node.joinedTextBlocks || [], tb = node.textBlocks || [];
    const rep = { page: node.sheetPageNumber ?? null, node_id: node.id ?? null, image: [W, H],
      flags: [], blocks: [], orphans: [], totals: {}, totals_check: {}, titles: [] };
    if (!(W && H) || !jb.length || jb.length !== tb.length) {
      rep.flags.push(`count_mismatch:joined=${jb.length},text=${tb.length}`);
      return rep;
    }
    const blks = jb.map((j, i) => mkBlk(i, j, tb[i]));

    // 1. заголовки столбцов (по тексту) задают x-границы столбцов
    const heads = [];
    for (const b of blks) {
      if (b.cy < 0.4 * H) { const k = headerKind(b.text); if (k) heads.push([k, b]); }
    }
    const mains = heads.filter((h) => h[0][0] === 'main').sort((p, q) => p[1].cx - q[1].cx);
    if (!mains.length) { rep.flags.push('no_main_header'); return rep; }
    const hdrBottom = Math.max(...heads.map((h) => h[1].y1));
    const headIds = new Set(heads.map((h) => h[1].id));
    rep.titles = blks.filter((b) => !headIds.has(b.id) && b.cy <= hdrBottom)
      .sort((p, q) => p.cy - q.cy || p.cx - q.cx).map((b) => b.text);
    const splitX = a.splitX || W / 2;
    const groupOf = (b) => (mains.length === 1 ? mains[0][0][1]
      : (b.cx < splitX ? mains[0][0][1] : mains[mains.length - 1][0][1]));

    const cols = [], byGroup = new Map();
    for (const [k, b] of heads) push(byGroup, groupOf(b), [k[0], b]);
    for (const [g, items] of byGroup) {
      items.sort((p, q) => p[1].cx - q[1].cx);
      const ages = items.filter((t) => t[0] === 'age').map((t) => t[1]);
      for (const [k, b] of items) {
        let role = k;
        if (k === 'age') role = b === ages[ages.length - 1] ? 'age_now' : (b === ages[0] ? 'age_prev' : 'age_mid');
        cols.push({ role, x0: b.x0, x1: b.x1, group: g });
      }
    }
    const body = blks.filter((b) => !headIds.has(b.id) && b.cy > hdrBottom);

    // 2. итоги внизу страницы
    const footIds = new Set();
    for (const f of body.filter((b) => norm(b.text).startsWith('итого'))) {
      footIds.add(f.id);
      const nf = norm(f.text);
      const sex = nf.includes('мужеск') ? 'M' : (nf.includes('женск') ? 'F' : null);
      const right = body.filter((b) => !footIds.has(b.id) && b.x0 >= f.x1 - 0.01 * W
        && Math.abs(b.cy - f.cy) < 0.5 * Math.max(b.h, f.h)).sort((p, q) => p.x0 - q.x0);
      if (sex && right.length) { footIds.add(right[0].id); rep.totals[sex] = toInt(right[0].text); }
    }
    const rest = body.filter((b) => !footIds.has(b.id));

    // 3. блок -> столбец по x
    const colOf = (b) => {
      let best = null;
      for (const c of cols) {
        const dx = Math.max(c.x0 - b.cx, 0, b.cx - c.x1);
        if (best === null || dx < best[0]) best = [dx, c];
      }
      return best;
    };
    const mainb = new Map(), fam = new Map(), cells = [];
    for (const b of rest) {
      const [dx, c] = colOf(b);
      if (dx > a.colTol * W) { rep.orphans.push({ text: b.text, id: b.id, bbox: [b.x0, b.y0, b.x1, b.y1], why: 'no_column' }); continue; }
      if (c.role === 'main') push(mainb, c.group, b);
      else if (c.role === 'family') push(fam, c.group, b);
      else cells.push([b, c]);
    }
    for (const v of mainb.values()) v.sort((p, q) => p.y0 - q.y0);
    for (const v of fam.values()) v.sort((p, q) => p.cy - q.cy);

    // 4. ячейка -> Основной блок своей половины (самый низкий из содержащих по y)
    const assigned = new Map();
    for (const v of mainb.values()) for (const m of v) assigned.set(m.id, []);
    const nested = new Set();
    for (const [b, c] of cells) {
      const cand = mainb.get(c.group) || [];
      const inside = cand.filter((m) => m.y0 <= b.cy && b.cy <= m.y1);
      let m, flag = null;
      if (inside.length) {
        m = inside.reduce((p, q) => (q.h < p.h ? q : p));
        if (inside.length > 1) nested.add(m.id);
      } else {
        const dist = (mm) => Math.min(Math.abs(b.cy - mm.y0), Math.abs(b.cy - mm.y1));
        const near = cand.slice().sort((p, q) => dist(p) - dist(q));
        if (near.length && dist(near[0]) <= a.nearTol * H) { m = near[0]; flag = 'y_near'; }
        else { rep.orphans.push({ text: b.text, id: b.id, bbox: [b.x0, b.y0, b.x1, b.y1], why: 'no_main_block', role: c.role }); continue; }
      }
      const ov = b.w ? (Math.min(b.x1, c.x1) - Math.max(b.x0, c.x0)) / b.w : 1;
      assigned.get(m.id).push({ b, c, fls: [flag, ov < 0.5 ? 'weak_col' : null].filter(Boolean) });
    }

    const med = median(cells.map((t) => t[0].h));
    const usedFam = new Set();

    // 5. блоки и строки
    const groups = [...mainb.keys()].sort((p, q) => mainb.get(p)[0].cx - mainb.get(q)[0].cx);
    for (const g of groups) {
      mainb.get(g).forEach((m, idx) => {
        const bl = { block_no: idx + 1, sex: g, block_id: m.id, text_block_id: m.tid,
          bbox: [m.x0, m.y0, m.x1, m.y1], names_text: m.text, flags: [], rows: [] };
        if (nested.has(m.id)) bl.flags.push('nested_block');

        const famList = fam.get(g) || [];
        const fl = famList.filter((f) => f.cy <= m.y0 + a.famTol * H);
        if (fl.length) {
          const f = fl[fl.length - 1];
          usedFam.add(f.id);
          bl.family = f.text; bl.family_no = toInt(f.text); bl.family_cell_id = f.id;
        } else {
          bl.family = null; bl.family_no = null; bl.flags.push('family_missing');
        }
        if (famList.some((f) => m.y0 + a.famTol * H < f.cy && f.cy <= m.y1)) bl.flags.push('family_changes_inside_block');

        const items = assigned.get(m.id);
        const normal = items.filter((t) => !(t.b.h > a.tall * med)).sort((p, q) => p.b.cy - q.b.cy);
        const tall = items.filter((t) => t.b.h > a.tall * med);
        const rows = [];
        for (const { b, c, fls } of normal) {
          const last = rows.length ? rows[rows.length - 1] : null;
          let r = null;
          if (last) {
            const ov = Math.min(b.y1, last.y1) - Math.max(b.y0, last.y0);
            if (ov > 0 && ov / Math.min(b.h, last.y1 - last.y0) >= a.rowOverlap) r = last;
          }
          if (!r) { r = { y0: b.y0, y1: b.y1, cells: new Map(), flags: [] }; rows.push(r); }
          r.y0 = Math.min(r.y0, b.y0); r.y1 = Math.max(r.y1, b.y1);
          push(r.cells, c.role, b);
          for (const x of fls) r.flags.push(`${x}:${c.role}`);
        }
        for (const { b, c } of tall) {            // высокая ячейка («новорожденные.») покрывает несколько строк
          let hit = rows.filter((r) => (Math.min(b.y1, r.y1) - Math.max(b.y0, r.y0)) / Math.max(r.y1 - r.y0, 1) >= 0.4);
          if (!hit.length) { const nr = { y0: b.y0, y1: b.y1, cells: new Map(), flags: [] }; rows.push(nr); hit = [nr]; }
          for (const r of hit) {
            if (r.cells.has(c.role)) r.flags.push(`span_conflict:${c.role}`);
            else { r.cells.set(c.role, [b]); r.flags.push(`span:${c.role}`); }
          }
        }
        rows.sort((p, q) => p.y0 - q.y0);
        rows.forEach((r, k) => {
          const out = { row_no: k + 1, y: [r.y0, r.y1], flags: r.flags.slice(), cells: {} };
          for (const [role, bs] of r.cells) {
            out.cells[role] = cellJson(bs);
            if (bs.length > 1) out.flags.push(`multi_cell:${role}`);
          }
          bl.rows.push(out);
        });
        bl.rows_in_block = rows.length;
        rep.blocks.push(bl);
      });
    }
    for (const fs of fam.values()) for (const f of fs) if (!usedFam.has(f.id)) rep.flags.push(`family_cell_unused:${f.text}`);

    // 6. проверка по итогам: строки без «Когда именно» (не выбывшие) == «Итого налицо»
    for (const sex of Object.keys(rep.totals)) {
      const total = rep.totals[sex];
      const all = rep.blocks.filter((bl) => bl.sex === sex).flatMap((bl) => bl.rows);
      const present = all.filter((r) => !r.cells.departed_when);
      rep.totals_check[sex] = { expected: total, rows_total: all.length, rows_present: present.length, ok: total === present.length };
      if (total !== present.length) rep.flags.push(`totals_mismatch:${sex}:expected=${total},rows=${present.length}`);
    }
    if (rep.orphans.length) rep.flags.push(`orphans:${rep.orphans.length}`);
    return rep;
  }

  const api = { process, headerKind, norm, ROLES, DEFAULTS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RevisionRows = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
