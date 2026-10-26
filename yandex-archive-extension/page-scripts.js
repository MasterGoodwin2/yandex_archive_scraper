// Функции, которые выполняются ВНУТРИ страницы Яндекс Архива (через chrome.scripting.executeScript).
// Каждая функция должна быть самодостаточной: внешние переменные ей недоступны.
(function (root) {

  // ----------------------------------------------------------------------------
  // Извлечение: (1) окно распознанного (расшифрованного) текста — справа от окна
  // со сканом; (2) строка даты документа — выше панели инструментов окна со сканом.
  // opts: {textSel, dateSel, maxWait, wantHtml}
  //   textSel/dateSel — CSS-селекторы, заданные вручную (калибровка); пусто — автоопределение.
  // ----------------------------------------------------------------------------
  async function pageExtractor(opts) {
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

  // ----------------------------------------------------------------------------
  // Выбор элемента мышью (калибровка). Возвращает Promise с {selector,count,sample} или null.
  // Мышь — подсветка, ↑ — выбрать родителя, ↓ — вернуться к вложенному, клик — выбрать, Esc — отмена.
  // ----------------------------------------------------------------------------
  function pickerRun(label) {
    return new Promise((resolve) => {
      const Z = 2147483647;
      const box = document.createElement('div');
      box.style.cssText = `position:fixed;pointer-events:none;z-index:${Z};border:2px solid #2563eb;` +
        'background:rgba(37,99,235,.18);transition:all .05s;display:none;';
      const bar = document.createElement('div');
      bar.style.cssText = `position:fixed;top:8px;left:50%;transform:translateX(-50%);z-index:${Z};` +
        'background:#111;color:#fff;font:14px/1.4 sans-serif;padding:8px 14px;border-radius:8px;' +
        'pointer-events:none;box-shadow:0 2px 10px rgba(0,0,0,.4);max-width:90vw;text-align:center;';
      bar.textContent = `Наведите на «${label}» и кликните. ↑ — выбрать контейнер выше, ↓ — вернуться, Esc — отмена`;
      document.documentElement.append(box, bar);

      let cur = null;
      let stack = [];
      const ours = (e) => e === box || e === bar;

      const draw = () => {
        if (!cur) { box.style.display = 'none'; return; }
        const r = cur.getBoundingClientRect();
        box.style.display = 'block';
        box.style.left = r.left + 'px';
        box.style.top = r.top + 'px';
        box.style.width = r.width + 'px';
        box.style.height = r.height + 'px';
      };

      const esc = (s) => (window.CSS && CSS.escape ? CSS.escape(s) : s.replace(/[^\w-]/g, '\\$&'));
      const unique = (sel) => { try { return document.querySelectorAll(sel).length === 1; } catch (e) { return false; } };
      const part = (e) => {
        let p = e.tagName.toLowerCase();
        const cls = [...e.classList].filter((c) => !/^\d/.test(c)).slice(0, 3);
        if (cls.length) p += '.' + cls.map(esc).join('.');
        return p;
      };
      const cssPath = (el) => {
        if (el.id && unique('#' + esc(el.id))) return '#' + esc(el.id);
        const path = [];
        for (let e = el; e && e !== document.body; e = e.parentElement) {
          path.unshift(part(e));
          const s = path.join(' > ');
          if (unique(s)) return s;
        }
        // запасной вариант: полный путь с номерами потомков
        const full = [];
        for (let e = el; e && e !== document.body; e = e.parentElement) {
          const idx = [...e.parentElement.children].indexOf(e) + 1;
          full.unshift(`${e.tagName.toLowerCase()}:nth-child(${idx})`);
        }
        return 'body > ' + full.join(' > ');
      };

      const cleanup = () => {
        document.removeEventListener('mousemove', onMove, true);
        document.removeEventListener('click', onClick, true);
        document.removeEventListener('keydown', onKey, true);
        box.remove();
        bar.remove();
      };
      const onMove = (ev) => {
        const e = document.elementFromPoint(ev.clientX, ev.clientY);
        if (!e || ours(e)) return;
        cur = e;
        stack = [];
        draw();
      };
      const onClick = (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        if (!cur) return;
        const selector = cssPath(cur);
        const out = {
          selector,
          count: (() => { try { return document.querySelectorAll(selector).length; } catch (e) { return 0; } })(),
          sample: (cur.innerText || cur.textContent || '').trim().slice(0, 300),
        };
        cleanup();
        resolve(out);
      };
      const onKey = (ev) => {
        if (ev.key === 'Escape') { ev.preventDefault(); cleanup(); resolve(null); }
        else if (ev.key === 'ArrowUp' && cur && cur.parentElement && cur.parentElement !== document.documentElement) {
          ev.preventDefault(); stack.push(cur); cur = cur.parentElement; draw();
        } else if (ev.key === 'ArrowDown' && stack.length) {
          ev.preventDefault(); cur = stack.pop(); draw();
        }
      };
      document.addEventListener('mousemove', onMove, true);
      document.addEventListener('click', onClick, true);
      document.addEventListener('keydown', onKey, true);
      setTimeout(() => { cleanup(); resolve(null); }, 5 * 60 * 1000);
    });
  }

  const api = { pageExtractor, pickerRun };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PageScripts = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
