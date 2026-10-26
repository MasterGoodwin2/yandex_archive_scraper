// Страница приложения: загружает страницы Яндекс Архива в рабочей вкладке Chrome
// (где вы уже вошли в Яндекс), разбирает текст и сохраняет результат в Excel.
(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);

  const CFG = {
    PAGE_URL: (fileId, n) => `https://ya.ru/archive/catalog/${fileId}/${n}`,
    DELAY: [1000, 2500],     // пауза между страницами, мс
    PAGE_TIMEOUT: 45000,     // ожидание загрузки страницы, мс
    MAX_RETRIES: 2,
    RENDER_WAIT: 15000,      // сколько ждать появления распознанного текста, мс
    BIG_RANGE: 300,          // предупреждение о слишком большом диапазоне
    MAX_MISSES: 3,           // подряд страниц без окна текста — остановка
  };
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const NEEDS_USER = /captcha|passport\.yandex|\/auth/i;

  let linksDoc = null;       // разобранный HTML со ссылками
  let items = [];            // ссылки
  let running = false;
  let stopRequested = false;
  let resumeResolver = null;
  let workTabId = null;
  let appTabId = null;
  let fileHandle = null;
  let debugSaved = false;
  let busyPicking = false;

  let settings = { mode: Logic.DEFAULT_MODE, textSel: '', dateSel: '' };

  let people = [];
  let sheetRows = [];
  let linkRows = [];

  // ---------------------------------------------------------------- UI
  function log(msg) {
    const t = new Date().toTimeString().slice(0, 8);
    const el = $('log');
    el.textContent += `[${t}] ${msg}\n`;
    el.scrollTop = el.scrollHeight;
  }
  function setStatus(s) { $('status').textContent = s; }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const rand = ([a, b]) => a + Math.random() * (b - a);
  const mode = () => Logic.getMode(settings.mode);

  function setRunningUi(on) {
    $('startBtn').disabled = on;
    $('linksFile').disabled = on;
    document.querySelectorAll('input[name="mode"]').forEach((r) => { r.disabled = on; });
    $('pickTextBtn').disabled = on;
    $('pickDateBtn').disabled = on;
    $('resetSelBtn').disabled = on;
    $('saveBtn').disabled = !on;
    $('stopBtn').disabled = !on;
  }

  function showSelectors() {
    $('textSelInfo').textContent = settings.textSel || 'автоопределение (справа от окна со сканом)';
    $('dateSelInfo').textContent = settings.dateSel || 'автоопределение (выше панели инструментов окна со сканом)';
  }

  // ---------------------------------------------------------------- настройки
  async function loadSettings() {
    try {
      const got = await chrome.storage.local.get(['mode', 'textSel', 'dateSel']);
      if (got.mode && Logic.MODES[got.mode]) settings.mode = got.mode;
      settings.textSel = got.textSel || '';
      settings.dateSel = got.dateSel || '';
    } catch (e) { /* без сохранения — не критично */ }
  }
  async function saveSettings() {
    try { await chrome.storage.local.set(settings); } catch (e) { /* ignore */ }
  }

  function applyModeUi() {
    $('descConfession').textContent = Logic.MODES.confession.description;
    $('descRevision').textContent = Logic.MODES.revision.description;
    document.querySelectorAll('input[name="mode"]').forEach((r) => { r.checked = r.value === settings.mode; });
  }

  document.querySelectorAll('input[name="mode"]').forEach((r) => {
    r.addEventListener('change', async () => {
      if (!r.checked) return;
      settings.mode = r.value;
      await saveSettings();
      reparseLinks();
      log(`Тип обработки: ${mode().label}.`);
    });
  });

  // ------------------------------------------------------------- ожидание пользователя
  function waitForUser(text, showWorkTab) {
    $('bannerText').textContent = text;
    $('banner').style.display = 'block';
    if (showWorkTab && workTabId !== null) chrome.tabs.update(workTabId, { active: true }).catch(() => {});
    return new Promise((resolve) => {
      resumeResolver = () => {
        $('banner').style.display = 'none';
        resumeResolver = null;
        if (showWorkTab && appTabId !== null && !$('showTab').checked) {
          chrome.tabs.update(appTabId, { active: true }).catch(() => {});
        }
        resolve();
      };
    });
  }
  $('continueBtn').addEventListener('click', () => resumeResolver && resumeResolver());
  $('stopBtn').addEventListener('click', () => {
    stopRequested = true;
    log('Остановка запрошена — дорабатываю текущую страницу и сохраняю результат...');
    if (resumeResolver) resumeResolver();
  });
  $('saveBtn').addEventListener('click', async () => { await saveResult(true); });
  window.addEventListener('beforeunload', (e) => {
    if (running) { e.preventDefault(); e.returnValue = ''; }
  });

  // ------------------------------------------------------ список ссылок
  function reparseLinks() {
    items = linksDoc ? Logic.parseLinks(linksDoc, settings.mode) : [];
    if (!linksDoc) return;
    const ok = items.filter((i) => i.fileId && i.pFrom !== null).length;
    $('linksInfo').textContent = items.length
      ? `Найдено ссылок Яндекс Архива: ${items.length} (с номером документа и диапазоном: ${ok})`
      : 'В файле не найдено ссылок Яндекс Архива.';
  }

  $('linksFile').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    linksDoc = null;
    items = [];
    if (!f) { $('linksInfo').textContent = 'Файл не выбран'; return; }
    try {
      linksDoc = new DOMParser().parseFromString(await f.text(), 'text/html');
    } catch (err) {
      $('linksInfo').textContent = 'Не удалось прочитать файл: ' + err.message;
      return;
    }
    reparseLinks();
  });

  // ------------------------------------------------------ рабочая вкладка
  async function getWorkTab() {
    if (workTabId !== null) {
      try { await chrome.tabs.get(workTabId); return workTabId; } catch (e) { workTabId = null; }
    }
    const t = await chrome.tabs.create({ url: 'about:blank', active: $('showTab').checked });
    workTabId = t.id;
    return workTabId;
  }

  function navigate(tabId, url, timeoutMs) {
    return new Promise((resolve, reject) => {
      let finished = false;
      let sawLoading = false;
      const done = (err) => {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        chrome.tabs.onUpdated.removeListener(onUpd);
        err ? reject(err) : resolve();
      };
      const timer = setTimeout(() => done(new Error('таймаут загрузки страницы')), timeoutMs);
      const onUpd = (id, info) => {
        if (id !== tabId) return;
        if (info.status === 'loading') sawLoading = true;
        if (info.status === 'complete' && sawLoading) done();
      };
      chrome.tabs.onUpdated.addListener(onUpd);
      chrome.tabs.update(tabId, { url }).catch((e) => done(e));
    });
  }

  async function fetchSheet(fileId, n) {
    const url = CFG.PAGE_URL(fileId, n);
    let lastErr = '';
    for (let attempt = 0; attempt <= CFG.MAX_RETRIES; attempt++) {
      if (stopRequested) return { url, text: '', date: '', found: false, err: 'остановлено пользователем' };
      try {
        const tabId = await getWorkTab();
        await navigate(tabId, url, CFG.PAGE_TIMEOUT);
        let tab = await chrome.tabs.get(tabId);
        let guard = 0;
        while (NEEDS_USER.test(tab.url || '') && guard++ < 3 && !stopRequested) {
          const captcha = /captcha/i.test(tab.url);
          await waitForUser(
            captcha
              ? 'Яндекс просит капчу. Решите её в рабочей вкладке и нажмите «Продолжить».'
              : 'Сессия Яндекса сбросилась. Войдите заново в рабочей вкладке и нажмите «Продолжить».',
            true
          );
          await navigate(tabId, url, CFG.PAGE_TIMEOUT);
          tab = await chrome.tabs.get(tabId);
        }
        if ((tab.url || '').startsWith('chrome-error')) throw new Error('страница не загрузилась (ошибка сети)');

        const wantHtml = $('debugHtml').checked && !debugSaved;
        const [inj] = await chrome.scripting.executeScript({
          target: { tabId },
          func: PageScripts.pageExtractor,
          args: [{ textSel: settings.textSel, dateSel: settings.dateSel, maxWait: CFG.RENDER_WAIT, wantHtml }],
        });
        const res = (inj && inj.result) || { found: false, text: '', date: '' };
        if (wantHtml && res.html) {
          downloadBlob(new Blob([res.html], { type: 'text/html;charset=utf-8' }), 'first_page.html');
          debugSaved = true;
          log('Сохранён отладочный файл first_page.html (папка «Загрузки»).');
        }
        if (n !== undefined && res.found) {
          log(`  стр. ${n}: окно текста — ${res.textHow}, знаков: ${res.text.length}` +
              (res.date ? `; дата: «${res.date}» (${res.dateHow})` : '; дата не найдена'));
        }
        return {
          url, text: res.text || '', date: res.date || '', found: !!res.found,
          err: res.found ? '' : 'окно распознанного текста не найдено',
        };
      } catch (e) {
        lastErr = e && e.message ? e.message : String(e);
        log(`  страница ${n}: ошибка «${lastErr}», попытка ${attempt + 1}/${CFG.MAX_RETRIES + 1}`);
        await sleep(2000);
      }
    }
    return { url, text: '', date: '', found: false, err: lastErr || 'ошибка' };
  }

  async function ensureLogin() {
    const tabId = await getWorkTab();
    log('Проверяю вход в Яндекс...');
    try {
      await navigate(tabId, 'https://ya.ru/archive', 60000);
    } catch (e) {
      log('Страница ya.ru/archive открылась с ошибкой: ' + e.message);
    }
    const tab = await chrome.tabs.get(tabId);
    const onArchive = /^https:\/\/(ya\.ru|yandex\.ru)\/archive/.test(tab.url || '');
    if (!onArchive || NEEDS_USER.test(tab.url || '')) {
      await waitForUser(
        'Войдите в Яндекс в рабочей вкладке (она сейчас открыта; убедитесь, что там открыт ya.ru/archive), ' +
        'затем вернитесь на эту вкладку и нажмите «Продолжить».',
        true
      );
      if (stopRequested) return;
    }
    // прогреваем сессию на втором домене (yandex.ru) — на случай ссылок с ним
    try {
      await navigate(tabId, 'https://yandex.ru/archive', 30000);
      await navigate(tabId, 'https://ya.ru/archive', 30000);
    } catch (e) { /* не критично */ }
    log('Вход проверен.');
  }

  // ------------------------------------------------------ ручная калибровка
  async function calibrate(kind) {
    if (running || busyPicking) return;
    const it = items.find((i) => i.fileId && i.pFrom !== null);
    if (!it) { setStatus('Сначала выберите HTML-файл со ссылками: расширение откроет первую из них.'); return; }
    busyPicking = true;
    const label = kind === 'text' ? 'окно расшифрованного текста (справа от скана)' : 'строку с датой документа (над панелью инструментов скана)';
    try {
      if (appTabId === null) { const cur = await chrome.tabs.getCurrent(); appTabId = cur ? cur.id : null; }
      const tabId = await getWorkTab();
      await chrome.tabs.update(tabId, { active: true });
      log(`Калибровка: открываю страницу ${it.pFrom} документа ${it.fileId}...`);
      try { await navigate(tabId, CFG.PAGE_URL(it.fileId, it.pFrom), CFG.PAGE_TIMEOUT); }
      catch (e) { log('Страница открылась с ошибкой: ' + e.message); }
      const tab = await chrome.tabs.get(tabId);
      if (NEEDS_USER.test(tab.url || '')) {
        log('Нужен вход или капча: выполните их в открытой вкладке и нажмите кнопку калибровки ещё раз.');
        return;
      }
      await sleep(3500);   // дать странице дорисоваться
      log(`В рабочей вкладке наведите мышь на ${label} и кликните. Esc — отмена.`);
      const [inj] = await chrome.scripting.executeScript({
        target: { tabId }, func: PageScripts.pickerRun, args: [label],
      });
      const res = inj && inj.result;
      if (!res) { log('Выбор отменён.'); return; }
      if (kind === 'text') settings.textSel = res.selector; else settings.dateSel = res.selector;
      await saveSettings();
      showSelectors();
      log(`Выбрано: ${res.selector} (совпадений: ${res.count}).`);
      log(`Начало содержимого: «${res.sample.replace(/\s+/g, ' ').slice(0, 160)}»`);
      if (res.count !== 1) log('Внимание: селектор подходит не к одному элементу — проверьте результат.');
    } catch (e) {
      log('Калибровка не удалась: ' + (e && e.message ? e.message : e));
    } finally {
      busyPicking = false;
      if (appTabId !== null) chrome.tabs.update(appTabId, { active: true }).catch(() => {});
    }
  }
  $('pickTextBtn').addEventListener('click', () => calibrate('text'));
  $('pickDateBtn').addEventListener('click', () => calibrate('date'));
  $('resetSelBtn').addEventListener('click', async () => {
    settings.textSel = '';
    settings.dateSel = '';
    await saveSettings();
    showSelectors();
    log('Сброшено на автоопределение.');
  });

  // ------------------------------------------------------------- Excel
  const trunc = (v) => (typeof v === 'string' ? v.slice(0, 32000) : v === null || v === undefined ? '' : v);

  function widthsFor(rows) {
    const cols = Math.max(...rows.map((r) => r.length));
    const out = [];
    for (let c = 0; c < cols; c++) {
      let w = 8;
      rows.slice(0, 200).forEach((r) => {
        const v = r[c];
        if (v !== '' && v !== null && v !== undefined) w = Math.max(w, String(v).length);
      });
      out.push(Math.min(Math.max(w + 2, 10), 70));
    }
    return out;
  }

  function buildSheets() {
    const revision = mode().id === 'revision';
    const peopleRows = [['Год', 'Тип документа', 'Дата документа (со страницы)', 'Населённый пункт', 'Двор №',
      'Страница', 'Строка на странице', 'Числа-возрасты (черновик)',
      revision ? 'Возраст: 1-е число' : '', revision ? 'Возраст: 2-е число' : '',
      'Распознанная строка', 'Ссылка на страницу']];
    people.forEach((r) => peopleRows.push([r.year, r.docType, r.docDate, r.settlement, r.household,
      r.page, r.lineNo, r.ages, r.age1, r.age2, r.text, r.url].map(trunc)));

    const pageRows = [['Год', 'Тип документа', 'Страница', 'Дата документа (со страницы)',
      'Найденные НП округи', 'Ссылка', 'Распознанный текст страницы']];
    sheetRows.forEach((r) => pageRows.push([r.year, r.docType, r.page, r.docDate, r.found, r.url, r.text].map(trunc)));

    const linkOut = [['Год', 'Тип документа', 'Исходная ссылка', 'Диапазон из ссылки', 'Просмотрено страниц',
      'Найдено НП (шт.)', 'Какие НП', 'Решение', 'Людей собрано', 'Дата документа (1-я страница)', 'Примечание']];
    linkRows.forEach((r) => linkOut.push(r.map(trunc)));

    return [
      { name: 'Люди', rows: peopleRows, widths: widthsFor(peopleRows) },
      { name: 'Страницы', rows: pageRows, widths: widthsFor(pageRows) },
      { name: 'Ссылки', rows: linkOut, widths: widthsFor(linkOut) },
    ];
  }

  function downloadBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  }

  // download=true — если файл не выбран заранее, скачать копию в «Загрузки»
  async function saveResult(download) {
    const bytes = await XlsxWriter.build(buildSheets());
    if (fileHandle) {
      try {
        const w = await fileHandle.createWritable();
        await w.write(bytes);
        await w.close();
        log(`Сохранено: людей ${people.length}, страниц ${sheetRows.length}.`);
        return true;
      } catch (e) {
        log('Не удалось записать в выбранный файл (' + e.message + '). Скачиваю копию в «Загрузки».');
        download = true;
      }
    }
    if (download) {
      downloadBlob(new Blob([bytes], { type: XLSX_MIME }), 'result.xlsx');
      log('Файл result.xlsx скачан в папку «Загрузки».');
    }
    return true;
  }

  // ------------------------------------------------------------ основной цикл
  async function run() {
    running = true;
    stopRequested = false;
    debugSaved = false;
    people = [];
    sheetRows = [];
    linkRows = [];
    const m = mode();
    const pageCache = new Map();     // `${fileId}:${page}` -> {url,text,date,found,err}
    const seenPeople = new Set();
    let misses = 0;
    setRunningUi(true);
    log(`Тип обработки: ${m.label}. Окно текста: ${settings.textSel || 'авто'}; дата: ${settings.dateSel || 'авто'}.`);

    try {
      const cur = await chrome.tabs.getCurrent();
      appTabId = cur ? cur.id : null;
      await ensureLogin();

      for (let idx = 0; idx < items.length && !stopRequested; idx++) {
        const it = items[idx];
        const tag = `[${idx + 1}/${items.length}] ${it.year ?? '—'}`;
        const rng = it.pFrom !== null ? `${it.pFrom}–${it.pTo}` : '—';

        if (!it.fileId || it.pFrom === null) {
          linkRows.push([it.year, m.label, it.url, rng, 0, 0, '', 'пропущена: нет fileId или диапазона', 0, '', it.note]);
          log(`${tag}: пропущена (нет fileId/диапазона)`);
          continue;
        }

        const lo = it.pFrom;
        const hi = it.pTo;
        if (hi - lo + 1 > CFG.BIG_RANGE) log(`${tag}: внимание, диапазон очень большой (${hi - lo + 1} страниц)`);

        // 1. загрузка страниц диапазона из ссылки
        const texts = new Map();
        for (let n = lo; n <= hi && !stopRequested; n++) {
          const key = `${it.fileId}:${n}`;
          if (!pageCache.has(key)) {
            setStatus(`${tag}: страница ${n} (диапазон ${lo}–${hi})`);
            log(`${tag}: загружаю страницу ${n} (диапазон ${lo}–${hi})...`);
            const r = await fetchSheet(it.fileId, n);
            pageCache.set(key, r);
            if (!r.found && !stopRequested) {
              if (++misses >= CFG.MAX_MISSES) {
                stopRequested = true;
                log(`Окно распознанного текста не найдено на ${CFG.MAX_MISSES} страницах подряд — останавливаюсь, ` +
                    'чтобы не собирать лишнее. Нажмите «Указать окно текста вручную» и запустите снова.');
              }
            } else if (r.found) {
              misses = 0;
            }
            if (n < hi) await sleep(rand(CFG.DELAY));
          }
          texts.set(n, pageCache.get(key));
        }

        // 2. сбор людей
        const foundAll = new Set();
        const perPage = new Map();
        for (const [n, r] of texts) {
          const f = Logic.findSettlements(r.text, m.id);
          perPage.set(n, f);
          f.forEach((x) => foundAll.add(x));
        }

        const state = { current: null, other: null, started: true, household: null };
        let collected = 0;
        let firstDate = '';
        for (const n of [...texts.keys()].sort((a, b) => a - b)) {
          const r = texts.get(n);
          if (!firstDate && r.date) firstDate = r.date;
          sheetRows.push({
            year: it.year, docType: m.label, page: n, docDate: r.date,
            found: [...perPage.get(n)].sort().join('; '),
            url: r.url, text: r.err ? `ОШИБКА: ${r.err}` : r.text,
          });
          for (const p of Logic.extractPeople(r.text, state, true, m.id)) {
            const key = `${it.fileId}|${n}|${p.lineNo}|${p.text}`;
            if (seenPeople.has(key)) continue;
            seenPeople.add(key);
            people.push({ ...p, year: it.year, docType: m.label, docDate: r.date, page: n, url: r.url });
            collected++;
          }
        }

        const total = hi - lo + 1;
        const verdict = texts.size < total
          ? `ПРЕРВАНА (страниц ${texts.size} из ${total})`
          : 'ОБРАБОТАНА';
        linkRows.push([it.year, m.label, it.url, rng, texts.size, foundAll.size,
          [...foundAll].sort().join('; '), verdict, collected, firstDate, it.note]);
        log(`${tag}: диапазон ${rng} — НП ${foundAll.size}, людей: ${collected}`);
        await saveResult(false);
      }
    } catch (e) {
      log('ОШИБКА: ' + (e && e.stack ? e.stack : e));
    } finally {
      try { await saveResult(true); } catch (e) { log('Не удалось сохранить: ' + e.message); }
      if (workTabId !== null) { chrome.tabs.remove(workTabId).catch(() => {}); workTabId = null; }
      running = false;
      setRunningUi(false);
      setStatus(`Готово. Людей: ${people.length}.`);
      log(`Готово. Людей: ${people.length}.`);
    }
  }

  $('startBtn').addEventListener('click', async () => {
    if (running || busyPicking) return;
    if (!items.length) { setStatus('Сначала выберите HTML-файл со ссылками.'); return; }
    // окно выбора файла должно открываться прямо по клику
    fileHandle = null;
    if (window.showSaveFilePicker) {
      try {
        fileHandle = await window.showSaveFilePicker({
          suggestedName: 'result.xlsx',
          types: [{ description: 'Excel', accept: { [XLSX_MIME]: ['.xlsx'] } }],
        });
      } catch (e) {
        if (e && e.name === 'AbortError') { setStatus('Место сохранения не выбрано.'); return; }
        log('Окно выбора файла недоступно (' + (e && e.message) + '). Результат будет скачан в «Загрузки».');
      }
    }
    await run();
  });

  // ------------------------------------------------------------- старт
  (async () => {
    await loadSettings();
    applyModeUi();
    showSelectors();
    try { const cur = await chrome.tabs.getCurrent(); appTabId = cur ? cur.id : null; } catch (e) { /* ignore */ }
  })();
})();
