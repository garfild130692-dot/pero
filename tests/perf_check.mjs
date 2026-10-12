// Замеры скорости на больших документах (этап 17): открытие, первое разбиение, набор, сохранение.
//   1) npx vite;  2) $env:VITE_DEV_SERVER_URL='http://127.0.0.1:5183'; npx electron . --remote-debugging-port=9223
//   3) node tests/perf_check.mjs [файл.docx]   (по умолчанию tests/fixtures/big.docx от tests/make_big_fixture.py)
// Время «набора» — работа главного потока на одно нажатие: транзакция ProseMirror (синхронно) + замер разбиения
// (window.__peroMeasureMs, его пишет pagination.js в режиме разработки) + прочие плагины после разбиения.
const file = process.argv[2] || 'tests/fixtures/big.docx';
const list = await (await fetch('http://127.0.0.1:9223/json')).json();
const ws = new WebSocket(list.find((p) => p.type === 'page').webSocketDebuggerUrl);
await new Promise((r) => ws.addEventListener('open', r));
let seq = 0;
const send = (expression) => new Promise((res) => {
  const id = ++seq;
  ws.addEventListener('message', function on(e) {
    const m = JSON.parse(e.data);
    if (m.id !== id) return;
    ws.removeEventListener('message', on);
    res(m.result && m.result.result ? m.result.result.value : JSON.stringify({ error: JSON.stringify(m).slice(0, 300) }));
  });
  ws.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, awaitPromise: true, returnByValue: true } }));
});
const out = await send(`(async () => { try {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  document.querySelector('.start-tile')?.click(); await sleep(800);
  const ed = document.querySelector('.ProseMirror').editor;
  const view = ed.view;
  // ожидание с пределом: не дождались — ошибка с названием этапа (а не вечное ожидание)
  const guard = (p, label, ms = 15000) => Promise.race([p, sleep(ms).then(() => { throw new Error('не дождались: ' + label); })]);
  const paginated = () => new Promise((r) => view.dom.addEventListener('pero-paginated', r, { once: true }));
  // ждать, пока разбиение не утихнет (самопроверка полос, раскладка рисунков и т. п.)
  const settle = async () => { for (;;) { const p = paginated(); const t = await Promise.race([p.then(() => 1), sleep(700).then(() => 0)]); if (!t) return; } };
  const { importDocx } = await import('/src/docx/import.js');
  const { exportDocx } = await import('/src/docx/export.js');
  const { loadDocument } = await import('/src/editor/extensions.js');
  // плагин разбиения — из состояния редактора (после горячей замены модулей import дал бы другой экземпляр)
  const pagin = () => view.state.plugins.find((p) => p.key.startsWith('pagination$')).getState(view.state);
  const bytes = await (await fetch('/${file}')).arrayBuffer();
  const r = {};
  // эталонная нагрузка (одна и та же работа): насколько машина сейчас занята другими программами
  { const t0 = performance.now(); let x = 0; for (let i = 0; i < 3e7; i++) x = (x + i * 7) % 1000003; r.cpu = performance.now() - t0 + (x < 0 ? 1 : 0); }
  let t = performance.now();
  const { doc } = await importDocx(bytes);
  r.import = performance.now() - t;
  t = performance.now();
  const first = paginated();
  loadDocument(ed, doc);
  r.load = performance.now() - t;
  await guard(first, 'первое разбиение');
  r.firstPaginate = performance.now() - t;
  r.firstMeasure = window.__peroMeasureMs;
  await settle();
  r.settled = performance.now() - t;
  r.pages = pagin().total;
  r.blocks = view.state.doc.childCount;
  // набор: символ в абзаце в начале, середине и конце документа (по 5 нажатий, медиана)
  const textPos = (frac) => { let want = Math.floor(view.state.doc.content.size * frac); let found = null;
    view.state.doc.descendants((n, p) => { if (found == null && n.isText && p >= want) found = p + 1; return found == null; }); return found; };
  const med = (a) => a.slice().sort((x, y) => x - y)[a.length >> 1];
  for (const [name, frac] of [['typeStart', 0.001], ['typeMiddle', 0.5], ['typeEnd', 0.98]]) {
    const pos = textPos(frac);
    const sync = []; const meas = []; const total = [];
    for (let i = 0; i < 5; i++) {
      const p = paginated();
      const t0 = performance.now();
      view.dispatch(view.state.tr.insertText('ж', pos + i));
      sync.push(performance.now() - t0);
      await guard(p, name + ' #' + i);
      meas.push(window.__peroMeasureMs);
      total.push(sync[i] + window.__peroMeasureMs);
      await settle();
    }
    r[name] = { sync: med(sync), measure: med(meas), total: med(total) };
  }
  // набор текста подряд (~8 знаков в секунду, как быстрый набор): задержка каждого нажатия — насколько позже
  // запланированного момента главный поток смог его обработать (занят разбиением и т. п.) + время самой транзакции
  for (const [name, frac] of [['burstMiddle', 0.5], ['burstStart', 0.001]]) {
    const pos = textPos(frac);
    const lags = [];
    const t0 = performance.now() + 50;
    await new Promise((done) => {
      let i = 0;
      const step = () => {
        const planned = t0 + i * 120;
        const late = performance.now() - planned;
        const a = performance.now();
        const text = i % 6 === 5 ? ' ' : 'ж';
        view.dispatch(view.state.tr.insertText(text, pos + i));
        lags.push(late + performance.now() - a);
        i++;
        if (i >= 40) { done(); return; }
        setTimeout(step, Math.max(0, t0 + i * 120 - performance.now()));
      };
      setTimeout(step, 50);
    });
    await settle();
    const sorted = lags.slice().sort((x, y) => x - y);
    r[name] = { mean: lags.reduce((x, y) => x + y, 0) / lags.length, p90: sorted[Math.floor(sorted.length * 0.9)], max: sorted[sorted.length - 1] };
  }
  // согласованность: после случайных правок (быстрый путь и обычные замеры) полный замер даёт то же разбиение
  {
    let seed = 7; const rnd = (n) => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed % n; };
    const key = () => pagin().key;
    let bad = 0; let runs = 0;
    for (let round = 0; round < 6; round++) {
      for (let s = 0; s < 12; s++) {
        const pos = textPos(rnd(1000) / 1000);
        const op = rnd(10);
        const tr = view.state.tr;
        if (op < 6) tr.insertText(op < 5 ? 'жжж ' : 'з', pos);
        else if (op < 8) tr.delete(pos, Math.min(pos + 3, view.state.doc.resolve(pos).end()));
        else if (op === 8) tr.split(pos);
        else { const $p = view.state.doc.resolve(pos); if ($p.depth === 1 && $p.index(0) > 0) tr.join($p.before(1)); }
        view.dispatch(tr);
        await sleep(60 + rnd(120));
      }
      await settle();
      const a = key();
      const p = paginated(); view.dom.dispatchEvent(new Event('pero-remeasure')); await guard(p, 'полный замер'); await settle();
      runs++;
      if (key() !== a) { bad++; const b = key(); if (!r.diff) { if (a == null || b == null) r.diff = 'ключ ' + (a == null ? 'до: null' : 'после: null'); else { let i = 0; while (a[i] === b[i]) i++; r.diff = a.slice(Math.max(0, i - 60), i + 60) + ' ≠ ' + b.slice(Math.max(0, i - 60), i + 60); } } }
    }
    r.consistency = runs - bad + ' из ' + runs;
  }
  t = performance.now();
  const saved = await exportDocx(view.state.doc.toJSON(), { source: bytes });
  r.export = performance.now() - t;
  r.exportKB = Math.round(saved.byteLength / 1024);
  for (const k of Object.keys(r)) if (typeof r[k] === 'number') r[k] = Math.round(r[k]);
  for (const k of Object.keys(r)) if (typeof r[k] === 'object') for (const j of Object.keys(r[k])) r[k][j] = Math.round(r[k][j]);
  return JSON.stringify(r);
} catch (e) { return JSON.stringify({ error: e.message + ' ' + e.stack }); } })()`);
ws.close();
const r = JSON.parse(out);
if (r.error) { console.log('ошибка в окне:', r.error); process.exit(1); }
console.log(`${file}: ${r.pages} стр., ${r.blocks} блоков верхнего уровня; эталонная нагрузка ${r.cpu} мс`);
console.log(`открытие: разбор ${r.import} мс, загрузка в редактор ${r.load} мс, первое разбиение готово через ${r.firstPaginate} мс (замер ${r.firstMeasure} мс), всё утихло через ${r.settled} мс`);
for (const k of ['typeStart', 'typeMiddle', 'typeEnd']) console.log(`набор (${{ typeStart: 'начало', typeMiddle: 'середина', typeEnd: 'конец' }[k]}): транзакция ${r[k].sync} мс + разбиение ${r[k].measure} мс = ${r[k].total} мс на нажатие`);
for (const k of ['burstMiddle', 'burstStart']) console.log(`набор подряд 40 знаков (${k === 'burstMiddle' ? 'середина' : 'начало'}): задержка нажатия в среднем ${r[k].mean} мс, 90 % — до ${r[k].p90} мс, худшее ${r[k].max} мс`);
console.log(`согласованность (после случайных правок полный замер даёт то же разбиение): ${r.consistency}`);
if (r.diff) console.log('первое расхождение:', r.diff);
console.log(`сохранение: ${r.export} мс (${r.exportKB} КБ)`);
