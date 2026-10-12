// Сверка обтекания рисунков Перо с Word: начало каждой строки (лист, x, y в pt от края листа, первое слово).
//   1) npx vite;  2) $env:VITE_DEV_SERVER_URL='http://127.0.0.1:5183'; npx electron . --remote-debugging-port=9223
//   3) node tests/wrap_check.mjs   (эталон — tests/fixtures/wrap_word.json от tests/make_wrap_fixture.ps1)
import { readFileSync } from 'node:fs';
const word = JSON.parse(readFileSync('tests/fixtures/wrap_word.json', 'utf8')).lines;
const list = await (await fetch('http://127.0.0.1:9223/json')).json();
const ws = new WebSocket(list.find((p) => p.type === 'page').webSocketDebuggerUrl);
await new Promise((r) => ws.addEventListener('open', r));
const send = (expression) => new Promise((res) => {
  ws.addEventListener('message', function on(e) {
    const m = JSON.parse(e.data);
    if (m.id !== 1) return;
    ws.removeEventListener('message', on);
    res(m.result && m.result.result ? m.result.result.value : JSON.stringify({ error: JSON.stringify(m).slice(0, 300) }));
  });
  ws.send(JSON.stringify({ id: 1, method: 'Runtime.evaluate', params: { expression, awaitPromise: true, returnByValue: true } }));
});
const pero = await send(`(async () => { try {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  document.querySelector('.start-tile')?.click(); await sleep(1000);
  const ed = document.querySelector('.ProseMirror').editor;
  const { importDocx } = await import('/src/docx/import.js');
  const { loadDocument } = await import('/src/editor/extensions.js');
  const { doc } = await importDocx(await (await fetch('/tests/fixtures/wrap.docx')).arrayBuffer());
  loadDocument(ed, doc); await sleep(3000);
  const root = ed.view.dom; const page = root.closest('.page');
  const pr = page.getBoundingClientRect(); const scale = pr.width / page.offsetWidth;
  const bands = [...page.querySelectorAll(':scope > .pero-bands > .pero-page-band')].sort((a, b) => a.offsetTop - b.offsetTop);
  // верх листа: после нижнего поля и промежутка полосы (как pageTopOf), не после верхнего поля
  const tops = [0, ...bands.map((b) => b.offsetTop + b.children[0].offsetHeight + b.children[1].offsetHeight)];
  const pageOf = (y) => { let k = 0; while (k + 1 < tops.length && tops[k + 1] <= y + 1) k++; return k; };
  const lines = []; let prev = null;
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest('[contenteditable=false]') ? 2 : 1) });
  const r = document.createRange();
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    for (const m of n.data.matchAll(/\\S+/g)) {
      r.setStart(n, m.index); r.setEnd(n, m.index + 1);
      const x = r.getClientRects()[0]; if (!x) continue;
      const y = (x.top - pr.top) / scale; const k = pageOf(y);
      const pt = { page: k + 1, x: (x.left - pr.left) / scale * 0.75, y: (y - tops[k]) * 0.75, text: m[0] };
      if (!prev || prev.page !== pt.page || Math.abs(pt.y - prev.y) > 1) lines.push(pt);
      prev = pt;
    }
  }
  return JSON.stringify(lines);
} catch (e) { return JSON.stringify({ error: e.message }); } })()`);
ws.close();
const ours = JSON.parse(pero || '[]');
if (ours.error) { console.log('ошибка в окне:', ours.error); process.exit(1); }
// сравнение по абзацам (каскад от одного абзаца не портит остальные): строки абзаца — первое слово, x и высота от
// первой строки абзаца (±3 pt по x, ±4 pt по высоте)
const paras = (lines) => {
  const out = [];
  for (const l of lines) { if (l.text === 'Абзац' || !out.length) out.push([]); out[out.length - 1].push(l); }
  return out;
};
const W = paras(word); const P = paras(ours);
let okLines = 0; let total = 0;
W.forEach((wl, i) => {
  const pl = P[i] || [];
  const w0 = wl[0]; const p0 = pl[0];
  let ok = 0;
  wl.forEach((a, j) => {
    total++;
    const b = pl[j];
    if (b && a.text === b.text && Math.abs(a.x - b.x) <= 3 && Math.abs((a.y - w0.y) - (b.y - p0.y)) <= 4) ok++;
  });
  okLines += ok;
  const first = wl.find((a, j) => { const b = pl[j]; return !(b && a.text === b.text && Math.abs(a.x - b.x) <= 3 && Math.abs((a.y - w0.y) - (b.y - p0.y)) <= 4); });
  const j = first ? wl.indexOf(first) : -1;
  console.log(`${ok === wl.length && pl.length === wl.length ? '=' : '≠'} абзац ${i + 1}: строк Word ${wl.length}, Перо ${pl.length}, совпало ${ok}` +
    (j >= 0 ? `; первое расхождение: Word x${first.x} +${(first.y - w0.y).toFixed(1)} ${first.text} | Перо ${pl[j] ? `x${pl[j].x.toFixed(1)} +${(pl[j].y - p0.y).toFixed(1)} ${pl[j].text}` : '—'}` : ''));
});
console.log(`строк совпало: ${okLines} из ${total}`);
