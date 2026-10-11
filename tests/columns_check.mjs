// Сверка колонок Перо с Word: начало каждой колонки на каждом листе.
//   1) npx vite;  2) $env:VITE_DEV_SERVER_URL='http://127.0.0.1:5183'; npx electron . --remote-debugging-port=9223
//   3) node tests/columns_check.mjs   (эталон — tests/fixtures/columns_word.json от tests/make_columns_fixture.ps1)
// Правило начала колонки (как в make_columns_fixture.ps1): слово на новом листе или «поднялось» на том же листе правее.
import { readFileSync } from 'node:fs';
const word = JSON.parse(readFileSync('tests/fixtures/columns_word.json', 'utf8'));
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
  const { doc } = await importDocx(await (await fetch('/tests/fixtures/columns.docx')).arrayBuffer());
  loadDocument(ed, doc); await sleep(2500);
  const root = ed.view.dom; const page = root.closest('.page');
  const pr = page.getBoundingClientRect(); const scale = pr.width / page.offsetWidth;
  const bands = [...page.querySelectorAll(':scope > .pero-bands > .pero-page-band')].sort((a, b) => a.offsetTop - b.offsetTop);
  const tops = [0, ...bands.map((b) => b.offsetTop + b.offsetHeight)];
  const pageOf = (y) => { let k = 0; while (k + 1 < tops.length && tops[k + 1] <= y + 1) k++; return k + 1; };
  const starts = []; let prev = null;
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest('[contenteditable=false]') ? 2 : 1) });
  const r = document.createRange();
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    for (const m of n.data.matchAll(/\\S+/g)) {
      r.setStart(n, m.index); r.setEnd(n, m.index + 1);
      const x = r.getClientRects()[0]; if (!x) continue;
      const y = (x.top - pr.top) / scale; const xx = (x.left - pr.left) / scale; const pg = pageOf(y);
      if (!prev || pg !== prev.pg || (y < prev.y - 1 && xx > prev.x + 20)) starts.push({ page: pg, text: (n.data.slice(m.index) + ' ').slice(0, 25) });
      prev = { pg, y, x: xx };
    }
  }
  return JSON.stringify(starts);
} catch (e) { return JSON.stringify({ error: e.message }); } })()`);
ws.close();
const ours = JSON.parse(pero || '[]');
if (ours.error) { console.log('ошибка в окне:', ours.error); process.exit(1); }
const norm = (s) => s.replace(/\s+/g, ' ').trim().slice(0, 18);
let ok = 0;
const n = Math.max(word.length, ours.length);
for (let i = 0; i < n; i++) {
  const a = word[i]; const b = ours[i];
  const same = a && b && a.page === b.page && norm(a.text) === norm(b.text);
  if (same) ok++;
  console.log(`${same ? '=' : '≠'} ${a ? `стр.${a.page} ${norm(a.text)}` : '—'}  |  ${b ? `стр.${b.page} ${norm(b.text)}` : '—'}`);
}
console.log(`совпало начал колонок: ${ok} из ${word.length}`);
