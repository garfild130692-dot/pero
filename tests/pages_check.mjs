// Сверка разбиения на страницы Перо с Word (Windows, установлен Word).
//
//   1) npx vite                                   — dev-сервер интерфейса
//   2) $env:VITE_DEV_SERVER_URL='http://127.0.0.1:5183'; npx electron . --remote-debugging-port=9223
//   3) node tests/pages_check.mjs tests/fixtures/pages.docx
//
// Word (через COM, tests/word_pages.ps1) даёт начало каждой страницы: смещение в абзаце и первые символы.
// Перо — то же из состояния плагина разбиения (документ загружается в открытое окно через CDP).
import { execFileSync } from 'child_process';
import { resolve } from 'path';

const file = process.argv[2] || 'tests/fixtures/pages.docx';

function wordPages() {
  const out = execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'tests/word_pages.ps1', '-path', resolve(file)]).toString('utf8');
  return JSON.parse(out.replace(/^﻿/, ''));
}

async function peroPages() {
  const list = await (await fetch('http://127.0.0.1:9223/json')).json();
  const page = list.find((p) => p.type === 'page');
  const ws = new WebSocket(page.webSocketDebuggerUrl);
  let id = 0;
  const pending = new Map();
  ws.onmessage = (m) => { const d = JSON.parse(m.data); if (pending.has(d.id)) { pending.get(d.id)(d); pending.delete(d.id); } };
  await new Promise((r) => { ws.onopen = r; });
  const run = async (expression) => {
    const i = ++id;
    const res = await new Promise((r) => { pending.set(i, r); ws.send(JSON.stringify({ id: i, method: 'Runtime.evaluate', params: { expression, awaitPromise: true, returnByValue: true } })); });
    if (res.result.exceptionDetails) throw new Error(res.result.exceptionDetails.exception?.description || 'ошибка в окне');
    return res.result.result.value;
  };
  await run(`(async () => {
    const { importDocx } = await import('/src/docx/import.js');
    const { loadDocument } = await import('/src/editor/extensions.js');
    const ed = document.querySelector('.ProseMirror').editor;
    loadDocument(ed, (await importDocx(await (await fetch('/${file.replace(/\\/g, '/')}')).arrayBuffer())).doc);
    await new Promise((r) => setTimeout(r, 3000));
  })()`);
  const result = await run(`(() => {
    const { state } = document.querySelector('.ProseMirror').editor;
    const st = state.plugins.find((p) => p.key.startsWith('pagination')).getState(state);
    return { total: st.total, breaks: st.breaks.map((pos) => {
      let $p = state.doc.resolve(Math.min(pos, state.doc.content.size));
      if (!$p.parent.isTextblock) {
        let tb = null;
        state.doc.nodesBetween(pos, Math.min(pos + 200, state.doc.content.size), (n, p) => { if (tb == null && n.isTextblock) tb = p; return tb == null; });
        if (tb != null) $p = state.doc.resolve(tb + 1);
      }
      const offset = $p.pos - $p.start();
      return { offset, text: $p.parent.textBetween(offset, $p.parent.content.size, '', '\\t').slice(0, 25) };
    }) };
  })()`);
  ws.close();
  return result;
}

const norm = (s) => String(s || '').replace(/[\t​]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 18);
const w = wordPages();
const p = await peroPages();
console.log(`страниц: Word ${w.pages}, Перо ${p.total}`);
let same = 0;
for (let i = 0; i < Math.max(w.breaks.length, p.breaks.length); i++) {
  const a = w.breaks[i]; const b = p.breaks[i];
  const ok = a && b && a.offset === b.offset && norm(a.text) === norm(b.text);
  if (ok) same++;
  console.log(`${ok ? '  ' : '≠ '}стр.${i + 2}: Word ${a ? `[${a.offset}] ${norm(a.text)}` : '—'}  |  Перо ${b ? `[${b.offset}] ${norm(b.text)}` : '—'}`);
}
console.log(`совпало начал страниц: ${same} из ${w.breaks.length}`);
