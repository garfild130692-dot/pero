// Классы параметров функций для записи формул .xls (этап 19.3): берутся из байт-кода, который сам Excel записал в
// tests/fixtures/formulas.xls (870 формул всех групп). Класс аргумента — класс его верхнего токена: R — ссылка, V —
// значение, A — массив. Запись формулы с неверным классом Excel считает иначе (неявное пересечение, #ЗНАЧ!).
// Запуск: node tests/gen_xls_classes.mjs → src/sheet/xlsClasses.js
import { readFileSync, writeFileSync } from 'node:fs';
import { readCfb } from '../src/formats/cfb.js';
import { FTAB, UDF } from '../src/sheet/xlsFtab.js';

const s = readCfb(readFileSync('tests/fixtures/formulas.xls')).stream('Workbook');
const u16 = (d, o) => d[o] | (d[o + 1] << 8);
const recs = [];
for (let p = 0; p + 4 <= s.length;) { const t = u16(s, p); const l = u16(s, p + 2); recs.push({ t, d: s.subarray(p + 4, p + 4 + l) }); p += 4 + l; }
// имена (NAME) и внешние имена (EXTERNNAME) — для вызовов по имени
const names = []; const ext = [];
for (const { t, d } of recs) {
  if (t === 0x18) { const cch = d[3]; const high = d[14] & 1; let n = ''; for (let i = 0; i < cch; i++) n += String.fromCharCode(high ? u16(d, 15 + i * 2) : d[15 + i]); names.push(n); }
  if (t === 0x23) { const cch = d[6]; const high = d[7] & 1; let n = ''; for (let i = 0; i < cch; i++) n += String.fromCharCode(high ? u16(d, 8 + i * 2) : d[8 + i]); ext.push(n); }
}
const clsOf = (t) => ['', 'R', 'V', 'A'][(t >> 5) & 3];
const stat = new Map(); // имя → [Map(класс → сколько)]
const note = (name, i, c) => {
  if (!c) return;
  if (!stat.has(name)) stat.set(name, []);
  const a = stat.get(name);
  while (a.length <= i) a.push(new Map());
  a[i].set(c, (a[i].get(c) || 0) + 1);
};
function walk(ce, onlyR = false) {
  const st = []; let p = 0;
  while (p < ce.length) {
    const t = ce[p];
    if (t >= 0x03 && t <= 0x11) { st.pop(); st.pop(); st.push({ c: null }); p++; continue; }
    switch (t) {
      case 0x01: case 0x02: return;
      case 0x12: case 0x13: case 0x14: st.pop(); st.push({ c: null }); p++; continue;
      case 0x15: p++; continue;
      case 0x16: st.push({ c: null }); p++; continue;
      case 0x17: { const cch = ce[p + 1]; const high = ce[p + 2] & 1; st.push({ c: null }); p += 3 + cch * (high ? 2 : 1); continue; }
      case 0x19: { const g = ce[p + 1]; if (g & 0x04) { p += 4 + 2 * (u16(ce, p + 2) + 1); continue; } if (g & 0x10) { st.pop(); st.push({ c: 'V' }); } p += 4; continue; }
      case 0x1C: case 0x1D: st.push({ c: null }); p += 2; continue;
      case 0x1E: st.push({ c: null }); p += 3; continue;
      case 0x1F: st.push({ c: null }); p += 9; continue;
      default: break;
    }
    const base = (t & 0x1F) | 0x20; const c = clsOf(t);
    const size = { 0x20: 8, 0x21: 3, 0x22: 4, 0x23: 5, 0x24: 5, 0x25: 9, 0x26: 7, 0x27: 7, 0x28: 7, 0x29: 3, 0x2A: 5, 0x2B: 9, 0x2C: 5, 0x2D: 9, 0x39: 7, 0x3A: 7, 0x3B: 11, 0x3C: 7, 0x3D: 11 }[base];
    if (!size) return;
    if (base === 0x21 || base === 0x22) {
      const iftab = u16(ce, base === 0x21 ? p + 1 : p + 2) & 0x7FFF;
      const f = FTAB.get(iftab);
      const n = base === 0x21 ? (f ? f.argc : 0) : ce[p + 1] & 0x7F;
      const args = st.splice(st.length - n, n);
      let name = f && iftab !== UDF ? f.name : null;
      if (iftab === UDF) { const first = args.shift(); name = first && first.name ? first.name.replace(/^_xlfn\.(_xlws\.)?/, '') : null; }
      if (name) args.forEach((a, i) => { if (!onlyR || a.c === 'R') note(name, i, a.c); });
      st.push({ c });
    } else if (base === 0x26 || base === 0x27 || base === 0x28 || base === 0x29) { /* подвыражение следует дальше */ } else {
      const e = { c };
      if (base === 0x23) e.name = names[(u16(ce, p + 1) | (u16(ce, p + 3) << 16)) - 1];
      if (base === 0x39) e.name = ext[u16(ce, p + 3) - 1];
      st.push(e);
    }
    p += size;
  }
}
for (const { t, d } of recs) {
  if (t === 0x06) { const cce = u16(d, 20); const ce = d.subarray(22, 22 + cce); if (ce[0] !== 0x01) walk(ce); }
  if (t === 0x04BC) { const cce = u16(d, 8); walk(d.subarray(10, 10 + cce)); }
  // формулы массива: там значения пишутся классом A — берём только параметры-ссылки (R)
  if (t === 0x0221) { const cce = u16(d, 12); walk(d.subarray(14, 14 + cce), true); }
}
const table = {};
for (const [name, a] of [...stat].sort((x, y) => (x[0] < y[0] ? -1 : 1))) {
  table[name] = a.map((m) => [...m].sort((x, y) => y[1] - x[1])[0]?.[0] || 'V').join('');
}
// Параметры, которые по грамматике Ftab ([MS-XLS] 2.5.198.17) принимают только ссылку (ref), — класс R, даже если в
// эталоне Excel им попалась константа или их нет вовсе (последний класс действует для остальных аргументов)
const REF_ONLY = {
  ROW: 'R', COLUMN: 'R', AREAS: 'R', OFFSET: 'RV', CELL: 'VR', RANK: 'VRV', SUBTOTAL: 'VR', SUMIF: 'RVR', COUNTIF: 'RV',
  COUNTBLANK: 'R', PHONETIC: 'R', ISREF: 'R',
  DCOUNT: 'RRR', DSUM: 'RRR', DAVERAGE: 'RRR', DMIN: 'RRR', DMAX: 'RRR', DSTDEV: 'RRR', DVAR: 'RRR', DPRODUCT: 'RRR',
  DSTDEVP: 'RRR', DVARP: 'RRR', DCOUNTA: 'RRR', DGET: 'RRR',
};
for (const [name, cls] of Object.entries(REF_ONLY)) {
  const cur = table[name] || '';
  const n = Math.max(cur.length, cls.length);
  let out = '';
  for (let i = 0; i < n; i++) out += cls[Math.min(i, cls.length - 1)] === 'R' ? 'R' : (cur[i] || cls[Math.min(i, cls.length - 1)]);
  table[name] = out;
}
// русские имена пакета анализа → английские: канонические имена дописывает xls.js при чтении; здесь — как записал Excel
const lines = Object.entries(table).map(([k, v]) => `  ${JSON.stringify(k)}: '${v}',`);
writeFileSync('src/sheet/xlsClasses.js', `// Классы параметров функций в формулах .xls: по позиции аргумента — R (ссылка), V (значение), A (массив); последний
// класс действует для остальных аргументов. Получено из байт-кода, который записал сам Excel (tests/fixtures/formulas.xls).
// Генерирует tests/gen_xls_classes.mjs — вручную не править.
export const PARAM_CLASS = {
${lines.join('\n')}
};
`);
console.log('функций:', Object.keys(table).length);
