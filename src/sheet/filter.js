// Автофильтр (как в Excel). Лист: autoFilter = { ref: { r1, c1, r2, c2 } — с строкой заголовков,
// cols: { [номер столбца от c1]: условие }, raw: { [номер]: исходный XML непонятного условия (top10, цвет…) } }.
// Условие: { values: [тексты как на экране], blank: bool, dates: [{ year, month?, day?, hour?, minute?, second? }] }
// или { custom: { and: bool, conds: [{ op, val }] } } (op как в файле: equal, notEqual, greaterThan, …; в val у
// равенства — подстановочные * ?). Строки, не прошедшие фильтр, скрыты: rowInfo[r] = { hidden, filtered }.
import { formatValue } from './format.js';
import { ERR_RU, serialToDate, wildcard } from './formula.js';
import { key } from './address.js';

const isErr = (v) => !!v && typeof v === 'object' && 'error' in v;
// текст ячейки, как на экране (по нему Excel сравнивает списки значений)
export function shownText(cell) {
  if (!cell || cell.v == null || cell.v === '') return '';
  if (isErr(cell.v)) return ERR_RU[cell.v.error] || cell.v.error;
  if (typeof cell.v === 'boolean') return cell.v ? 'ИСТИНА' : 'ЛОЖЬ';
  return formatValue(cell.v, cell.s && cell.s.numFmt);
}
const isBlank = (cell) => !cell || cell.v == null || cell.v === '';
const numOf = (s) => { const t = String(s).replace(/\s/g, '').replace(',', '.'); return t !== '' && isFinite(+t) ? +t : null; };

function customMatch({ op, val }, cell) {
  const v = cell ? cell.v : null;
  const n = numOf(val);
  if (n !== null && typeof v === 'number') {
    switch (op) {
      case 'equal': return v === n; case 'notEqual': return v !== n;
      case 'greaterThan': return v > n; case 'greaterThanOrEqual': return v >= n;
      case 'lessThan': return v < n; case 'lessThanOrEqual': return v <= n;
      default: return true;
    }
  }
  const text = shownText(cell);
  if (op === 'equal' || op === 'notEqual') {
    const re = wildcard(String(val));
    const hit = re ? re.test(text) : text.toLowerCase() === String(val).toLowerCase();
    return op === 'equal' ? hit : !hit;
  }
  if (typeof v === 'number' || n !== null) return false; // число со строкой не сравнивается
  const c = text.localeCompare(String(val), 'ru', { sensitivity: 'base' });
  switch (op) {
    case 'greaterThan': return c > 0; case 'greaterThanOrEqual': return c >= 0;
    case 'lessThan': return c < 0; case 'lessThanOrEqual': return c <= 0;
    default: return true;
  }
}
const DATE_PARTS = ['year', 'month', 'day', 'hour', 'minute', 'second'];
function dateMatch(item, v) {
  if (typeof v !== 'number') return false;
  const d = serialToDate(v);
  const got = { year: d.y, month: d.m, day: d.d, hour: d.hh, minute: d.mm, second: d.ss };
  return DATE_PARTS.every((p) => item[p] == null || item[p] === got[p]);
}
export function matches(crit, cell) {
  if (!crit) return true;
  if (crit.custom) {
    const rs = crit.custom.conds.map((c) => customMatch(c, cell));
    return crit.custom.and ? rs.every(Boolean) : rs.some(Boolean);
  }
  if (isBlank(cell)) return !!crit.blank;
  if (crit.dates && crit.dates.some((it) => dateMatch(it, cell.v))) return true;
  const t = shownText(cell).toLowerCase();
  return (crit.values || []).some((x) => String(x).toLowerCase() === t);
}

// Строки данных фильтра: от строки под заголовком до последней непустой строки «текущей области» (как Excel —
// фильтр захватывает дописанные снизу строки)
export function dataRows(sheet, af) {
  const { r1, c1, c2 } = af.ref;
  let r2 = af.ref.r2;
  const filled = (r) => { for (let c = c1; c <= c2; c++) { const x = sheet.cells.get(key(r, c)); if (x && x.v != null && x.v !== '') return true; } return false; };
  while (r2 < 1048575 && filled(r2 + 1)) r2++;
  return { r1: r1 + 1, r2 };
}
// Новые rowInfo: строки, не прошедшие условия всех столбцов, скрыты (filtered); прошедшие — показаны
export function applyFilter(sheet, af) {
  const rowInfo = { ...(sheet.rowInfo || {}) };
  const { r1, r2 } = dataRows(sheet, af);
  const cols = Object.entries(af.cols || {});
  for (let r = r1; r <= r2; r++) {
    const pass = cols.every(([j, crit]) => matches(crit, sheet.cells.get(key(r, af.ref.c1 + +j))));
    const info = { ...(rowInfo[r] || {}) };
    if (pass) { delete info.hidden; delete info.filtered; } else { info.hidden = true; info.filtered = true; }
    if (Object.keys(info).length) rowInfo[r] = info; else delete rowInfo[r];
  }
  return { rowInfo, ref: { ...af.ref, r2 } };
}
// Значения столбца для списка в меню фильтра: [{ text, sort }] без повторов + есть ли пустые.
// Строки, скрытые фильтром по ДРУГИМ столбцам, не показываются (как Excel)
export function columnValues(sheet, af, j) {
  const { r1, r2 } = dataRows(sheet, af);
  const other = Object.entries(af.cols || {}).filter(([k]) => +k !== j);
  const seen = new Map(); let blank = false;
  for (let r = r1; r <= r2; r++) {
    if (!other.every(([k, crit]) => matches(crit, sheet.cells.get(key(r, af.ref.c1 + +k))))) continue;
    const cell = sheet.cells.get(key(r, af.ref.c1 + j));
    if (isBlank(cell)) { blank = true; continue; }
    const text = shownText(cell);
    const k = text.toLowerCase();
    if (!seen.has(k)) seen.set(k, { text, num: typeof cell.v === 'number' ? cell.v : null });
  }
  const list = [...seen.values()].sort((a, b) => (a.num !== null && b.num !== null ? a.num - b.num : a.num !== null ? -1 : b.num !== null ? 1 : a.text.localeCompare(b.text, 'ru', { sensitivity: 'base' })));
  return { list, blank };
}
// Сколько строк данных видно / всего (строка состояния: «Найдено записей: 5 из 20»)
export function filterCount(sheet, af) {
  const { r1, r2 } = dataRows(sheet, af);
  let shown = 0;
  for (let r = r1; r <= r2; r++) if (!(sheet.rowInfo && sheet.rowInfo[r] && sheet.rowInfo[r].hidden)) shown++;
  return { shown, total: r2 - r1 + 1 };
}

// ---------- файл ----------
const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
// <filterColumn> → условие (непонятное — null, XML сохраняется вызывающим)
export function readFilterColumn(el) {
  const child = [...el.childNodes].find((n) => n.nodeType === 1);
  if (!child) return null;
  const name = child.localName || child.nodeName.replace(/^.*:/, '');
  const kids = (n) => [...n.childNodes].filter((x) => x.nodeType === 1);
  if (name === 'filters') {
    const crit = { values: [], blank: child.getAttribute('blank') === '1' };
    for (const f of kids(child)) {
      const fn = f.localName || f.nodeName;
      if (fn === 'filter') crit.values.push(f.getAttribute('val') || '');
      else if (fn === 'dateGroupItem') {
        const it = {};
        const g = f.getAttribute('dateTimeGrouping');
        for (const p of DATE_PARTS) { if (f.getAttribute(p) != null) it[p] = +f.getAttribute(p); if (p === g) break; }
        (crit.dates = crit.dates || []).push(it);
      }
    }
    return crit;
  }
  if (name === 'customFilters') {
    return { custom: { and: child.getAttribute('and') === '1', conds: kids(child).map((c) => ({ op: c.getAttribute('operator') || 'equal', val: c.getAttribute('val') || '' })) } };
  }
  return null;
}
export function filterColumnXml(j, crit) {
  if (crit.custom) {
    const conds = crit.custom.conds.map((c) => `<customFilter${c.op !== 'equal' ? ` operator="${c.op}"` : ''} val="${xmlEsc(c.val)}"/>`).join('');
    return `<filterColumn colId="${j}"><customFilters${crit.custom.and ? ' and="1"' : ''}>${conds}</customFilters></filterColumn>`;
  }
  const vals = (crit.values || []).map((v) => `<filter val="${xmlEsc(v)}"/>`).join('');
  const dates = (crit.dates || []).map((d) => {
    const parts = DATE_PARTS.filter((p) => d[p] != null);
    return `<dateGroupItem${parts.map((p) => ` ${p}="${d[p]}"`).join('')} dateTimeGrouping="${parts[parts.length - 1]}"/>`;
  }).join('');
  return `<filterColumn colId="${j}"><filters${crit.blank ? ' blank="1"' : ''}>${vals}${dates}</filters></filterColumn>`;
}
