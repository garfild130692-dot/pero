// Проверка данных (как в Excel). Лист: dv = [правило]. Правило: { ranges, type: any|whole|decimal|list|date|time|
//   textLength|custom, operator (between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual,
//   lessThanOrEqual), formula1, formula2 (канонические формулы: "a,b,c" в кавычках или $A$1:$A$5 для списка),
//   allowBlank, inCellDropdown (в файле наоборот: showDropDown="1" — список СКРЫТ), errorStyle: stop|warning|information,
//   showError, errorTitle, error, showPrompt, promptTitle, prompt, raw — исходная разметка (неизменённое правило) }.
import { key } from './address.js';
import { canonical, evaluate, fileFormula, isErr, offsetFormula, parse } from './formula.js';
import { viewContext } from './model.js';

const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const textEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); // текст элемента (кавычки — как есть, как у Excel)
const kids = (el, name) => [...(el ? el.childNodes : [])].filter((n) => n.nodeType === 1 && (!name || (n.localName || n.nodeName) === name));
const kid = (el, name) => kids(el, name)[0] || null;
const on = (el, a) => ['1', 'true'].includes(el.getAttribute(a));

// правило, которое действует в ячейке (первое по порядку, как Excel)
export function dvAt(sheet, r, c) {
  for (const rule of sheet.dv || []) if (rule.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)) return rule;
  return null;
}

// ---------- файл ----------
export function readDv(root, { parseRange, serialize }) {
  const out = [];
  for (const el of kids(kid(root, 'dataValidations'), 'dataValidation')) {
    const ranges = String(el.getAttribute('sqref') || '').split(/\s+/).filter(Boolean).map(parseRange).filter(Boolean);
    if (!ranges.length) continue;
    const f = (n) => { const x = kid(el, n); return x ? canonical(x.textContent) : null; };
    const rule = {
      ranges, type: el.getAttribute('type') || 'any', operator: el.getAttribute('operator') || 'between',
      formula1: f('formula1'), formula2: f('formula2'), allowBlank: on(el, 'allowBlank'), inCellDropdown: !on(el, 'showDropDown'),
      errorStyle: el.getAttribute('errorStyle') || 'stop', showError: on(el, 'showErrorMessage'), errorTitle: el.getAttribute('errorTitle') || '',
      error: el.getAttribute('error') || '', showPrompt: on(el, 'showInputMessage'), promptTitle: el.getAttribute('promptTitle') || '', prompt: el.getAttribute('prompt') || '',
    };
    // формулы в расширениях (x14:dataValidation со ссылкой на другой лист) Перо не читает — правило сохраняется как есть
    rule.raw = serialize(el);
    out.push(rule);
  }
  return out;
}
export const dvSignature = (dv) => JSON.stringify((dv || []).map(({ raw, edited, ...r }) => r));
export function dvXml(dv, rangeName) {
  if (!dv || !dv.length) return '';
  const items = dv.map((r) => {
    const sqref = r.ranges.map(rangeName).join(' ');
    if (r.raw && !r.edited) return r.raw.replace(/\ssqref="[^"]*"/, ` sqref="${sqref}"`);
    const a = [];
    if (r.type && r.type !== 'any') a.push(`type="${r.type}"`);
    if (r.errorStyle && r.errorStyle !== 'stop') a.push(`errorStyle="${r.errorStyle}"`);
    if (r.operator && r.operator !== 'between' && r.type !== 'list' && r.type !== 'custom' && r.type !== 'any') a.push(`operator="${r.operator}"`);
    if (r.allowBlank) a.push('allowBlank="1"');
    if (r.type === 'list' && !r.inCellDropdown) a.push('showDropDown="1"');
    if (r.showPrompt) a.push('showInputMessage="1"');
    if (r.showError) a.push('showErrorMessage="1"');
    for (const [k, n] of [['errorTitle', 'errorTitle'], ['error', 'error'], ['promptTitle', 'promptTitle'], ['prompt', 'prompt']]) if (r[k]) a.push(`${n}="${xmlEsc(r[k])}"`);
    a.push(`sqref="${sqref}"`);
    const fs = (r.formula1 != null ? `<formula1>${textEsc(fileFormula(r.formula1))}</formula1>` : '') + (r.formula2 != null && ['between', 'notBetween'].includes(r.operator) && r.type !== 'list' ? `<formula2>${textEsc(fileFormula(r.formula2))}</formula2>` : '');
    return `<dataValidation ${a.join(' ')}${fs ? `>${fs}</dataValidation>` : '/>'}`;
  });
  return `<dataValidations count="${items.length}">${items.join('')}</dataValidations>`;
}

// ---------- значения ----------
// формула правила, вычисленная для ячейки (r, c) — ссылки сдвигаются относительно левой верхней ячейки диапазона
function evalFor(wb, si, rule, f, r, c) {
  if (f == null) return null;
  const g = rule.ranges[0];
  let ast;
  try { ast = parse(offsetFormula(f, r - g.r1, c - g.c1)); } catch { return null; }
  return evaluate(ast, viewContext(wb, si, r, c));
}
// Элементы выпадающего списка: «а,б,в» или значения диапазона/имени
export function listItems(wb, si, rule, r, c) {
  const f = rule.formula1;
  if (f == null) return [];
  const lit = /^"(.*)"$/s.exec(f);
  if (lit) return lit[1].replace(/""/g, '"').split(',').map((x) => x.trim()).filter((x) => x !== '');
  const v = evalFor(wb, si, rule, f, r, c);
  const flat = Array.isArray(v) ? v.flat() : [v];
  return flat.filter((x) => x != null && x !== '' && !isErr(x)).map((x) => (typeof x === 'number' ? String(x).replace('.', ',') : String(x)));
}
const num = (v) => (typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && isFinite(+v.replace(',', '.')) ? +v.replace(',', '.') : null);
// Проверить значение (уже разобранное, как его сохранит ячейка): true — подходит
export function validate(wb, si, r, c, value) {
  const rule = dvAt(wb.sheets[si], r, c);
  if (!rule || rule.type === 'any') return true;
  if (value == null || value === '') return true; // пустую ячейку Excel не проверяет (очистить можно всегда)
  if (rule.type === 'list') {
    // список текстом («да,нет») Excel сравнивает с учётом регистра, список из диапазона — без (проверено через COM)
    const literal = /^".*"$/s.test(rule.formula1 || '');
    const norm = (x) => (literal ? x : x.toLowerCase());
    const items = listItems(wb, si, rule, r, c).map(norm);
    const t = typeof value === 'number' ? String(value).replace('.', ',') : norm(String(value));
    return items.includes(t) || (typeof value === 'number' && items.some((x) => num(x) === value));
  }
  if (rule.type === 'custom') { const x = evalFor(wb, si, rule, rule.formula1, r, c); return x === true || (typeof x === 'number' && x !== 0); }
  let x = rule.type === 'textLength' ? String(value).length : num(value);
  if (x == null) return false;
  if (rule.type === 'whole' && !Number.isInteger(x)) return false;
  if (rule.type === 'time') x -= Math.floor(x); // время — доля суток
  const a = num(evalFor(wb, si, rule, rule.formula1, r, c));
  const b = num(evalFor(wb, si, rule, rule.formula2, r, c));
  switch (rule.operator || 'between') {
    case 'between': return a != null && b != null && x >= Math.min(a, b) && x <= Math.max(a, b);
    case 'notBetween': return a != null && b != null && (x < Math.min(a, b) || x > Math.max(a, b));
    case 'equal': return x === a; case 'notEqual': return x !== a;
    case 'greaterThan': return x > a; case 'lessThan': return x < a;
    case 'greaterThanOrEqual': return x >= a; case 'lessThanOrEqual': return x <= a;
    default: return true;
  }
}
export const DV_DEFAULT_ERROR = 'Значение не соответствует ограничениям проверки данных, установленным для этой ячейки.';
export { key };
