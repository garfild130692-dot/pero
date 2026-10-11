// Формулы Word (OMML, m:oMath / m:oMathPara) → MathML для показа на экране и при печати (Chromium рисует MathML Core сам).
// Файл не меняется: узел opaque хранит исходный XML и пишет его обратно как есть; здесь — только вид.
import { DOMParser } from '@xmldom/xmldom';

const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const kids = (n) => {
  const out = [];
  for (let c = n && n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c);
  return out;
};
const isM = (n, name) => n && n.namespaceURI === M_NS && n.localName === name;
const child = (n, name) => kids(n).find((c) => isM(c, name)) || null;
// значение m:val свойства (…Pr/имя); нет элемента — undefined, есть без m:val — '' (для chr — «нет символа»)
const prop = (n, pr, name) => {
  const p = child(n, pr);
  const e = p && child(p, name);
  if (!e) return undefined;
  const v = e.getAttributeNS(M_NS, 'val') || e.getAttribute('m:val');
  return v == null ? '' : v;
};
const on = (v) => v !== undefined && v !== '0' && v !== 'off' && v !== 'false';

const mrow = (parts) => (parts.length === 1 ? parts[0] : `<mrow>${parts.join('')}</mrow>`);
const mo = (ch, attrs = '') => `<mo${attrs}>${esc(ch)}</mo>`;

// текст прогона: числа — mn, буквы — mi (по одной, курсивом; «прямой» стиль — словом), прочее — mo; дефис — минус
function runText(text, sty) {
  const plain = sty === 'p' || sty === 'b';
  const variant = sty === 'b' ? 'bold' : sty === 'bi' ? 'bold-italic' : null;
  const out = [];
  const re = /(\d+(?:[.,]\d+)*)|(\p{L}+)|(\s+)|([^])/gu;
  let m;
  while ((m = re.exec(text))) {
    if (m[1]) out.push(`<mn${variant === 'bold' || variant === 'bold-italic' ? ' mathvariant="bold"' : ''}>${esc(m[1])}</mn>`);
    else if (m[2]) {
      const v = (w) => (variant ? ` mathvariant="${variant}"` : plain && w.length === 1 ? ' mathvariant="normal"' : '');
      if (plain) out.push(`<mi${v(m[2])}>${esc(m[2])}</mi>`);
      else for (const ch of m[2]) out.push(`<mi${v(ch)}>${esc(ch)}</mi>`);
    } else if (m[3]) out.push('<mspace width="0.25em"/>');
    else out.push(mo(m[4] === '-' ? '−' : m[4]));
  }
  return out;
}

// комбинирующие знаки акцентов Word → самостоятельные символы (так их растягивает и ставит mover)
const ACCENTS = {
  '̀': '`', '́': '´', '̂': '^', '̃': '~', '̄': '¯', '̅': '‾', '̆': '˘',
  '̇': '˙', '̈': '¨', '̌': 'ˇ', '⃐': '↼', '⃑': '⇀', '⃖': '←',
  '⃗': '→', '⃡': '↔',
};
// черта над/под выражением: растянутый символ Chromium рисует едва заметным — линия рамкой
const OVERLINES = '̄̅¯‾';
const overline = (body, side) => `<mrow style="border-${side}:0.06em solid currentColor;padding-${side}:0.08em">${body}</mrow>`;
const INTEGRALS = '∫∬∭∮∯∰∱∲∳';

// содержимое аргумента списком элементов MathML (прогон — отдельными лексемами, иначе mrow не обернёт «x+1»)
const items = (el) => kids(el).flatMap((c) => (isM(c, 'r') ? runOf(c) : [conv(c)])).filter(Boolean);
const group = (el) => (el ? mrow(items(el)) : '<mrow/>');
function runOf(n) {
  const t = kids(n).filter((c) => isM(c, 't')).map((c) => c.textContent).join('');
  return runText(t, on(prop(n, 'rPr', 'nor')) ? 'p' : prop(n, 'rPr', 'sty'));
}

function conv(n) {
  if (n.namespaceURI !== M_NS) return kids(n).map(conv).join(''); // w:… внутри формулы — по содержимому
  const arg = (name) => group(child(n, name));
  switch (n.localName) {
    case 'r': return runOf(n).join('');
    case 'f': {
      const type = prop(n, 'fPr', 'type');
      if (type === 'lin' || type === 'skw') return `<mrow>${arg('num')}${mo('/')}${arg('den')}</mrow>`;
      return `<mfrac${type === 'noBar' ? ' linethickness="0"' : ''}>${arg('num')}${arg('den')}</mfrac>`;
    }
    case 'sSup': return `<msup>${arg('e')}${arg('sup')}</msup>`;
    case 'sSub': return `<msub>${arg('e')}${arg('sub')}</msub>`;
    case 'sSubSup': return `<msubsup>${arg('e')}${arg('sub')}${arg('sup')}</msubsup>`;
    case 'sPre': return `<mmultiscripts>${arg('e')}<mprescripts/>${arg('sub')}${arg('sup')}</mmultiscripts>`;
    case 'rad': {
      const deg = child(n, 'deg');
      if (on(prop(n, 'radPr', 'degHide')) || !deg || !kids(deg).length) return `<msqrt>${arg('e')}</msqrt>`;
      return `<mroot>${arg('e')}${arg('deg')}</mroot>`;
    }
    case 'd': {
      const beg = prop(n, 'dPr', 'begChr'); const end = prop(n, 'dPr', 'endChr'); const sep = prop(n, 'dPr', 'sepChr');
      const parts = [];
      const f = ' fence="true" stretchy="true" lspace="0" rspace="0"';
      if ((beg ?? '(') !== '') parts.push(mo(beg ?? '(', f));
      kids(n).filter((c) => isM(c, 'e')).forEach((e, i) => {
        if (i) parts.push(mo(sep ?? '|', ' separator="true" stretchy="true"'));
        parts.push(group(e));
      });
      if ((end ?? ')') !== '') parts.push(mo(end ?? ')', f));
      return `<mrow>${parts.join('')}</mrow>`;
    }
    case 'nary': {
      const chr = prop(n, 'naryPr', 'chr') || '∫';
      const loc = prop(n, 'naryPr', 'limLoc') || (INTEGRALS.includes(chr) ? 'subSup' : 'undOvr');
      const sub = !on(prop(n, 'naryPr', 'subHide')) && child(n, 'sub') && kids(child(n, 'sub')).length ? arg('sub') : null;
      const sup = !on(prop(n, 'naryPr', 'supHide')) && child(n, 'sup') && kids(child(n, 'sup')).length ? arg('sup') : null;
      const op = mo(chr, ' largeop="true" movablelimits="false"');
      const und = loc === 'undOvr';
      let base = op;
      if (sub && sup) base = und ? `<munderover>${op}${sub}${sup}</munderover>` : `<msubsup>${op}${sub}${sup}</msubsup>`;
      else if (sub) base = und ? `<munder>${op}${sub}</munder>` : `<msub>${op}${sub}</msub>`;
      else if (sup) base = und ? `<mover>${op}${sup}</mover>` : `<msup>${op}${sup}</msup>`;
      return `<mrow>${base}${arg('e')}</mrow>`;
    }
    case 'func': return `<mrow>${arg('fName')}${mo('⁡')}${arg('e')}</mrow>`;
    case 'limLow': return `<munder>${arg('e')}${arg('lim')}</munder>`;
    case 'limUpp': return `<mover>${arg('e')}${arg('lim')}</mover>`;
    case 'acc': {
      const c = prop(n, 'accPr', 'chr') ?? '̂';
      if (OVERLINES.includes(c)) return overline(arg('e'), 'top');
      return `<mover accent="true">${arg('e')}${mo(ACCENTS[c] || c, ' stretchy="true"')}</mover>`;
    }
    case 'bar': {
      return overline(arg('e'), prop(n, 'barPr', 'pos') === 'top' ? 'top' : 'bottom');
    }
    case 'groupChr': {
      const c = prop(n, 'groupChrPr', 'chr') ?? '⏟';
      const top = prop(n, 'groupChrPr', 'pos') === 'top';
      const g = mo(c, ' stretchy="true"');
      return top ? `<mover>${arg('e')}${g}</mover>` : `<munder>${arg('e')}${g}</munder>`;
    }
    case 'borderBox': return `<mrow style="border:1px solid currentColor;padding:1px 2px">${arg('e')}</mrow>`;
    case 'box': return arg('e');
    case 'phant': return on(prop(n, 'phantPr', 'show')) || prop(n, 'phantPr', 'show') === undefined ? arg('e') : `<mphantom>${arg('e')}</mphantom>`;
    case 'm': {
      const rows = kids(n).filter((c) => isM(c, 'mr')).map((r) => `<mtr>${kids(r).filter((c) => isM(c, 'e'))
        .map((e) => `<mtd>${group(e)}</mtd>`).join('')}</mtr>`);
      return `<mtable>${rows.join('')}</mtable>`;
    }
    case 'eqArr': {
      const rows = kids(n).filter((c) => isM(c, 'e')).map((e) => `<mtr><mtd>${group(e)}</mtd></mtr>`);
      return `<mtable>${rows.join('')}</mtable>`;
    }
    case 'oMath': return `<math>${kids(n).map(conv).join('')}</math>`;
    case 'oMathPara': return kids(n).filter((c) => isM(c, 'oMath')).map((c) => `<math display="block">${kids(c).map(conv).join('')}</math>`).join('');
    default:
      // свойства (…Pr, ctrlPr) не показываются; прочее — по содержимому
      if (/Pr$/.test(n.localName)) return '';
      return kids(n).map(conv).join('');
  }
}

// OMML (строка с m:oMath или m:oMathPara) → MathML; не разобрать — null
export function ommlToMathml(xml) {
  try {
    const doc = new DOMParser({ onError: () => {} }).parseFromString(
      `<x xmlns:m="${M_NS}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</x>`, 'text/xml');
    const root = doc.documentElement;
    const out = kids(root).map(conv).join('');
    return out.includes('<math') ? out : null;
  } catch {
    return null;
  }
}

// MathML → DOM-узел для показа (кеш по XML: renderHTML вызывается при каждой перерисовке абзаца)
const cache = new Map();
export function mathDom(xml) {
  if (typeof document === 'undefined') return null;
  let ml = cache.get(xml);
  if (ml === undefined) {
    ml = ommlToMathml(xml);
    if (cache.size > 500) cache.clear();
    cache.set(xml, ml);
  }
  if (!ml) return null;
  const span = document.createElement('span');
  span.className = 'math-body';
  span.innerHTML = ml;
  return span;
}
