// Страница: геометрия из w:sectPr, колонтитулы (модель и разбор XML), изменение параметров страницы.
// Общий модуль для редактора, импорта и экспорта (без DOM — разбор регулярными выражениями и xmldom-узлами).

export const TWIPS_PER_PX = 15;   // 1 px = 1/96 дюйма = 15 твипов
export const TWIPS_PER_CM = 567;

// Параметры страницы по умолчанию (A4, поля как в русском Word) — в твипах
export const DEFAULT_PAGE = { w: 11906, h: 16838, top: 1134, right: 850, bottom: 1134, left: 1701, header: 708, footer: 708 };

export const PAPER = [
  ['A4', 11906, 16838], ['A5', 8391, 11906], ['A3', 16838, 23811], ['Letter', 12240, 15840], ['Legal', 12240, 20160],
];

// Предустановки полей (твипы): верх, право, низ, лево
export const MARGIN_PRESETS = [
  ['gost', 'Обычные (ГОСТ)', { top: 1134, right: 850, bottom: 1134, left: 1701 }],
  ['narrow', 'Узкие', { top: 720, right: 720, bottom: 720, left: 720 }],
  ['moderate', 'Средние', { top: 1440, right: 1080, bottom: 1440, left: 1080 }],
  ['wide', 'Широкие', { top: 1440, right: 2880, bottom: 1440, left: 2880 }],
];

const attr = (tag, name) => {
  const m = new RegExp(`\\bw:${name}="(-?\\d+)"`).exec(tag || '');
  return m ? +m[1] : null;
};

// Геометрия страницы из XML <w:sectPr> (или умолчание). Значения — в твипах.
export function parseSectPr(xml) {
  const g = { ...DEFAULT_PAGE, orient: 'portrait', titlePg: false };
  if (!xml) return g;
  const sz = (/<w:pgSz\b[^>]*>/.exec(xml) || [])[0];
  const mar = (/<w:pgMar\b[^>]*>/.exec(xml) || [])[0];
  if (sz) {
    g.w = attr(sz, 'w') ?? g.w;
    g.h = attr(sz, 'h') ?? g.h;
    if (/w:orient="landscape"/.test(sz)) g.orient = 'landscape';
  }
  if (mar) for (const k of ['top', 'right', 'bottom', 'left', 'header', 'footer']) g[k] = attr(mar, k) ?? g[k];
  // в Word поля сверху/снизу могут быть отрицательными («не сдвигать текст») — берём модуль
  g.top = Math.abs(g.top);
  g.bottom = Math.abs(g.bottom);
  g.titlePg = /<w:titlePg\b(?![^>]*w:val="(?:0|false)")/.test(xml);
  return g;
}

// Геометрия в px для экрана
export function pagePx(g) {
  const px = (t) => t / TWIPS_PER_PX;
  return {
    width: px(g.w), height: px(g.h), top: px(g.top), right: px(g.right), bottom: px(g.bottom), left: px(g.left),
    header: px(g.header), footer: px(g.footer),
  };
}

// Изменить размер/ориентацию/поля в XML sectPr (остальное — колонтитулы, колонки, сетка — как было)
export function patchSectPr(xml, { w, h, orient, top, right, bottom, left, header, footer } = {}) {
  let out = xml || `<w:sectPr><w:pgSz w:w="${DEFAULT_PAGE.w}" w:h="${DEFAULT_PAGE.h}"/><w:pgMar w:top="${DEFAULT_PAGE.top}" w:right="${DEFAULT_PAGE.right}" w:bottom="${DEFAULT_PAGE.bottom}" w:left="${DEFAULT_PAGE.left}" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
  const setAttr = (tag, name, v) => {
    if (v == null) return tag;
    const re = new RegExp(`\\bw:${name}="[^"]*"`);
    return re.test(tag) ? tag.replace(re, `w:${name}="${v}"`) : tag.replace(/\s*\/?>$/, (end) => ` w:${name}="${v}"${end}`);
  };
  out = out.replace(/<w:pgSz\b[^>]*\/>/, (tag) => {
    let t = setAttr(setAttr(tag, 'w', w), 'h', h);
    if (orient === 'landscape') t = setAttr(t, 'orient', 'landscape');
    if (orient === 'portrait') t = t.replace(/\s*w:orient="[^"]*"/, '');
    return t;
  });
  out = out.replace(/<w:pgMar\b[^>]*\/>/, (tag) => {
    let t = tag;
    for (const [k, v] of Object.entries({ top, right, bottom, left, header, footer })) t = setAttr(t, k, v);
    return t;
  });
  return out;
}

// ---------- Колонтитулы ----------
// Модель: { header: Para[] | null, footer: Para[] | null, own: { header: bool, footer: bool } }
// Para: { text, align: 'left'|'center'|'right' } — в тексте поля {PAGE} (номер страницы) и {NUMPAGES} (всего страниц).
// own — колонтитул задан в Пере: при сохранении пишется своя часть и ссылки на неё во всех разделах.
// Колонтитулы из файла (own = false) только показываются; их части сохраняются из пакета без изменений.
export const EMPTY_HF = { header: null, footer: null, own: { header: false, footer: false } };

// Текст колонтитула для страницы n из total
export const hfText = (text, n, total) => String(text || '').replace(/\{PAGE\}/g, String(n)).replace(/\{NUMPAGES\}/g, String(total));

const JC = { center: 'center', right: 'right', end: 'right', both: 'left', left: 'left', start: 'left' };

// Разбор части header*.xml/footer*.xml (xmldom-документ) → Para[]; поля PAGE/NUMPAGES → {PAGE}/{NUMPAGES}
export function readHfPart(doc) {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const ln = (n) => n.localName || String(n.nodeName).split(':').pop();
  const kids = (el) => { const out = []; for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c); return out; };
  const fieldToken = (instr) => {
    const name = String(instr || '').trim().split(/\s+/)[0].toUpperCase();
    return name === 'PAGE' ? '{PAGE}' : name === 'NUMPAGES' || name === 'SECTIONPAGES' ? '{NUMPAGES}' : null;
  };
  const paras = [];
  const allP = Array.from(doc.getElementsByTagNameNS(W, 'p'));
  for (const p of allP) {
    if (p.parentNode && ln(p.parentNode) === 'txbxContent') continue; // надписи (часто — номер страницы в рамке) — ниже
    let text = '';
    let field = null; // { instr, inResult }
    const walk = (el) => {
      for (const c of kids(el)) {
        const name = ln(c);
        if (name === 'fldSimple') {
          const tok = fieldToken(c.getAttributeNS(W, 'instr') || c.getAttribute('w:instr'));
          if (tok) text += tok; else walk(c);
        } else if (name === 'fldChar') {
          const t = c.getAttributeNS(W, 'fldCharType') || c.getAttribute('w:fldCharType');
          if (t === 'begin') field = { instr: '', inResult: false };
          else if (t === 'separate' && field) {
            field.inResult = true;
            const tok = fieldToken(field.instr);
            if (tok) { text += tok; field.skip = true; }
          } else if (t === 'end') field = null;
        } else if (name === 'instrText') {
          if (field && !field.inResult) field.instr += c.textContent;
        } else if (name === 't') {
          if (!(field && field.skip)) text += c.textContent;
        } else if (name === 'tab') text += '\t';
        else if (name === 'pPr' || name === 'rPr') continue;
        else walk(c);
      }
    };
    walk(p);
    const pPr = kids(p).find((c) => ln(c) === 'pPr');
    const jc = pPr && kids(pPr).find((c) => ln(c) === 'jc');
    const align = JC[jc && (jc.getAttributeNS(W, 'val') || jc.getAttribute('w:val'))] || 'left';
    paras.push({ text, align });
  }
  // номер страницы часто стоит в надписи (txbxContent) — достаём хотя бы поле
  if (!paras.some((p) => /\{PAGE\}/.test(p.text))) {
    for (const p of allP) {
      if (!(p.parentNode && ln(p.parentNode) === 'txbxContent')) continue;
      const xml = String(p);
      if (/PAGE/.test(p.textContent) || /w:instr="\s*PAGE/.test(xml)) { paras.push({ text: '{PAGE}', align: 'right' }); break; }
    }
  }
  // убрать пустые абзацы по краям
  while (paras.length && !paras[0].text.trim()) paras.shift();
  while (paras.length && !paras[paras.length - 1].text.trim()) paras.pop();
  return paras;
}

// Para[] → XML части колонтитула (w:hdr / w:ftr); поля — w:fldSimple
export function writeHfPart(kind, paras, ns) {
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  const run = (t) => (t ? `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>` : '');
  const body = (paras.length ? paras : [{ text: '', align: 'left' }]).map((p) => {
    const parts = String(p.text).split(/(\{PAGE\}|\{NUMPAGES\})/);
    const inner = parts.map((x) => (x === '{PAGE}' ? '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>'
      : x === '{NUMPAGES}' ? '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple>' : run(x))).join('');
    const jc = p.align === 'center' || p.align === 'right' ? `<w:jc w:val="${p.align}"/>` : '';
    return `<w:p><w:pPr><w:pStyle w:val="${kind === 'hdr' ? 'Header' : 'Footer'}"/>${jc}</w:pPr>${inner}</w:p>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${kind} ${ns}>${body}</w:${kind}>`;
}

// Заменить ссылки на колонтитулы «по умолчанию» в XML sectPr; ids: { header?, footer? } (null — убрать)
export function setHfRefs(sectXml, ids) {
  let out = sectXml;
  for (const kind of ['header', 'footer']) {
    if (!(kind in ids)) continue;
    out = out.replace(new RegExp(`<w:${kind}Reference\\b[^>]*w:type="default"[^>]*/>`, 'g'), '');
    if (ids[kind]) out = out.replace(/(<w:sectPr\b[^>]*>)/, `$1<w:${kind}Reference w:type="default" r:id="${ids[kind]}"/>`);
  }
  // ссылки на колонтитулы идут в начале sectPr: header* затем footer* — Word терпит любой порядок внутри группы
  return out;
}
