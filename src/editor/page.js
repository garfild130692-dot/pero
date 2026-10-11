// Страница: геометрия из w:sectPr, колонтитулы (модель и разбор XML), изменение параметров страницы.
// Общий модуль для редактора, импорта и экспорта (без DOM — разбор регулярными выражениями и xmldom-узлами).

export const TWIPS_PER_PX = 15;   // 1 px = 1/96 дюйма = 15 твипов
export const TWIPS_PER_CM = 567;

// Параметры страницы по умолчанию (A4, поля как в русском Word) — в твипах
export const DEFAULT_PAGE = { w: 11906, h: 16838, top: 1134, right: 850, bottom: 1134, left: 1701, header: 708, footer: 708 };

// [имя, ширина, высота (твипы, книжная), код размера бумаги Word (w:code)]
export const PAPER = [
  ['A4', 11906, 16838, 9], ['A5', 8391, 11906, 11], ['A3', 16838, 23811, 8], ['A6', 5953, 8391, 70],
  ['B4 (JIS)', 14570, 20636, 12], ['B5 (JIS)', 10318, 14570, 13],
  ['Letter', 12240, 15840, 1], ['Legal', 12240, 20160, 5], ['Executive', 10440, 15120, 7], ['Tabloid', 15840, 24480, 3],
];
// формат бумаги по размеру (любая ориентация, допуск ~0,5 мм)
export const paperOf = (w, h) => PAPER.find(([, pw, ph]) => (Math.abs(pw - w) < 30 && Math.abs(ph - h) < 30) || (Math.abs(ph - w) < 30 && Math.abs(pw - h) < 30)) || null;

// Предустановки полей (твипы): верх, право, низ, лево
export const MARGIN_PRESETS = [
  ['gost', 'Обычные (ГОСТ)', { top: 1134, right: 850, bottom: 1134, left: 1701 }],
  ['narrow', 'Узкие', { top: 720, right: 720, bottom: 720, left: 720 }],
  ['moderate', 'Средние', { top: 1440, right: 1080, bottom: 1440, left: 1080 }],
  ['wide', 'Широкие', { top: 1440, right: 2880, bottom: 1440, left: 2880 }],
  ['mirror', 'Зеркальные', { top: 1440, right: 1440, bottom: 1440, left: 1800 }, true], // внутри 3,18 / снаружи 2,54 (как Word)
];

const attr = (tag, name) => {
  const m = new RegExp(`\\bw:${name}="(-?\\d+)"`).exec(tag || '');
  return m ? +m[1] : null;
};

// Геометрия страницы из XML <w:sectPr> (или умолчание). Значения — в твипах.
export function parseSectPr(xml) {
  const g = { ...DEFAULT_PAGE, orient: 'portrait', titlePg: false, gutter: 0, marLeft: DEFAULT_PAGE.left };
  if (!xml) return g;
  const sz = (/<w:pgSz\b[^>]*>/.exec(xml) || [])[0];
  const mar = (/<w:pgMar\b[^>]*>/.exec(xml) || [])[0];
  if (sz) {
    g.w = attr(sz, 'w') ?? g.w;
    g.h = attr(sz, 'h') ?? g.h;
    if (/w:orient="landscape"/.test(sz)) g.orient = 'landscape';
  }
  if (mar) for (const k of ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter']) g[k] = attr(mar, k) ?? g[k];
  // переплёт (слева) — часть левого поля для раскладки; marLeft — поле без переплёта (как в диалоге Word)
  g.gutter = Math.max(0, g.gutter || 0);
  g.marLeft = g.left;
  g.left += g.gutter;
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

// Изменить размер/ориентацию/поля в XML sectPr (остальное — колонтитулы, колонки, сетка — как было).
// left — поле вместе с переплётом (как в раскладке); gutter — переплёт (нет — прежний); code — код бумаги (w:code; при
// смене размера без кода прежний код убирается — иначе принтер взял бы старый формат)
export function patchSectPr(xml, { w, h, orient, top, right, bottom, left, header, footer, gutter, code } = {}) {
  let out = xml || `<w:sectPr><w:pgSz w:w="${DEFAULT_PAGE.w}" w:h="${DEFAULT_PAGE.h}"/><w:pgMar w:top="${DEFAULT_PAGE.top}" w:right="${DEFAULT_PAGE.right}" w:bottom="${DEFAULT_PAGE.bottom}" w:left="${DEFAULT_PAGE.left}" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
  const setAttr = (tag, name, v) => {
    if (v == null) return tag;
    const re = new RegExp(`\\bw:${name}="[^"]*"`);
    return re.test(tag) ? tag.replace(re, `w:${name}="${v}"`) : tag.replace(/\s*\/?>$/, (end) => ` w:${name}="${v}"${end}`);
  };
  const gut = gutter ?? Math.max(0, attr((/<w:pgMar\b[^>]*>/.exec(out) || [])[0], 'gutter') || 0);
  if (left != null) left = Math.max(0, left - gut);
  if (gutter != null) gutter = Math.max(0, gutter);
  out = out.replace(/<w:pgSz\b[^>]*\/>/, (tag) => {
    const sizeChanged = (w != null && w !== attr(tag, 'w')) || (h != null && h !== attr(tag, 'h'));
    let t = setAttr(setAttr(tag, 'w', w), 'h', h);
    if (code != null) t = setAttr(t, 'code', code);
    else if (sizeChanged) t = t.replace(/\s*w:code="[^"]*"/, '');
    if (orient === 'landscape') t = setAttr(t, 'orient', 'landscape');
    if (orient === 'portrait') t = t.replace(/\s*w:orient="[^"]*"/, '');
    return t;
  });
  out = out.replace(/<w:pgMar\b[^>]*\/>/, (tag) => {
    let t = tag;
    for (const [k, v] of Object.entries({ top, right, bottom, left, header, footer, gutter })) t = setAttr(t, k, v);
    return t;
  });
  return out;
}

// ---------- Колонтитулы ----------
// Модель: { header, footer: Para[] | null — «по умолчанию» (нечётные/все страницы),
//           first: { header, footer } | null — первая страница раздела (если у раздела w:titlePg),
//           even: { header, footer } | null — чётные страницы (если в settings.xml w:evenAndOddHeaders),
//           titlePg: bool, evenOdd: bool — «особый для первой» (значение из Пера, см. own) и «разные для чётных»,
//           own: { header, footer, first, even, titlePg, evenOdd } — что изменено в Пере }
// Para: { text, align: 'left'|'center'|'right' } — в тексте поля {PAGE} (номер страницы) и {NUMPAGES} (всего страниц).
// own — задано в Пере: при сохранении пишутся свои части и ссылки на них во всех разделах (titlePg — во всех разделах,
// evenOdd — в settings.xml). Колонтитулы из файла (own = false) только показываются; их части сохраняются как были.
export const EMPTY_HF = {
  header: null, footer: null, first: null, even: null, titlePg: false, evenOdd: false,
  own: { header: false, footer: false, first: false, even: false, titlePg: false, evenOdd: false },
};

// Какой колонтитул на странице (правила Word): первая страница раздела с «особым для первой» — first (не задан — пусто),
// чётная при «разных для чётных и нечётных» — even (не задан — пусто), иначе — «по умолчанию».
// page — номер страницы (с 1), firstOfSection — страница первая в разделе, sectTitlePg — w:titlePg раздела в файле.
export function hfFor(hf, kind, { page, firstOfSection, sectTitlePg }) {
  const h = hf || EMPTY_HF;
  const own = h.own || {};
  const titlePg = own.titlePg ? !!h.titlePg : !!sectTitlePg;
  if (firstOfSection && titlePg) return (h.first && h.first[kind]) || [];
  if (h.evenOdd && page % 2 === 0) return (h.even && h.even[kind]) || [];
  return h[kind] || [];
}

// Включить/выключить «особый колонтитул для первой страницы» в XML sectPr
export function setTitlePg(sectXml, on) {
  let out = sectXml.replace(/<w:titlePg\b[^>]*\/>/g, '');
  if (on) out = /<w:docGrid\b/.test(out) ? out.replace(/<w:docGrid\b/, '<w:titlePg/><w:docGrid') : out.replace(/<\/w:sectPr>/, '<w:titlePg/></w:sectPr>');
  return out;
}

// Разделы документа по порядку: XML sectPr (абзацы с разрывом раздела, затем последний — у документа)
export function sectionList(doc) {
  const list = [];
  doc.forEach((n) => { if (n.attrs && n.attrs.sectPr) list.push(n.attrs.sectPr); });
  list.push(doc.attrs.sectPr || null);
  return list;
}

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

// Заменить ссылки на колонтитулы типа type (default | first | even) в XML sectPr; ids: { header?, footer? } (null — убрать)
export function setHfRefs(sectXml, ids, type = 'default') {
  let out = sectXml;
  for (const kind of ['header', 'footer']) {
    if (!(kind in ids)) continue;
    out = out.replace(new RegExp(`<w:${kind}Reference\\b[^>]*w:type="${type}"[^>]*/>`, 'g'), '');
    if (ids[kind]) out = out.replace(/(<w:sectPr\b[^>]*>)/, `$1<w:${kind}Reference w:type="${type}" r:id="${ids[kind]}"/>`);
  }
  // ссылки на колонтитулы идут в начале sectPr: header* затем footer* — Word терпит любой порядок внутри группы
  return out;
}

// ---------- Разделы с разной геометрией листа ----------
// Рамка экрана для разделов с разными размерами/полями (книжная + альбомная и т.п.): лист на экране шириной с самый
// широкий раздел, узкие листы — по центру; область текста редактора — от самого левого до самого правого края текста;
// блоки раздела сдвигаются полями --sec-l/--sec-r (px). Один раздел или одинаковые — сдвигов нет.
export function sectionFrame(sects) {
  const list = sects.map((xml) => pagePx(parseSectPr(xml)));
  const maxW = Math.max(...list.map((p) => p.width));
  const off = (p) => (maxW - p.width) / 2;
  const left = Math.min(...list.map((p) => off(p) + p.left));
  const right = Math.max(...list.map((p) => off(p) + p.width - p.right)); // правый край текста от левого края листа
  const mixed = list.some((p) => p.width !== list[0].width || p.height !== list[0].height || p.left !== list[0].left || p.right !== list[0].right);
  return {
    mixed,
    width: maxW,
    padLeft: left,
    padRight: maxW - right,
    sections: list.map((p) => ({ ...p, offset: off(p), secL: off(p) + p.left - left, secR: right - (off(p) + p.width - p.right) })),
  };
}

// Раздел под курсором: { pos — абзац с его w:sectPr (null — последний раздел, у документа), sectPr }
export function sectionAt(state) {
  const top = state.selection.$from.index(0);
  let found = null;
  state.doc.forEach((n, pos, idx) => { if (!found && idx >= top && n.attrs && n.attrs.sectPr) found = { pos, sectPr: n.attrs.sectPr }; });
  return found || { pos: null, sectPr: state.doc.attrs.sectPr };
}

// Флажок в settings.xml (w:mirrorMargins и т.п.) на месте по схеме: after — элементы, после которых он стоит (ближайший первым)
export function patchSettingsFlag(xml, name, on, after) {
  const s = String(xml).replace(new RegExp(`<w:${name}\\b[^>]*/>`, 'g'), '');
  if (!on) return s;
  let at = -1;
  for (const tag of after) {
    const m = new RegExp(`<w:${tag}\\b[^>]*(?:/>|>[\\s\\S]*?</w:${tag}>)`).exec(s);
    if (m) { at = m.index + m[0].length; break; }
  }
  if (at < 0) { const m = /<w:settings\b[^>]*>/.exec(s); at = m.index + m[0].length; }
  return `${s.slice(0, at)}<w:${name}/>${s.slice(at)}`;
}
// предшественники w:mirrorMargins в CT_Settings (ближайший первым)
export const MIRROR_AFTER = ['saveFormsData', 'saveSubsetFonts', 'embedSystemFonts', 'embedTrueTypeFonts', 'printFormsData',
  'printFractionalCharacterWidth', 'printPostScriptOverText', 'displayBackgroundShape', 'doNotDisplayPageBoundaries',
  'removeDateAndTime', 'removePersonalInformation', 'zoom', 'view', 'writeProtection'];
