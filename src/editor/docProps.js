// Свойства документа (Файл → Свойства; поля AUTHOR, TITLE, SUBJECT, KEYWORDS, COMMENTS…) ↔ docProps/core.xml.
// doc.attrs.props = { title, subject, author, keywords, comments, lastModifiedBy, created, modified } (строки; даты ISO),
// propsChanged — менялись в Пере (при сохранении core.xml исходного файла правятся только эти элементы).
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

// свойство → элемент core.xml (порядок — как пишет Word)
export const CORE = [
  ['title', 'dc:title'], ['subject', 'dc:subject'], ['author', 'dc:creator'], ['keywords', 'cp:keywords'],
  ['comments', 'dc:description'], ['lastModifiedBy', 'cp:lastModifiedBy'], ['created', 'dcterms:created'], ['modified', 'dcterms:modified'],
];
export const PROP_LABELS = { title: 'Название', subject: 'Тема', author: 'Автор', keywords: 'Ключевые слова', comments: 'Примечания' };
const isDate = (k) => k === 'created' || k === 'modified';

export function readCoreProps(xml) {
  const out = {};
  if (!xml) return out;
  for (const [k, tag] of CORE) {
    if (k === 'modified') continue; // дата изменения — при каждом сохранении своя; в модель не берём
    const m = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml);
    if (m && m[1].trim()) out[k] = unesc(m[1].trim());
  }
  return out;
}

const element = (k, tag, v) => (isDate(k) ? `<${tag} xsi:type="dcterms:W3CDTF">${esc(v)}</${tag}>` : `<${tag}>${esc(v)}</${tag}>`);

// core.xml: исходный (или пустой) с заменёнными свойствами; пустое значение — элемент убирается
export function patchCoreProps(xml, props, now = new Date()) {
  const stamp = now.toISOString().replace(/\.\d+Z$/, 'Z');
  let s = xml || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>';
  // даты: из свойств, если заданы (документ из файла), иначе — сейчас
  const values = { created: stamp, modified: stamp, ...(props || {}) };
  for (const [k, tag] of CORE) {
    const re = new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`);
    const v = values[k];
    if (v == null) continue; // не задано в Пере — как в файле
    if (re.test(s)) s = s.replace(re, v === '' ? '' : element(k, tag, v));
    else if (v !== '') s = s.replace(/<\/cp:coreProperties>\s*$/, `${element(k, tag, v)}</cp:coreProperties>`);
  }
  return s;
}
