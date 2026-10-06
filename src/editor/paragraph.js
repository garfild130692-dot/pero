// Абзац: единицы, умолчания и преобразования — общие для редактора, импорта и экспорта.
// Внутри всё храним в пунктах (pt). OOXML — в твипах (1 pt = 20 твипов), множитель интервала — в 1/240 строки.

export const TWIPS_PER_PT = 20;
export const TWIPS_PER_CM = 567;            // как считает Word: 1 см = 567 твипов
export const PT_PER_CM = TWIPS_PER_CM / TWIPS_PER_PT;
export const INDENT_STEP = 708 / TWIPS_PER_PT; // шаг «Увеличить отступ» = стандартная позиция табуляции 1,25 см

// Атрибуты абзаца (null — не задано, действует умолчание документа/стиля)
//   spaceBefore, spaceAfter — pt; lineRule: 'auto' (множитель) | 'exact' | 'atLeast'; lineSpacing — множитель или pt;
//   indentLeft, indentRight — pt; indentFirst — pt (отрицательный = выступ).
export const PARA_KEYS = ['spaceBefore', 'spaceAfter', 'lineRule', 'lineSpacing', 'indentLeft', 'indentRight', 'indentFirst'];

// Умолчания нового документа (как пишет наш экспорт и Word 2013+: после 8 pt, множитель 1,08)
export const NEW_DOC_PARA = {
  spaceBefore: 0, spaceAfter: 8, lineRule: 'auto', lineSpacing: 259 / 240, indentLeft: 0, indentRight: 0, indentFirst: 0,
};
// Если в файле не задано ничего — так считает Word
export const WORD_FALLBACK_PARA = {
  spaceBefore: 0, spaceAfter: 0, lineRule: 'auto', lineSpacing: 1, indentLeft: 0, indentRight: 0, indentFirst: 0,
};

// Интервалы заголовков (совпадают со стилями Heading1–6 в export.js)
export const HEADING_PARA = (level) => ({
  spaceBefore: level === 1 ? 12 : 2, spaceAfter: 0, lineRule: 'auto', lineSpacing: 259 / 240,
});

// База, от которой отсчитывается прямое форматирование абзаца: умолчания документа,
// у заголовка — его интервалы, у элемента списка — нулевые интервалы и отступы (их задаёт сам список)
export function paraBase(docPara, { headingLevel = 0, inList = false } = {}) {
  const base = { ...NEW_DOC_PARA, ...(docPara || {}) };
  if (headingLevel) Object.assign(base, HEADING_PARA(headingLevel));
  if (inList) Object.assign(base, { indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0 });
  return base;
}

// Оставить только отличия от базы; межстрочный интервал (правило + значение) сравнивается парой
export function paraDiff(values, base) {
  const out = {};
  for (const k of PARA_KEYS) out[k] = null;
  for (const k of ['spaceBefore', 'spaceAfter', 'indentLeft', 'indentRight', 'indentFirst']) {
    if (values[k] != null && Math.abs(values[k] - base[k]) > 1e-9) out[k] = values[k];
  }
  if (values.lineSpacing != null) {
    const rule = values.lineRule || 'auto';
    if (rule !== base.lineRule || Math.abs(values.lineSpacing - base.lineSpacing) > 1e-9) {
      out.lineRule = rule;
      out.lineSpacing = values.lineSpacing;
    }
  }
  return out;
}

export const twipsToPt = (v) => (v === null || v === undefined || v === '' || isNaN(+v) ? null : +v / TWIPS_PER_PT);
export const ptToTwips = (v) => Math.round(v * TWIPS_PER_PT);
export const cmToPt = (cm) => Math.round(cm * TWIPS_PER_CM) / TWIPS_PER_PT; // округляем до твипа — round-trip без дрейфа
export const ptToCm = (pt) => pt / PT_PER_CM;

// Нормализация отдельных значений
export const LINE_RULES = ['auto', 'exact', 'atLeast'];
export function normalizeLine(rule, value) {
  const r = LINE_RULES.includes(rule) ? rule : 'auto';
  const v = +value;
  if (!isFinite(v) || v <= 0) return null;
  return { lineRule: r, lineSpacing: r === 'auto' ? Math.round(v * 240) / 240 : Math.round(v * 20) / 20 };
}

// Подпись для выпадающего списка «Интервал»: 1,0 / 1,15 / 1,5 …
export const LINE_PRESETS = [1, 1.15, 1.5, 2, 2.5, 3];
export const fmtNum = (n, digits = 2) => String(+n.toFixed(digits)).replace('.', ',');
