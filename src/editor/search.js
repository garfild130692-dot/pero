// Поиск и замена как в Word: подсветка всех совпадений, переход, замена текущего и всех; специальные знаки
// (^p — конец абзаца, ^t — табуляция, ^l — разрыв строки, ^m — разрыв страницы, ^s — неразрывный пробел, ^~ — неразрывный
// дефис, ^- — мягкий перенос, ^? — любой знак, ^# — любая цифра, ^$ — любая буква, ^w — пробелы, ^^ — «^»),
// подстановочные знаки Word (? * [ ] [!] {n,m} @ < > ( ) \1), поиск и замена формата (полужирный, курсив, …, стиль).
import { Extension } from '@tiptap/core';
import { Fragment, Slice } from '@tiptap/pm/model';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

export const searchKey = new PluginKey('search');
const OBJ = '￼'; // нетекстовые узлы (рисунки, значки) — совпадения через них не идут
const BLOCK = '\n';   // граница абзацев (^p)
const LINE = '\v';    // разрыв строки (^l)
const PAGE = '\f';    // разрыв страницы (^m)
const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');

// Специальные знаки Word (^x) в обычном поиске → фрагмент регулярного выражения
const SPECIAL_FIND = {
  p: BLOCK, t: '\t', l: LINE, m: PAGE, s: ' ', '~': '‑', '-': '­', '^': '\\^',
  '?': `[^${BLOCK}${OBJ}]`, '#': '\\p{N}', $: '\\p{L}', w: '[ \\t\\u00a0]+',
};
// В замене
const SPECIAL_REPL = { p: BLOCK, t: '\t', l: LINE, s: ' ', '~': '‑', '-': '­', '^': '^' };

// Обычный поиск: текст + ^коды
function plainPattern(q) {
  let out = '';
  for (let i = 0; i < q.length; i++) {
    const c = q[i];
    if (c === '^' && i + 1 < q.length && SPECIAL_FIND[q[i + 1]] != null) { const x = SPECIAL_FIND[q[i + 1]]; out += x === BLOCK || x === LINE || x === PAGE ? reEsc(x) : x; i++; continue; }
    out += reEsc(c);
  }
  return out;
}

// Подстановочные знаки Word → регулярное выражение JS (в пределах абзаца; регистр всегда учитывается — как в Word)
export function wildcardPattern(q) {
  let out = '';
  for (let i = 0; i < q.length; i++) {
    const c = q[i];
    if (c === '\\' && i + 1 < q.length) { out += reEsc(q[++i]); continue; }
    if (c === '^' && i + 1 < q.length && SPECIAL_FIND[q[i + 1]] != null) { const x = SPECIAL_FIND[q[i + 1]]; out += x === BLOCK ? reEsc(x) : x; i++; continue; }
    if (c === '?') { out += `[^${BLOCK}]`; continue; }
    if (c === '*') { out += `[^${BLOCK}]*?`; continue; }
    if (c === '@') { out += '+'; continue; }
    if (c === '<') { out += '(?<![\\p{L}\\p{N}_])'; continue; }
    if (c === '>') { out += '(?![\\p{L}\\p{N}_])'; continue; }
    if (c === '(' || c === ')') { out += c; continue; }
    if (c === '{') {
      const m = /^\{(\d*)(,?)(\d*)\}/.exec(q.slice(i));
      if (m) { out += `{${m[1] || 0}${m[2]}${m[3]}}`; i += m[0].length - 1; continue; }
    }
    if (c === '[') {
      const end = q.indexOf(']', i + 1);
      if (end > i) {
        let body = q.slice(i + 1, end);
        const neg = body.startsWith('!');
        if (neg) body = body.slice(1);
        out += `[${neg ? '^' : ''}${body.replace(/[\\\]^]/g, '\\$&')}]`;
        i = end;
        continue;
      }
    }
    out += reEsc(c);
  }
  return out;
}

// Регулярное выражение поиска; «слово целиком» — по буквам/цифрам Юникода (\b в JS кириллицу не понимает)
export function buildRegex(query, { caseSensitive = false, wholeWord = false, wildcards = false } = {}) {
  if (!query) return null;
  let src;
  try {
    src = wildcards ? wildcardPattern(query) : plainPattern(query);
    if (wholeWord && !wildcards) src = `(?<![\\p{L}\\p{N}_])${src}(?![\\p{L}\\p{N}_])`;
    return new RegExp(src, `gu${caseSensitive || wildcards ? '' : 'i'}`);
  } catch { return null; } // неверный шаблон подстановочных знаков — нет совпадений
}

// ---------- Формат ----------
// Условия формата: { bold, italic, underline, strike, highlight, superscript, subscript: true|false } и styleId (абзаца)
export const FORMAT_FLAGS = [['bold', 'Полужирный'], ['italic', 'Курсив'], ['underline', 'Подчёркнутый'], ['strike', 'Зачёркнутый'],
  ['highlight', 'Выделение цветом'], ['superscript', 'Надстрочный'], ['subscript', 'Подстрочный']];
const hasFormat = (f) => !!f && (FORMAT_FLAGS.some(([k]) => f[k] != null) || f.styleId !== undefined);
const nodeMatches = (marks, parent, f) => FORMAT_FLAGS.every(([k]) => f[k] == null || marks.some((m) => m.type.name === k) === f[k])
  && (f.styleId === undefined || (parent.type.name === 'paragraph' ? parent.attrs.styleId || null : `h${parent.attrs.level}`) === f.styleId);

// Текст документа для поиска: абзацы через \n; у каждого знака — диапазон позиций [from, to)
function docText(doc) {
  let text = '';
  const from = []; const to = []; const fmt = []; // fmt: [marks, parent] у знака
  let prevEnd = null;
  doc.descendants((block, blockPos) => {
    if (!block.isTextblock) return true;
    if (prevEnd != null) { text += BLOCK; from.push(prevEnd); to.push(blockPos + 1); fmt.push(null); }
    block.forEach((child, off) => {
      const pos = blockPos + 1 + off;
      if (child.isText) {
        for (let i = 0; i < child.text.length; i++) { from.push(pos + i); to.push(pos + i + 1); fmt.push([child.marks, block]); }
        text += child.text;
      } else {
        const name = child.type.name;
        text += name === 'tab' ? '\t' : name === 'hardBreak' ? LINE : name === 'opaque' && child.attrs.kind === 'pageBreak' ? PAGE : OBJ;
        from.push(pos); to.push(pos + child.nodeSize); fmt.push(name === 'tab' ? [child.marks, block] : null);
      }
    });
    prevEnd = blockPos + block.nodeSize - 1;
    return false;
  });
  return { text, from, to, fmt };
}

// Все совпадения в документе: [{ from, to, m }] (позиции документа; m — совпадение регулярного выражения для замены \1…)
export function findAll(doc, query, opts = {}) {
  const f = opts.format;
  const { text, from, to, fmt } = docText(doc);
  const out = [];
  const okFmt = (a, b) => {
    if (!hasFormat(f)) return true;
    for (let i = a; i < b; i++) {
      const x = fmt[i];
      if (x && !nodeMatches(x[0], x[1], f)) return false;
    }
    return true;
  };
  if (!query) {
    // только формат: подряд идущие знаки абзаца с нужным оформлением
    if (!hasFormat(f)) return [];
    let start = -1;
    for (let i = 0; i <= text.length; i++) {
      const good = i < text.length && fmt[i] && text[i] !== OBJ && nodeMatches(fmt[i][0], fmt[i][1], f);
      if (good && start < 0) start = i;
      if (!good && start >= 0) { out.push({ from: from[start], to: to[i - 1], m: [text.slice(start, i)] }); start = -1; }
    }
    return out;
  }
  const re = buildRegex(query, opts);
  if (!re) return [];
  re.lastIndex = 0;
  let m;
  while ((m = re.exec(text))) {
    if (!m[0].length) { re.lastIndex++; continue; }
    const a = m.index; const b = a + m[0].length;
    if (!m[0].includes(OBJ) && okFmt(a, b)) out.push({ from: from[a], to: to[b - 1], m: [...m] });
  }
  return out;
}

// Текст замены: \1…\9 (подстановочные знаки), ^& — найденный текст, ^p ^t ^l ^s… — специальные знаки
export function replacementText(repl, m, wildcards) {
  let out = '';
  const s = String(repl || '');
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (wildcards && c === '\\' && /[1-9]/.test(s[i + 1] || '')) { out += (m && m[+s[i + 1]]) || ''; i++; continue; }
    if (c === '^' && s[i + 1] === '&') { out += m ? m[0] : ''; i++; continue; }
    if (c === '^' && SPECIAL_REPL[s[i + 1]] != null) { out += SPECIAL_REPL[s[i + 1]]; i++; continue; }
    out += c;
  }
  return out;
}

// Узлы замены: текст с метками места; \n — новые абзацы, \v — разрыв строки (табуляцию превращает в узел плагин табуляции)
function replacementSlice(schema, text, marks) {
  const inline = (line) => {
    const nodes = [];
    line.split(LINE).forEach((part, i) => {
      if (i > 0) nodes.push(schema.nodes.hardBreak.create());
      if (part) nodes.push(schema.text(part, marks));
    });
    return nodes;
  };
  const lines = text.split(BLOCK);
  if (lines.length === 1) return new Slice(Fragment.from(inline(lines[0])), 0, 0);
  return new Slice(Fragment.from(lines.map((l) => schema.nodes.paragraph.create(null, inline(l)))), 1, 1);
}

// Метки с заменой формата: { bold: true|false, … } — добавить/убрать
function applyFormat(state, marks, rf) {
  if (!hasFormat(rf)) return marks;
  let out = marks;
  for (const [k] of FORMAT_FLAGS) {
    if (rf[k] == null || !state.schema.marks[k]) continue;
    const type = state.schema.marks[k];
    out = rf[k] ? (k === 'highlight' ? type.create({ color: rf.highlightColor || '#ffff00' }).addToSet(out) : type.create().addToSet(out)) : type.removeFromSet(out);
  }
  return out;
}

const emptyState = { query: '', opts: {}, results: [], index: -1, deco: DecorationSet.empty };

function decorate(doc, results, index) {
  return DecorationSet.create(doc, results.map((r, i) => Decoration.inline(r.from, r.to, { class: i === index ? 'search-hit current' : 'search-hit' })));
}

// индекс ближайшего совпадения не раньше pos
const nearest = (results, pos) => {
  const i = results.findIndex((r) => r.from >= pos);
  return i < 0 ? (results.length ? 0 : -1) : i;
};

// Заменить одно совпадение r в tr (позиции — исходного документа, через mapping)
function replaceOne(state, tr, r, repl, opts) {
  const text = replacementText(repl, r.m, opts.wildcards);
  const marks = applyFormat(state, state.doc.resolve(Math.min(r.from + 1, r.to)).marks(), opts.replaceFormat);
  const a = tr.mapping.map(r.from); const b = tr.mapping.map(r.to);
  if (!text && hasFormat(opts.replaceFormat) && !repl) {
    // пустая замена с форматом — только формат найденного (как в Word)
    for (const [k] of FORMAT_FLAGS) {
      const v = opts.replaceFormat[k];
      if (v == null || !state.schema.marks[k]) continue;
      const type = state.schema.marks[k];
      if (v) tr.addMark(a, b, k === 'highlight' ? type.create({ color: opts.replaceFormat.highlightColor || '#ffff00' }) : type.create()); else tr.removeMark(a, b, type);
    }
    return b;
  }
  tr.replace(a, b, replacementSlice(state.schema, text, marks));
  return tr.mapping.map(r.to, 1);
}

// Заменить все совпадения текущего поиска в tr; возвращает их число
export function replaceAllTr(state, tr, text) {
  const s = searchKey.getState(state);
  if (!s || !s.results.length) return 0;
  for (const r of s.results) replaceOne(state, tr, r, text, s.opts || {});
  tr.setMeta(searchKey, {});
  return s.results.length;
}

export const SearchReplace = Extension.create({
  name: 'searchReplace',

  addProseMirrorPlugins() {
    return [new Plugin({
      key: searchKey,
      state: {
        init: () => emptyState,
        apply(tr, old, _oldState, newState) {
          const meta = tr.getMeta(searchKey);
          if (meta) {
            const query = meta.query ?? old.query;
            const opts = meta.opts ?? old.opts;
            const results = findAll(newState.doc, query, opts);
            let index = meta.index ?? (results.length ? nearest(results, newState.selection.from) : -1);
            if (index >= results.length) index = results.length ? 0 : -1;
            return { query, opts, results, index, deco: decorate(newState.doc, results, index) };
          }
          if (tr.docChanged && (old.query || hasFormat(old.opts.format))) {
            const results = findAll(newState.doc, old.query, old.opts);
            const index = results.length ? Math.min(Math.max(old.index, 0), results.length - 1) : -1;
            return { ...old, results, index, deco: decorate(newState.doc, results, index) };
          }
          return old;
        },
      },
      props: { decorations: (state) => searchKey.getState(state).deco },
    })];
  },

  addCommands() {
    const go = (dir) => ({ state, tr, dispatch, view }) => {
      const s = searchKey.getState(state);
      if (!s.results.length) return false;
      const index = s.index < 0 ? 0 : (s.index + dir + s.results.length) % s.results.length;
      const r = s.results[index];
      if (dispatch) {
        tr.setMeta(searchKey, { index }).setSelection(TextSelection.create(tr.doc, r.from, r.to)).scrollIntoView();
        dispatch(tr);
        // прокрутка к совпадению по центру области
        setTimeout(() => {
          const dom = view && view.domAtPos(r.from).node;
          const el = dom && (dom.nodeType === 1 ? dom : dom.parentElement);
          if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' });
        }, 0);
      }
      return true;
    };
    return {
      setSearch: (query, opts = {}) => ({ tr, dispatch }) => {
        if (dispatch) dispatch(tr.setMeta(searchKey, { query, opts }));
        return true;
      },
      clearSearch: () => ({ tr, dispatch }) => {
        if (dispatch) dispatch(tr.setMeta(searchKey, { query: '', opts: {} }));
        return true;
      },
      findNext: () => go(1),
      findPrev: () => go(-1),
      // Заменить текущее совпадение и перейти к следующему
      replaceCurrent: (text) => ({ state, tr, dispatch }) => {
        const s = searchKey.getState(state);
        const r = s.results[s.index];
        if (!r) return false;
        if (dispatch) {
          const after = replaceOne(state, tr, r, text, s.opts || {});
          tr.setMeta(searchKey, { index: undefined });
          tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(after, tr.doc.content.size))));
          dispatch(tr);
        }
        return true;
      },
      // Заменить все: одной транзакцией (одна отмена)
      replaceAll: (text) => ({ state, tr, dispatch }) => {
        const n = replaceAllTr(state, tr, text);
        if (!n) return false;
        if (dispatch) dispatch(tr);
        return n;
      },
    };
  },
});

export const searchInfo = (state) => {
  const s = searchKey.getState(state);
  return s ? { count: s.results.length, index: s.index } : { count: 0, index: -1 };
};
