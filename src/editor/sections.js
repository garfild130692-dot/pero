// Разделы с разной геометрией листа (книжная + альбомная, разные поля): блоки раздела получают сдвиг полей
// относительно общей области текста (--sec-l / --sec-r, px — см. sectionFrame в page.js) и имя страницы для печати
// (CSS page: pero-sN → @page pero-sN { size, margin } — см. sectionPrintCss).
// Один раздел или все одинаковые — декораций нет, вид как раньше.
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { pagePx, parseSectPr, sectionFrame, sectionList } from './page.js';

export const sectionsKey = new PluginKey('sections');

function sectionDecorations(doc) {
  const frame = sectionFrame(sectionList(doc));
  if (!frame.mixed) return { frame, deco: DecorationSet.empty };
  const decos = [];
  let s = 0;
  doc.forEach((node, pos) => {
    const sec = frame.sections[Math.min(s, frame.sections.length - 1)];
    decos.push(Decoration.node(pos, pos + node.nodeSize, {
      style: `--sec-l: ${sec.secL}px; --sec-r: ${sec.secR}px; page: pero-s${s}`,
      'data-section': String(s),
    }));
    if (node.attrs && node.attrs.sectPr) s++;
  });
  return { frame, deco: DecorationSet.create(doc, decos) };
}

export const Sections = Extension.create({
  name: 'sections',
  addProseMirrorPlugins() {
    return [new Plugin({
      key: sectionsKey,
      state: {
        init: (_c, state) => sectionDecorations(state.doc),
        apply: (tr, old) => (tr.docChanged ? sectionDecorations(tr.doc) : old),
      },
      props: { decorations: (state) => sectionsKey.getState(state).deco },
    })];
  },
});

// Печать разделов с разной геометрией: свой @page на каждый раздел (имена совпадают с декорациями)
export function sectionPrintCss(doc, mirror = !!doc.attrs.mirrorMargins) {
  const sects = sectionList(doc);
  const frame = sectionFrame(sects);
  if (!frame.mixed) return '';
  return sects.map((xml, i) => {
    const p = pagePx(parseSectPr(xml));
    const m = `@page pero-s${i} { size: ${p.width}px ${p.height}px; margin: ${p.top}px ${p.right}px ${p.bottom}px ${p.left}px; }`;
    return mirror ? `${m} @page pero-s${i}:left { margin-left: ${p.right}px; margin-right: ${p.left}px; }` : m;
  }).join(' ');
}
