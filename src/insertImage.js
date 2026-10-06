// Вставка картинок из файлов (буфер обмена, перетаскивание, «Вставка → Рисунок»).
import * as fileApi from './fileApi.js';
import { mimeOf, prepareImage } from './editor/image.js';
// pos — куда вставить (по умолчанию — в позицию курсора)
export async function insertImageFiles(editor, files, pos) {
  const nodes = [];
  for (const f of files) {
    try {
      const bytes = f.data || (await f.arrayBuffer());
      const mime = f.type || mimeOf(f.name || '');
      nodes.push({ type: 'image', attrs: { ...(await prepareImage(bytes, mime)), alt: '' } });
    } catch (e) {
      fileApi.showError('Не удалось вставить рисунок', `${f.name || ''}\n\n${e.message || e}`);
    }
  }
  if (!nodes.length) return;
  const chain = editor.chain().focus();
  (pos != null ? chain.insertContentAt(pos, nodes) : chain.insertContent(nodes)).run();
}
