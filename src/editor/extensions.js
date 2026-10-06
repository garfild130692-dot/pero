// Набор расширений редактора — единая схема документа для UI и для тестов.
import StarterKit from '@tiptap/starter-kit';
import TextAlign from '@tiptap/extension-text-align';
import { TableKit } from '@tiptap/extension-table';

export const extensions = [
  StarterKit.configure({
    // то, чего нет в Word как отдельных сущностей, отключаем
    code: false,
    codeBlock: false,
    blockquote: false,
    horizontalRule: false,
    heading: { levels: [1, 2, 3, 4, 5, 6] },
    link: { openOnClick: false, autolink: true, defaultProtocol: 'https' },
  }),
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  TableKit.configure({ table: { resizable: true } }),
];
