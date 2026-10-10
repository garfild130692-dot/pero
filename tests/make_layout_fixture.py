# Генерирует tests/fixtures/layout.docx средствами python-docx: положение абзаца на странице, границы и заливка абзаца.
# Запуск: python tests/make_layout_fixture.py tests/fixtures/layout.docx
import sys
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

d = Document()

# порядок элементов pPr по схеме — вставляем перед первым из «более поздних»
LATER_THAN_BDR = ['w:shd', 'w:tabs', 'w:spacing', 'w:ind', 'w:jc']
LATER_THAN_SHD = ['w:tabs', 'w:spacing', 'w:ind', 'w:jc']


def insert(pPr, el, later):
    for tag in later:
        a = pPr.find(qn(tag))
        if a is not None:
            a.addprevious(el)
            return
    pPr.append(el)


def border(p, sides, val='single', sz=4, space=1, color='auto'):
    bdr = OxmlElement('w:pBdr')
    for s in sides:
        e = OxmlElement(f'w:{s}')
        e.set(qn('w:val'), val); e.set(qn('w:sz'), str(sz)); e.set(qn('w:space'), str(space)); e.set(qn('w:color'), color)
        bdr.append(e)
    insert(p._p.get_or_add_pPr(), bdr, LATER_THAN_BDR)


def shade(p, fill):
    e = OxmlElement('w:shd'); e.set(qn('w:val'), 'clear'); e.set(qn('w:color'), 'auto'); e.set(qn('w:fill'), fill)
    insert(p._p.get_or_add_pPr(), e, LATER_THAN_SHD)


d.add_heading('Положение на странице', 1)
d.add_paragraph('Обычный абзац.')
p = d.add_paragraph('Не отрывать от следующего.'); p.paragraph_format.keep_with_next = True
p = d.add_paragraph('Не разрывать абзац.'); p.paragraph_format.keep_together = True
p = d.add_paragraph('С новой страницы.'); p.paragraph_format.page_break_before = True
p = d.add_paragraph('Без запрета висячих строк.'); p.paragraph_format.widow_control = False
h = d.add_heading('Заголовок, который можно отрывать', 2); h.paragraph_format.keep_with_next = False

# стиль со свойством «не отрывать» — значение переходит в абзац
st = d.styles.add_style('Подпись Перо', WD_STYLE_TYPE.PARAGRAPH)
st.base_style = d.styles['Normal']
st.paragraph_format.keep_with_next = True
d.add_paragraph('Абзац стиля с «не отрывать».', style='Подпись Перо')

p = d.add_paragraph('Рамка вокруг абзаца.'); border(p, ['top', 'left', 'bottom', 'right'], space=1)
p = d.add_paragraph('Нижняя граница двойная красная.'); border(p, ['bottom'], val='double', sz=6, color='C00000')
p = d.add_paragraph('Заливка жёлтым.'); shade(p, 'FFFF00')
p = d.add_paragraph('Рамка с заливкой, первый абзац группы.'); border(p, ['top', 'left', 'bottom', 'right', 'between']); shade(p, 'DEEBF7')
p = d.add_paragraph('Второй абзац группы.'); border(p, ['top', 'left', 'bottom', 'right', 'between']); shade(p, 'DEEBF7')

d.save(sys.argv[1])
