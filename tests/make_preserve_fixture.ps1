# Генерирует tests/fixtures/preserve.docx средствами самого Word (COM): «настоящий» документ с тем,
# что Перо пока не редактирует, но обязано сохранить: колонтитулы, сноска, комментарий, рисунок, закладка
# и ссылка на неё, поле, формула, символ, разрыв страницы, второй раздел с альбомной ориентацией, поля страницы.
# Файл в UTF-8 с BOM — иначе Windows PowerShell 5.1 прочитает кириллицу как ANSI.
# Запуск (нужен установленный Word):
#   powershell -ExecutionPolicy Bypass -File tests/make_preserve_fixture.ps1 tests/fixtures/preserve.docx
param([string]$Out = 'tests/fixtures/preserve.docx')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$img = [IO.Path]::GetFullPath('build/icon.png')
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  $d = $w.Documents.Add()
  $s = $d.Sections.Item(1)
  # поля страницы: 2,5 / 2 / 2,5 / 3 см; колонтитулы
  $s.PageSetup.TopMargin = $w.CentimetersToPoints(2.5)
  $s.PageSetup.BottomMargin = $w.CentimetersToPoints(2.5)
  $s.PageSetup.LeftMargin = $w.CentimetersToPoints(3)
  $s.PageSetup.RightMargin = $w.CentimetersToPoints(2)
  $s.Headers.Item(1).Range.Text = 'Верхний колонтитул Перо'
  $s.Footers.Item(1).Range.Text = 'Нижний колонтитул Перо'

  $r = $d.Content
  $r.Text = 'Первый абзац со сноской'
  $p1 = $d.Paragraphs.Item(1).Range
  $end = $d.Range($p1.End - 1, $p1.End - 1)
  $d.Footnotes.Add($end, [Type]::Missing, 'Текст сноски') | Out-Null

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'Абзац с комментарием'
  $d.Comments.Add($d.Paragraphs.Last.Range, 'Комментарий рецензента') | Out-Null

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'Рисунок: '
  $pos = $d.Range($d.Paragraphs.Last.Range.End - 1, $d.Paragraphs.Last.Range.End - 1)
  $d.InlineShapes.AddPicture($img, $false, $true, $pos) | Out-Null

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'Закладка здесь'
  $d.Bookmarks.Add('Цель', $d.Paragraphs.Last.Range) | Out-Null

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'Ссылка на закладку'
  $d.Hyperlinks.Add($d.Paragraphs.Last.Range, '', 'Цель') | Out-Null

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'Поле даты: '
  $pos = $d.Range($d.Paragraphs.Last.Range.End - 1, $d.Paragraphs.Last.Range.End - 1)
  $d.Fields.Add($pos, 31, 'DATE \@ "dd.MM.yyyy"', $false) | Out-Null   # 31 = wdFieldDate

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'E=mc2'
  $m = $d.OMaths.Add($d.Paragraphs.Last.Range)
  $m.OMaths.Item(1).BuildUp() | Out-Null

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'Символ: '
  $pos = $d.Range($d.Paragraphs.Last.Range.End - 1, $d.Paragraphs.Last.Range.End - 1)
  $pos.InsertSymbol(0xF0B7 - 65536, 'Symbol', $true)   # код — знаковое 16-битное число

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'Перед разрывом страницы'
  $pos = $d.Range($d.Paragraphs.Last.Range.End - 1, $d.Paragraphs.Last.Range.End - 1)
  $pos.InsertBreak(7)   # 7 = wdPageBreak

  $r = $d.Content; $r.InsertParagraphAfter(); $r = $d.Paragraphs.Last.Range
  $r.Text = 'Конец первого раздела'
  $pos = $d.Range($d.Paragraphs.Last.Range.End - 1, $d.Paragraphs.Last.Range.End - 1)
  $pos.InsertBreak(2)   # 2 = wdSectionBreakNextPage
  $last = $d.Sections.Item($d.Sections.Count)
  $last.PageSetup.Orientation = 1   # альбомная
  $last.Range.Paragraphs.Last.Range.Text = 'Второй раздел, альбомная ориентация'

  $d.SaveAs2($path, 16)   # 16 = wdFormatDocumentDefault (.docx)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.Quit()
}
