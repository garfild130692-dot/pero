# Генерирует tests/fixtures/toc.docx средствами Word (COM): оглавление (поле TOC), заголовки на разных страницах,
# закладка «Срок» и перекрёстные ссылки на неё (REF, PAGEREF), поля DATE и PAGE в тексте — эталон для полей этапа 14.
# Автор — нейтральный «Перо Тест». Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_toc_fixture.ps1 tests/fixtures/toc.docx
param([string]$Out = 'tests/fixtures/toc.docx')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$oldName = $w.UserName
$oldInitials = $w.UserInitials
$oldLocal = $w.Options.UseLocalUserInfo
$w.Options.UseLocalUserInfo = $true
$w.UserName = 'Перо Тест'
$w.UserInitials = 'ПТ'
try {
  $d = $w.Documents.Add()
  $paras = @('Место для оглавления', 'Общие положения', 'Срок поставки — 10 дней.', 'Предмет договора', 'Текст раздела.',
    'Ответственность сторон', 'См. срок: ', 'Стр. ')
  $d.Content.Text = ($paras -join "`r")
  $styles = @(-1, -2, -1, -3, -1, -2, -1, -1) # Обычный, Заголовок 1, Заголовок 2
  for ($i = 0; $i -lt $styles.Count; $i++) { $d.Paragraphs.Item($i + 1).Style = $styles[$i] }
  # заголовок «Ответственность сторон» — на новой странице
  $d.Paragraphs.Item(6).Format.PageBreakBefore = $true
  # закладка «Срок» на «10 дней»
  $rr = $d.Paragraphs.Item(3).Range
  $f = $rr.Find; $f.Execute('10 дней') | Out-Null
  $d.Bookmarks.Add('Срок', $rr) | Out-Null
  # ссылки на закладку: текст и номер страницы
  $p7 = $d.Paragraphs.Item(7).Range
  $at = $d.Range($p7.End - 1, $p7.End - 1)
  $d.Fields.Add($at, -1, 'REF Срок \h', $false) | Out-Null
  $p7 = $d.Paragraphs.Item(7).Range
  $at = $d.Range($p7.End - 1, $p7.End - 1)
  $at.InsertAfter(', стр. ')
  $p7 = $d.Paragraphs.Item(7).Range
  $at = $d.Range($p7.End - 1, $p7.End - 1)
  $d.Fields.Add($at, -1, 'PAGEREF Срок \h', $false) | Out-Null
  # поле PAGE и дата
  $p8 = $d.Paragraphs.Item(8).Range
  $at = $d.Range($p8.End - 1, $p8.End - 1)
  $d.Fields.Add($at, -1, 'PAGE', $false) | Out-Null
  $p8 = $d.Paragraphs.Item(8).Range
  $at = $d.Range($p8.End - 1, $p8.End - 1)
  $at.InsertAfter(', дата: ')
  $p8 = $d.Paragraphs.Item(8).Range
  $at = $d.Range($p8.End - 1, $p8.End - 1)
  $d.Fields.Add($at, -1, 'DATE \@ "dd.MM.yyyy"', $false) | Out-Null
  # оглавление вместо первого абзаца
  $first = $d.Paragraphs.Item(1).Range
  $d.TablesOfContents.Add($d.Range($first.Start, $first.End - 1), $true, 1, 3) | Out-Null
  $d.Fields.Update() | Out-Null
  $d.SaveAs2($path, 16)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
