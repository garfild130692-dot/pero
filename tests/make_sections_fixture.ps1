# Генерирует tests/fixtures/sections.docx средствами Word (COM): три раздела с разрывом «со следующей страницы» —
# книжный A4 (поля 2/1,5/2/3 см), альбомный A4 с широкой таблицей (поля 1,5 см), снова книжный. Автор — «Перо Тест».
# Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_sections_fixture.ps1 tests/fixtures/sections.docx
param([string]$Out = 'tests/fixtures/sections.docx')
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
  $text = ('Текст раздела для заполнения страниц и проверки разбиения. ' * 10).Trim()
  $sel = $w.Selection
  for ($i = 1; $i -le 14; $i++) { $sel.TypeText("Книжный раздел, абзац $i. $text"); $sel.TypeParagraph() }
  $sel.InsertBreak(2)   # wdSectionBreakNextPage
  for ($i = 1; $i -le 3; $i++) { $sel.TypeText("Альбомный раздел, абзац $i. $text"); $sel.TypeParagraph() }
  $t = $d.Tables.Add($sel.Range, 12, 6)
  $t.Borders.Enable = 1
  for ($r = 1; $r -le 12; $r++) { for ($c = 1; $c -le 6; $c++) { $t.Cell($r, $c).Range.Text = "Ячейка $r.$c данные" } }
  $sel.EndKey(6) | Out-Null   # конец документа
  $sel.TypeParagraph()
  $sel.InsertBreak(2)
  for ($i = 1; $i -le 10; $i++) { $sel.TypeText("Снова книжный раздел, абзац $i. $text"); $sel.TypeParagraph() }
  $cm = 28.3465
  $s1 = $d.Sections.Item(1).PageSetup
  $s1.TopMargin = 2 * $cm; $s1.BottomMargin = 2 * $cm; $s1.LeftMargin = 3 * $cm; $s1.RightMargin = 1.5 * $cm
  $s2 = $d.Sections.Item(2).PageSetup
  $s2.Orientation = 1   # wdOrientLandscape
  $s2.TopMargin = 1.5 * $cm; $s2.BottomMargin = 1.5 * $cm; $s2.LeftMargin = 1.5 * $cm; $s2.RightMargin = 1.5 * $cm
  $s3 = $d.Sections.Item(3).PageSetup
  $s3.Orientation = 0
  $s3.TopMargin = 2 * $cm; $s3.BottomMargin = 2 * $cm; $s3.LeftMargin = 3 * $cm; $s3.RightMargin = 1.5 * $cm
  $d.SaveAs2($path, 16)
  'sections=' + $d.Sections.Count + ' pages=' + $d.ComputeStatistics(2)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
