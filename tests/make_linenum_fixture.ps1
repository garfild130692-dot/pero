# Генерирует tests/fixtures/linenum.docx средствами Word (COM): нумерация строк (начать с 5, шаг 2, на каждой странице),
# длинные абзацы, пустой абзац, таблица, абзац с запретом нумерации. Автор — нейтральный «Перо Тест». UTF-8 с BOM.
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_linenum_fixture.ps1 tests/fixtures/linenum.docx
param([string]$Out = 'tests/fixtures/linenum.docx')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$oldName = $w.UserName; $oldInitials = $w.UserInitials; $oldLocal = $w.Options.UseLocalUserInfo
$w.Options.UseLocalUserInfo = $true; $w.UserName = 'Перо Тест'; $w.UserInitials = 'ПТ'
try {
  $d = $w.Documents.Add()
  $long = ('Договор поставки заключён между сторонами на условиях, изложенных ниже, и вступает в силу с даты подписания. ' * 4).Trim()
  $paras = @('Нумерация строк.', $long, '', 'Абзац без номеров строк — запрет в абзаце.', $long, 'ТАБЛИЦА', $long)
  $d.Content.Text = ($paras -join "`r")
  $r = $d.Paragraphs.Item(6).Range
  $t = $d.Tables.Add($r, 2, 2)
  $t.Cell(1, 1).Range.Text = 'ячейка 1'; $t.Cell(2, 2).Range.Text = 'ячейка 4'
  $ln = $d.Sections.Item(1).PageSetup.LineNumbering
  $ln.Active = $true; $ln.StartingNumber = 5; $ln.CountBy = 2; $ln.RestartMode = 0; $ln.DistanceFromText = 18
  # запрет нумерации у абзаца — после включения нумерации (иначе свойство недоступно)
  $d.Paragraphs.Item(4).Format.NoLineNumber = -1
  $d.SaveAs2($path, 16)
  Write-Output ('строк в абзацах: ' + (($d.Paragraphs | ForEach-Object { $_.Range.ComputeStatistics(1) }) -join ','))
  $d.Close($false)
} finally {
  $w.UserName = $oldName; $w.UserInitials = $oldInitials; $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
