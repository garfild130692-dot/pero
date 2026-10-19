# Эталон tests/fixtures/extra.docx (этап 19.4, после ревью): то, чего не было в прежних эталонах .doc —
# рисунок в верхнем колонтитуле, гиперссылки на веб/почту/файл/UNC, гиперссылка в сноске, закладка вокруг текста последней
# ячейки строки таблицы, широкая таблица (26 столбцов, цветные границы — у Word свойства строки уходят в sprmPHugePapx).
# Дальше: powershell -File tests/make_doc_fixtures.ps1 extra → extra.doc + extra.doc.docx.
# Автор — нейтральный «Перо Тест» (как в make_doc_fixtures.ps1). Файл в UTF-8 с BOM (PowerShell 5.1).
$ErrorActionPreference = 'Stop'
$dir = [IO.Path]::GetFullPath('tests/fixtures')
$out = [string](Join-Path $dir 'extra.docx')
$logo = [string]([IO.Path]::GetFullPath('build/icon.png'))
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$oldName = $w.UserName; $oldInitials = $w.UserInitials; $oldLocal = $w.Options.UseLocalUserInfo
if (-not $oldLocal) { $w.Options.UseLocalUserInfo = $true }
if ($oldName -ne 'Перо Тест') { $w.UserName = 'Перо Тест' }
if ($oldInitials -ne 'ПТ') { $w.UserInitials = 'ПТ' }
try {
  $d = $w.Documents.Add()
  # колонтитул с рисунком
  $h = $d.Sections.Item(1).Headers.Item(1).Range
  $h.Text = 'Бланк '
  $h.Collapse(0)
  [void]$d.Sections.Item(1).Headers.Item(1).Range.InlineShapes.AddPicture($logo, $false, $true)
  $s = $w.Selection
  $s.TypeText('Ссылки: ')
  [void]$d.Hyperlinks.Add($s.Range, 'https://example.com/путь?a=1&b=2', '', '', 'сайт')
  $s.TypeText(', ')
  [void]$d.Hyperlinks.Add($s.Range, 'mailto:test@example.com', '', '', 'почта')
  $s.TypeText(', ')
  [void]$d.Hyperlinks.Add($s.Range, 'file:///C:/Windows/win.ini', '', '', 'файл')
  $s.TypeText(', ')
  [void]$d.Hyperlinks.Add($s.Range, '\\server\share\doc.txt', '', '', 'сеть')
  $s.TypeText('.')
  $s.TypeParagraph()
  $s.TypeText('Абзац со сноской')
  $fn = $d.Footnotes.Add($s.Range)
  $fn.Range.Text = 'Сноска: '
  $r = $fn.Range; $r.Collapse(0)
  [void]$d.Hyperlinks.Add($r, 'https://example.org/', '', '', 'ссылка в сноске')
  $s.EndKey(6) | Out-Null
  $s.TypeParagraph()
  # таблица 2×3, закладка вокруг текста последней ячейки первой строки
  $t = $d.Tables.Add($s.Range, 2, 3)
  $t.Borders.Enable = $true
  for ($i = 1; $i -le 2; $i++) { for ($j = 1; $j -le 3; $j++) { $t.Cell($i, $j).Range.Text = "Я$i$j" } }
  $c = $t.Cell(1, 3).Range; $c.MoveEnd(1, -1) | Out-Null
  [void]$d.Bookmarks.Add('ПоследняяЯчейка', $c)
  $s.EndKey(6) | Out-Null
  $s.TypeParagraph()
  # широкая таблица: 26 столбцов, у каждой ячейки свой цвет границы и заливка
  $t2 = $d.Tables.Add($s.Range, 2, 26)
  for ($j = 1; $j -le 26; $j++) {
    for ($i = 1; $i -le 2; $i++) {
      $cell = $t2.Cell($i, $j)
      $cell.Range.Text = [string][char](0x0410 + $j - 1)
      $cell.Shading.BackgroundPatternColor = (($j * 9) % 256) * 65536 + 200 * 256 + (($i * 60) % 256)
      foreach ($b in -1, -2, -3, -4) { $br = $cell.Borders.Item($b); $br.LineStyle = 1; $br.Color = ($j * 7919) % 16777215 }
    }
  }
  $s.EndKey(6) | Out-Null
  $s.TypeText('Конец.')
  if (Test-Path $out) { Remove-Item $out }
  $d.SaveAs2($out, 12)
  $d.Close()
  Write-Output "ok: $out"
} finally {
  if ($w.UserName -ne $oldName) { $w.UserName = $oldName }
  if ($w.UserInitials -ne $oldInitials) { $w.UserInitials = $oldInitials }
  if ($w.Options.UseLocalUserInfo -ne $oldLocal) { $w.Options.UseLocalUserInfo = $oldLocal }
  $w.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
}
