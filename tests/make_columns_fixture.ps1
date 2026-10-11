# Генерирует tests/fixtures/columns.docx средствами Word (COM): колонки текста в разделах — 2 колонки с разделителем
# на полтора листа (после заголовка, разрыв «на текущей странице»), 3 короткие выровненные колонки, 2 колонки «Слева»
# (разной ширины) с новой страницы. Пишет tests/fixtures/columns_word.json — начало каждой колонки на каждом листе.
# Автор — нейтральный «Перо Тест». UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_columns_fixture.ps1
param([string]$Out = 'tests/fixtures/columns.docx', [string]$Json = 'tests/fixtures/columns_word.json')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$oldName = $w.UserName; $oldInitials = $w.UserInitials; $oldLocal = $w.Options.UseLocalUserInfo
$w.Options.UseLocalUserInfo = $true; $w.UserName = 'Перо Тест'; $w.UserInitials = 'ПТ'
try {
  $d = $w.Documents.Add()
  $s = $w.Selection
  $words = @('договор','поставка','сторона','оплата','срок','качество','товар','претензия','уведомление','гарантия','реквизиты','подпись')
  $k = 0
  $para = { param($n) $t = @(); for ($i = 0; $i -lt $n; $i++) { $t += $words[$script:k % $words.Count] + $script:k; $script:k++ }; ($t -join ' ') + '.' }
  $s.TypeText('Заголовок на всю ширину листа'); $s.TypeParagraph()
  $s.InsertBreak(3)   # разрыв раздела на текущей странице
  for ($i = 1; $i -le 16; $i++) { $s.TypeText("Колонка А$i " + (& $para 60)); $s.TypeParagraph() }
  $s.InsertBreak(3)
  $s.TypeText('Абзац на всю ширину между разделами с колонками.'); $s.TypeParagraph()
  $s.InsertBreak(3)
  for ($i = 1; $i -le 3; $i++) { $s.TypeText("Три Б$i " + (& $para 25)); $s.TypeParagraph() }
  $s.InsertBreak(3)
  $s.TypeText('Снова на всю ширину.'); $s.TypeParagraph()
  $s.InsertBreak(2)   # со следующей страницы
  for ($i = 1; $i -le 4; $i++) { $s.TypeText("Слева В$i " + (& $para 30)); $s.TypeParagraph() }
  $s.TypeText('Конец документа.')
  # колонки разделов 2, 4, 6
  $c = $d.Sections.Item(2).PageSetup.TextColumns; $c.SetCount(2); $c.LineBetween = -1; $c.Spacing = $w.CentimetersToPoints(1.25)
  $c = $d.Sections.Item(4).PageSetup.TextColumns; $c.SetCount(3); $c.Spacing = $w.CentimetersToPoints(0.75)
  $c = $d.Sections.Item(6).PageSetup.TextColumns; $c.SetCount(2); $c.EvenlySpaced = 0
  $c.Item(1).Width = $w.CentimetersToPoints(4.5); $c.Item(1).SpaceAfter = $w.CentimetersToPoints(1.25); $c.Item(2).Width = $w.CentimetersToPoints(10.75)
  $d.SaveAs2($path, 16)
  # начало каждой колонки: слово, у которого лист сменился или текст «поднялся» на том же листе (новая колонка)
  $starts = @(); $prevPage = 0; $prevY = -1; $prevX = -1
  foreach ($wd in $d.Words) {
    $t = $wd.Text.Trim(); if ($t.Length -eq 0) { continue }
    $pg = $wd.Information(3); $x = $wd.Information(5); $y = $wd.Information(6)
    if ($pg -ne $prevPage -or ($y -lt $prevY - 1 -and $x -gt $prevX + 20)) {
      $r = $d.Range($wd.Start, [Math]::Min($d.Content.End, $wd.Start + 25))
      $starts += [pscustomobject]@{ page = $pg; x = [Math]::Round($x, 1); y = [Math]::Round($y, 1); text = $r.Text }
    }
    $prevPage = $pg; $prevY = $y; $prevX = $x
  }
  $text = $starts | ConvertTo-Json
  [IO.File]::WriteAllText([IO.Path]::GetFullPath($Json), $text, (New-Object Text.UTF8Encoding($false)))
  Write-Output ('колонок: ' + $starts.Count + ', листов: ' + $d.ComputeStatistics(2))
  $d.Close($false)
} finally {
  $w.UserName = $oldName; $w.UserInitials = $oldInitials; $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
