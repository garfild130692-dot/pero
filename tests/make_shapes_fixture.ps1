# Генерирует tests/fixtures/shapes.docx средствами Word (COM): надпись с текстом (обтекание вокруг рамки слева),
# прямоугольник с заливкой за текстом, эллипс перед текстом, стрелка вправо, линия со стрелкой, скруглённый прямоугольник
# с текстом по центру, надпись сверху и снизу. Пишет tests/fixtures/shapes_word.json — начало каждой строки (как wrap).
# Автор — нейтральный «Перо Тест». UTF-8 с BOM. Запуск: powershell -ExecutionPolicy Bypass -File tests/make_shapes_fixture.ps1
param([string]$Out = 'tests/fixtures/shapes.docx', [string]$Json = 'tests/fixtures/shapes_word.json')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
$oldName = $w.UserName; $oldInitials = $w.UserInitials; $oldLocal = $w.Options.UseLocalUserInfo
$w.Options.UseLocalUserInfo = $true; $w.UserName = 'Перо Тест'; $w.UserInitials = 'ПТ'
try {
  $d = $w.Documents.Add()
  $words = @('договор','поставка','сторона','оплата','срок','качество','товар','претензия','уведомление','гарантия','реквизиты','подпись')
  $k = 0
  $para = { param($n) $t = @(); for ($i = 0; $i -lt $n; $i++) { $t += $words[$script:k % $words.Count] + $script:k; $script:k++ }; ($t -join ' ') + '.' }
  $texts = @(); for ($i = 1; $i -le 8; $i++) { $texts += ("Абзац $i " + (& $para 70)) }
  $d.Content.Text = ($texts -join "`r")
  $cm = { param($v) [single]$w.CentimetersToPoints($v) }
  $at = { param($p) $r = $d.Paragraphs.Item($p).Range; $r.Collapse(1); $r }
  # фигура: тип, абзац, слева, сверху (от колонки / абзаца), ширина, высота, обтекание
  $add = { param($type, $p, $l, $t, $wd, $ht, $wrap)
    $s = $d.Shapes.AddShape($type, 0, 0, (& $cm $wd), (& $cm $ht), (& $at $p))
    $s.LockAnchor = -1; $s.WrapFormat.Type = $wrap
    $s.RelativeHorizontalPosition = 2; $s.RelativeVerticalPosition = 2; $s.Left = (& $cm $l); $s.Top = (& $cm $t)
    $s }
  # 1 — прямоугольник, 5 — скруглённый, 9 — эллипс, 33 — стрелка вправо; обтекание: 0 квадрат, 3 перед, 4 сверху и снизу, 5 за
  $tb = $d.Shapes.AddTextbox(1, 0, 0, (& $cm 5), (& $cm 3), (& $at 2))
  $tb.LockAnchor = -1; $tb.WrapFormat.Type = 0; $tb.RelativeHorizontalPosition = 2; $tb.RelativeVerticalPosition = 2; $tb.Left = 0; $tb.Top = 0
  $tb.TextFrame.TextRange.Text = 'Надпись: текст внутри рамки'; $tb.TextFrame.TextRange.Font.Bold = -1
  $s = & $add 1 3 1 0.5 14 2.5 5; $s.Fill.ForeColor.RGB = 0xE6F0DA; $s.Line.Visible = 0
  $s = & $add 9 4 11 0.2 4 2.5 3; $s.Fill.ForeColor.RGB = 0x50B0F0
  $s = & $add 33 5 2 0.3 5 1.5 3
  $ln = $d.Shapes.AddLine(0, 0, (& $cm 6), (& $cm 1.5), (& $at 6)); $ln.LockAnchor = -1; $ln.RelativeHorizontalPosition = 2; $ln.RelativeVerticalPosition = 2
  $ln.Left = (& $cm 8); $ln.Top = (& $cm 0.4); $ln.Line.EndArrowheadStyle = 2; $ln.Line.Weight = 2
  $s = & $add 5 7 10 0 6 2 0; $s.TextFrame.TextRange.Text = 'Скруглённый'; $s.TextFrame.TextRange.ParagraphFormat.Alignment = 1
  $tb2 = $d.Shapes.AddTextbox(1, 0, 0, (& $cm 12), (& $cm 1.5), (& $at 8))
  $tb2.LockAnchor = -1; $tb2.WrapFormat.Type = 4; $tb2.RelativeHorizontalPosition = 2; $tb2.RelativeVerticalPosition = 2; $tb2.Left = (& $cm 2); $tb2.Top = 0
  $tb2.TextFrame.TextRange.Text = 'Надпись сверху и снизу'
  $d.SaveAs2($path, 16)
  $lines = @(); $prevY = -1; $prevPage = 0
  foreach ($wd in $d.Content.Words) {
    $t = $wd.Text.Trim(); if ($t.Length -eq 0) { continue }
    $pg = $wd.Information(3); $x = $wd.Information(5); $y = $wd.Information(6)
    if ($pg -ne $prevPage -or [Math]::Abs($y - $prevY) -gt 1) { $lines += [pscustomobject]@{ page = $pg; x = [Math]::Round($x, 1); y = [Math]::Round($y, 1); text = $t } }
    $prevPage = $pg; $prevY = $y
  }
  $text = [pscustomobject]@{ lines = $lines } | ConvertTo-Json -Depth 4
  [IO.File]::WriteAllText([IO.Path]::GetFullPath($Json), $text, (New-Object Text.UTF8Encoding($false)))
  Write-Output ('фигур: ' + $d.Shapes.Count + ', строк: ' + $lines.Count)
  $d.Close($false)
} finally {
  $w.UserName = $oldName; $w.UserInitials = $oldInitials; $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
