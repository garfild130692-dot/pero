# Генерирует tests/fixtures/wrap.docx средствами Word (COM): рисунки с обтеканием — квадрат слева у абзаца, квадрат
# справа от поля, по контуру по центру (текст с двух сторон), сверху и снизу, печать «за текстом», рисунок «перед текстом».
# Пишет tests/fixtures/wrap_word.json: начало каждой строки (лист, x, y в pt от края листа, первое слово) и рамки рисунков.
# Автор — нейтральный «Перо Тест». UTF-8 с BOM. Запуск: powershell -ExecutionPolicy Bypass -File tests/make_wrap_fixture.ps1
param([string]$Out = 'tests/fixtures/wrap.docx', [string]$Json = 'tests/fixtures/wrap_word.json')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$photo = [IO.Path]::GetFullPath('tests/fixtures/photo.png'); $stamp = [IO.Path]::GetFullPath('tests/fixtures/stamp.png')
$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$oldName = $w.UserName; $oldInitials = $w.UserInitials; $oldLocal = $w.Options.UseLocalUserInfo
$w.Options.UseLocalUserInfo = $true; $w.UserName = 'Перо Тест'; $w.UserInitials = 'ПТ'
try {
  $d = $w.Documents.Add()
  $words = @('договор','поставка','сторона','оплата','срок','качество','товар','претензия','уведомление','гарантия','реквизиты','подпись')
  $k = 0
  $para = { param($n) $t = @(); for ($i = 0; $i -lt $n; $i++) { $t += $words[$script:k % $words.Count] + $script:k; $script:k++ }; ($t -join ' ') + '.' }
  $texts = @()
  for ($i = 1; $i -le 12; $i++) { $texts += ("Абзац $i " + (& $para 70)) }
  $d.Content.Text = ($texts -join "`r")
  $cm = { param($v) $w.CentimetersToPoints($v) }
  # рисунок: абзац, файл, ширина (см), тип обтекания, относительно по горизонтали/вертикали, слева/сверху (pt или константы)
  $add = { param($p, $file, $wcm, $wrap, $relH, $left, $relV, $top)
    $r = $d.Paragraphs.Item($p).Range
    $r.Collapse(1)
    # в строку нужного абзаца, потом в плавающий: так привязка остаётся у этого абзаца
    $s = $d.InlineShapes.AddPicture($file, $false, $true, $r).ConvertToShape()
    $s.LockAnchor = -1   # иначе Word перепривязывает рисунок к абзацу, ближайшему к его положению
    $s.LockAspectRatio = -1; $s.Width = (& $cm $wcm)
    $s.WrapFormat.Type = $wrap
    $s.RelativeHorizontalPosition = $relH; $s.RelativeVerticalPosition = $relV
    $s.Left = [single]$left; $s.Top = [single]$top
    $s }
  # WdWrapType: 0 квадрат, 1 по контуру, 3 перед текстом, 4 сверху и снизу, 5 за текстом; позиции: 0 поле, 1 лист, 2 колонка; 2 абзац
  $null = & $add 2 $photo 5 0 2 0 2 0
  $null = & $add 4 $photo 4.5 0 0 -999996 2 (& $cm 0.5)
  $null = & $add 6 $photo 4 1 0 -999995 2 (& $cm 0.3)
  $null = & $add 8 $photo 6 4 2 (& $cm 2) 2 0
  $null = & $add 10 $stamp 4 5 2 (& $cm 8) 2 (& $cm 0.2)
  $null = & $add 12 $photo 4 3 2 (& $cm 3) 2 (& $cm 0.5)
  $d.SaveAs2($path, 16)
  # начало каждой строки
  $lines = @(); $prevY = -1; $prevPage = 0
  foreach ($wd in $d.Words) {
    $t = $wd.Text.Trim(); if ($t.Length -eq 0) { continue }
    $pg = $wd.Information(3); $x = $wd.Information(5); $y = $wd.Information(6)
    if ($pg -ne $prevPage -or [Math]::Abs($y - $prevY) -gt 1) { $lines += [pscustomobject]@{ page = $pg; x = [Math]::Round($x, 1); y = [Math]::Round($y, 1); text = $t } }
    $prevPage = $pg; $prevY = $y
  }
  $shapes = @(); foreach ($s in $d.Shapes) { $shapes += [pscustomobject]@{ name = $s.Name; wrap = $s.WrapFormat.Type; w = [Math]::Round($s.Width, 1); h = [Math]::Round($s.Height, 1) } }
  $text = [pscustomobject]@{ lines = $lines; shapes = $shapes } | ConvertTo-Json -Depth 4
  [IO.File]::WriteAllText([IO.Path]::GetFullPath($Json), $text, (New-Object Text.UTF8Encoding($false)))
  Write-Output ('строк: ' + $lines.Count + ', листов: ' + $d.ComputeStatistics(2))
  $d.Close($false)
} finally {
  $w.UserName = $oldName; $w.UserInitials = $oldInitials; $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
