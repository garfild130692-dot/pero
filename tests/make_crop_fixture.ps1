# Генерирует tests/fixtures/crop.docx средствами Word (COM): рисунок в строке с обрезкой (слева 1 см, сверху 0,5 см,
# справа 0,3 см, снизу 0,2 см), плавающий рисунок, повёрнутый на 30° и отражённый по горизонтали, рисунок, повёрнутый на 90°.
# Автор — нейтральный «Перо Тест». UTF-8 с BOM. Запуск: powershell -ExecutionPolicy Bypass -File tests/make_crop_fixture.ps1
param([string]$Out = 'tests/fixtures/crop.docx')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out); $photo = [IO.Path]::GetFullPath('tests/fixtures/photo.png')
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
$oldName = $w.UserName; $oldInitials = $w.UserInitials; $oldLocal = $w.Options.UseLocalUserInfo
$w.Options.UseLocalUserInfo = $true; $w.UserName = 'Перо Тест'; $w.UserInitials = 'ПТ'
try {
  $d = $w.Documents.Add()
  $d.Content.Text = "Обрезанный рисунок: `rПовёрнутый рисунок: `rПоворот на 90: "
  $cm = { param($v) [single]$w.CentimetersToPoints($v) }
  $r = $d.Paragraphs.Item(1).Range; $r.Collapse(0); $r.Move(1, -1) | Out-Null
  $i = $d.InlineShapes.AddPicture($photo, $false, $true, $r)
  $i.LockAspectRatio = 0; $i.Width = (& $cm 6); $i.Height = (& $cm 4.5)
  $f = $i.PictureFormat; $f.CropLeft = (& $cm 1); $f.CropTop = (& $cm 0.5); $f.CropRight = (& $cm 0.3); $f.CropBottom = (& $cm 0.2)
  $r = $d.Paragraphs.Item(2).Range; $r.Collapse(0); $r.Move(1, -1) | Out-Null
  $s = $d.InlineShapes.AddPicture($photo, $false, $true, $r).ConvertToShape()
  $s.LockAnchor = -1; $s.WrapFormat.Type = 0; $s.Width = (& $cm 5); $s.Rotation = 30; $s.Flip(0)
  $r = $d.Paragraphs.Item(3).Range; $r.Collapse(0); $r.Move(1, -1) | Out-Null
  $s2 = $d.InlineShapes.AddPicture($photo, $false, $true, $r).ConvertToShape()
  $s2.LockAnchor = -1; $s2.WrapFormat.Type = 0; $s2.Width = (& $cm 4); $s2.Rotation = 90
  $d.SaveAs2($path, 16)
  Write-Output ('рисунков: ' + $d.InlineShapes.Count + ' в строке, ' + $d.Shapes.Count + ' плавающих')
  $d.Close($false)
} finally {
  $w.UserName = $oldName; $w.UserInitials = $oldInitials; $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
