# Генерирует tests/fixtures/design.docx средствами Word (COM): подложка «ОБРАЗЕЦ» (как «Конструктор → Подложка»:
# надпись WordArt в колонтитуле, повёрнута на 315°, полупрозрачная серая), цвет страницы, рамка страницы.
# Автор — нейтральный «Перо Тест». Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_design_fixture.ps1 tests/fixtures/design.docx
param([string]$Out = 'tests/fixtures/design.docx')
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
  $d.Content.Text = "Договор поставки.`rТекст документа."
  # подложка — так, как её вставляет сам Word (запись макроса «Конструктор → Подложка»)
  $hdr = $d.Sections.Item(1).Headers.Item(1)
  $s = $hdr.Shapes.AddTextEffect(0, 'ОБРАЗЕЦ', 'Calibri', 1, $false, $false, 0, 0)
  $s.Name = 'PowerPlusWaterMarkObject1'
  $s.TextEffect.NormalizedHeight = $false
  $s.Line.Visible = $false
  $s.Fill.Visible = $true
  $s.Fill.Solid()
  $s.Fill.ForeColor.RGB = 12632256   # серебристый (C0C0C0)
  $s.Fill.Transparency = 0.5
  $s.Rotation = 315
  $s.LockAspectRatio = $false
  $s.Height = $w.CentimetersToPoints(4.65)
  $s.Width = $w.CentimetersToPoints(18.62)
  $s.LockAspectRatio = $true
  $s.WrapFormat.AllowOverlap = $true
  $s.WrapFormat.Side = 3
  $s.WrapFormat.Type = 3
  $s.RelativeHorizontalPosition = 0
  $s.RelativeVerticalPosition = 0
  $s.Left = -999995
  $s.Top = -999995
  # цвет страницы (светло-жёлтый) и рамка страницы 1 пт от края листа
  $d.Background.Fill.ForeColor.RGB = 13434879  # FFFFCC
  $d.Background.Fill.Visible = $true
  $d.Background.Fill.Solid()
  $b = $d.Sections.Item(1).Borders
  $b.DistanceFrom = 1
  $b.OutsideLineStyle = 1
  $b.OutsideLineWidth = 8
  $b.OutsideColor = 12611584   # синий 0070C0
  $d.SaveAs2($path, 16)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
