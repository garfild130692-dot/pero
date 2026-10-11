# Генерирует tests/fixtures/math.docx средствами Word (COM): формулы (OMML), построенные Word из линейной записи —
# дробь, корень и корень n-й степени, индексы, сумма и интеграл с пределами, предел, функции, матрица, система, акцент,
# черта; плюс формула внутри строки текста. Автор — нейтральный «Перо Тест». Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_math_fixture.ps1 tests/fixtures/math.docx
param([string]$Out = 'tests/fixtures/math.docx')
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
  $formulas = @(
    (-join ([char[]]@(120,61,40,45,98,177,8730,40,98,94,50,45,52,97,99,41,41,47,50,97))),
    (-join ([char[]]@(8721,95,40,105,61,49,41,94,110,9618,105,94,50,32,61,40,110,40,110,43,49,41,40,50,110,43,49,41,41,47,54))),
    (-join ([char[]]@(8747,95,48,94,49,9618,12310,120,94,50,32,100,120,12311,61,49,47,51))),
    (-join ([char[]]@(108,105,109,95,40,110,8594,8734,41,8289,12310,40,49,43,49,47,110,41,94,110,32,12311,61,101))),
    (-join ([char[]]@(115,105,110,94,50,8289,945,43,99,111,115,94,50,8289,945,61,49))),
    (-join ([char[]]@(65,61,91,9632,40,49,38,48,64,48,38,49,41,93))),
    (-join ([char[]]@(8730,40,51,38,120,43,49,41,43,120,95,105,94,50,43,40,120,770,41,43,40,121,773,41))),
    (-join ([char[]]@(124,120,124,61,123,9632,40,120,44,38,120,8805,48,64,45,120,44,38,120,60,48,41,9508)))
  )
  $d.Content.Text = 'Формулы из Word.'
  foreach ($f in $formulas) {
    $d.Content.InsertParagraphAfter()
    $r = $d.Paragraphs.Last.Range
    $r.Text = $f
    $m = $d.OMaths.Add($d.Paragraphs.Last.Range)
    $m.OMaths.Item($m.OMaths.Count).BuildUp() | Out-Null
  }
  # формула внутри строки текста
  $d.Content.InsertParagraphAfter()
  $p = $d.Paragraphs.Last.Range
  $p.Text = 'Площадь круга равна S=πr^2 при радиусе r.'
  $s = $p.Start + $p.Text.IndexOf('S=')
  $r = $d.Range($s, $s + 'S=πr^2'.Length)
  $m = $d.OMaths.Add($r)
  $m.OMaths.Item(1).BuildUp() | Out-Null
  $d.SaveAs2($path, 16)
  Write-Output ('формул: ' + $d.OMaths.Count)
  $d.Close($false)
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
