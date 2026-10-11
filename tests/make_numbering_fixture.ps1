# Генерирует tests/fixtures/numbering.docx средствами Word (COM): заголовки с многоуровневой нумерацией, связанной
# со стилями «Заголовок 1–3» (1 / 1.1 / 1.1.1), заголовок без номера, нумерованный список «a)» и «I.».
# Автор — нейтральный «Перо Тест». Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_numbering_fixture.ps1 tests/fixtures/numbering.docx
param([string]$Out = 'tests/fixtures/numbering.docx')
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
  $paras = @('Общие положения', 'Термины', 'Текст.', 'Определения', 'Подробнее', 'Обязанности сторон', 'Примечание без номера',
    'первый пункт', 'второй пункт', 'Римский один', 'Римский два')
  $d.Content.Text = ($paras -join "`r")
  $styles = @(-2, -3, -1, -3, -4, -2, -2, -1, -1, -1, -1)
  for ($i = 0; $i -lt $styles.Count; $i++) { $d.Paragraphs.Item($i + 1).Style = $styles[$i] }
  # многоуровневый список, связанный со стилями заголовков
  $lt = $d.ListTemplates.Add($true)
  for ($k = 1; $k -le 3; $k++) {
    $lv = $lt.ListLevels.Item($k)
    $lv.NumberFormat = ((1..$k | ForEach-Object { "%$_" }) -join '.')
    $lv.NumberStyle = 0
    $lv.StartAt = 1
    $lv.TrailingCharacter = 0
    $lv.NumberPosition = 0
    $lv.TextPosition = $w.CentimetersToPoints(0.76 + 0.25 * ($k - 1))
    $lv.TabPosition = $lv.TextPosition
    $lv.LinkedStyle = $d.Styles.Item(-1 - $k).NameLocal
  }
  $d.Paragraphs.Item(1).Range.ListFormat.ApplyListTemplate($lt, $false, 0)
  # заголовок без номера (numId 0)
  $d.Paragraphs.Item(7).Range.ListFormat.RemoveNumbers()
  # список «a)»
  $la = $d.ListTemplates.Add($false)
  $la.ListLevels.Item(1).NumberFormat = '%1)'
  $la.ListLevels.Item(1).NumberStyle = 4
  $r = $d.Range($d.Paragraphs.Item(8).Range.Start, $d.Paragraphs.Item(9).Range.End)
  $r.ListFormat.ApplyListTemplate($la, $false, 0)
  # список «I.»
  $lr = $d.ListTemplates.Add($false)
  $lr.ListLevels.Item(1).NumberFormat = '%1.'
  $lr.ListLevels.Item(1).NumberStyle = 1
  $r = $d.Range($d.Paragraphs.Item(10).Range.Start, $d.Paragraphs.Item(11).Range.End)
  $r.ListFormat.ApplyListTemplate($lr, $false, 0)
  $d.SaveAs2($path, 16)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
