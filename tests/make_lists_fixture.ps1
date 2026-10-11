# Генерирует tests/fixtures/lists.docx средствами Word (COM): маркированные списки со знаками «➢» и «✓» (Wingdings),
# нумерованный список 1–2, абзац, его продолжение 3–4, абзац, новый список 1–2 (начать заново).
# Автор — нейтральный «Перо Тест». Файл в UTF-8 с BOM (Windows PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_lists_fixture.ps1 tests/fixtures/lists.docx
param([string]$Out = 'tests/fixtures/lists.docx')
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
  $paras = @('стрелка один', 'стрелка два', 'галка один', 'галка два', 'пункт один', 'пункт два', 'Текст между списками.',
    'пункт три', 'пункт четыре', 'Ещё текст.', 'новый один', 'новый два')
  $d.Content.Text = ($paras -join "`r")
  $bullet = {
    param($code, $from, $to)
    $t = $d.ListTemplates.Add($false)
    $l = $t.ListLevels.Item(1)
    $l.NumberFormat = [string][char]$code
    $l.Font.Name = 'Wingdings'
    $l.NumberStyle = 23
    $r = $d.Range($d.Paragraphs.Item($from).Range.Start, $d.Paragraphs.Item($to).Range.End)
    $r.ListFormat.ApplyListTemplate($t, $false, 0)
  }
  & $bullet 0xF0D8 1 2
  & $bullet 0xF0FC 3 4
  $num = $d.ListTemplates.Add($false)
  $num.ListLevels.Item(1).NumberStyle = 0
  $num.ListLevels.Item(1).NumberFormat = '%1.'
  $r = $d.Range($d.Paragraphs.Item(5).Range.Start, $d.Paragraphs.Item(6).Range.End)
  $r.ListFormat.ApplyListTemplate($num, $false, 0)
  # продолжение того же списка
  $r = $d.Range($d.Paragraphs.Item(8).Range.Start, $d.Paragraphs.Item(9).Range.End)
  $r.ListFormat.ApplyListTemplate($num, $true, 0)
  # новый список (начать заново)
  $r = $d.Range($d.Paragraphs.Item(11).Range.Start, $d.Paragraphs.Item(12).Range.End)
  $r.ListFormat.ApplyListTemplate($num, $false, 0)
  $d.SaveAs2($path, 16)
  $d.Close(0)
  "ok: $path"
} finally {
  $w.UserName = $oldName
  $w.UserInitials = $oldInitials
  $w.Options.UseLocalUserInfo = $oldLocal
  $w.Quit()
}
