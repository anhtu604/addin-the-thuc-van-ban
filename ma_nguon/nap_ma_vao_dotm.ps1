param([string]$Src, [string]$Bas, [string]$Out)
$ErrorActionPreference = 'Stop'
$before = @(Get-Process WINWORD -EA 0 | % Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false; $word.DisplayAlerts = 0
try {
  $code = [IO.File]::ReadAllText($Bas, [Text.Encoding]::UTF8)
  $d = $word.Documents.Open($Src)
  $cm = $d.VBProject.VBComponents.Item('Module1').CodeModule
  $cm.DeleteLines(1, $cm.CountOfLines)
  $cm.AddFromString($code)
  "lines now: " + $cm.CountOfLines
  $d.SaveAs2([ref]$Out, [ref]15)
  $d.Close([ref]0)
} finally { $word.Quit([ref]0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
Start-Sleep 1; Get-Process WINWORD -EA 0 | ? { $before -notcontains $_.Id -and -not $_.MainWindowTitle } | Stop-Process -Force
