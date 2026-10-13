<#
.SYNOPSIS
  Wgrywa moduły standardowe z src\ (Module1-3) do PW_integrationg.xlsm.
.NOTES
  Wymaga Windows + Excel oraz włączonej opcji:
  Plik > Opcje > Centrum zaufania > Ustawienia makr > "Ufaj dostępowi do modelu obiektowego projektu VBA".
  Formularze i kod arkuszy (.frm/.cls) NIE są importowane - zostają w skoroszycie bez zmian.
#>
param([string]$Workbook = (Join-Path $PSScriptRoot 'PW_integrationg.xlsm'))
$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP 'pw_vba_import'
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
New-Item $tmp -ItemType Directory | Out-Null
$ansi = [System.Text.Encoding]::GetEncoding(1250)   # VBE oczekuje kodowania ANSI, nie UTF-8
$modules = 'Module1','Module2','Module3'
foreach ($m in $modules) {
    $text = (Get-Content (Join-Path $PSScriptRoot "src\$m.bas") -Raw -Encoding UTF8) -replace "`r?`n", "`r`n"
    [IO.File]::WriteAllText((Join-Path $tmp "$m.bas"), $text, $ansi)
}
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Workbook)
    $vbp = $wb.VBProject
    foreach ($m in $modules) { $vbp.VBComponents.Remove($vbp.VBComponents.Item($m)) }
    foreach ($m in $modules) { [void]$vbp.VBComponents.Import((Join-Path $tmp "$m.bas")) }
    $wb.Save()
    $wb.Close($true)
    Write-Host 'OK: moduly zaimportowane.' -ForegroundColor Green
} finally { $xl.Quit() }
