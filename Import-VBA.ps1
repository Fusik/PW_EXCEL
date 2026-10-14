<#
.SYNOPSIS
  Wgrywa kod VBA z src\ do PW_integration_imported.xlsm:
  - moduły standardowe (modConfig, Module1-3) są podmieniane w całości,
  - kod arkuszy (Sheet1, Sheet4) i formularzy (frmAttributes, frmNavigator) jest podmieniany w środku
    (układ graficzny formularzy zostaje bez zmian).
.NOTES
  Wymaga Windows + Excel oraz włączonej opcji:
  Plik > Opcje > Centrum zaufania > Ustawienia makr > "Ufaj dostępowi do modelu obiektowego projektu VBA".
  Uruchom na KOPII pliku. Po imporcie zapisz skoroszyt.
#>
param([string]$Workbook = (Join-Path $PSScriptRoot 'PW_integration_imported.xlsm'))
$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP 'pw_vba_import'
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
New-Item $tmp -ItemType Directory | Out-Null
$ansi = [System.Text.Encoding]::GetEncoding(1250)   # VBE oczekuje kodowania ANSI, nie UTF-8
function Read-Src($name) { (Get-Content (Join-Path $PSScriptRoot "src\$name") -Raw -Encoding UTF8) -replace "`r?`n", "`r`n" }

$modules = 'modConfig', 'Module1', 'Module2', 'Module3'
foreach ($m in $modules) { [IO.File]::WriteAllText((Join-Path $tmp "$m.bas"), (Read-Src "$m.bas"), $ansi) }

# komponenty z kodem "w środku": plik źródłowy -> nazwa komponentu w skoroszycie
$inplace = @{ 'Sheet1.cls' = 'Sheet1'; 'Sheet4.cls' = 'Sheet4'; 'frmAttributes.frm' = 'frmAttributes'; 'frmNavigator.frm' = 'frmNavigator' }

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Workbook)
    $vbp = $wb.VBProject
    foreach ($m in $modules) {
        $c = $null; try { $c = $vbp.VBComponents.Item($m) } catch {}
        if ($c) { $vbp.VBComponents.Remove($c) }
    }
    foreach ($m in $modules) { [void]$vbp.VBComponents.Import((Join-Path $tmp "$m.bas")) }
    foreach ($file in $inplace.Keys) {
        $code = (Read-Src $file) -split "`r`n" | Where-Object { $_ -notmatch '^Attribute VB_' } 
        $cm = $vbp.VBComponents.Item($inplace[$file]).CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString(($code -join "`r`n"))
    }
    $wb.Save()
    $wb.Close($true)
    Write-Host 'OK: kod VBA zaimportowany.' -ForegroundColor Green
} finally { $xl.Quit() }
