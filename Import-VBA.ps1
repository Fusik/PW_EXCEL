<#
.SYNOPSIS
  Wgrywa kod VBA z src\ do PW_integration.xlsm:
  - moduły standardowe (modConfig, Module1-3) są podmieniane w całości,
  - kod arkuszy (Sheet1, Sheet4) i formularzy (frmAttributes, frmNavigator) jest podmieniany w środku
    (układ graficzny formularzy zostaje bez zmian).
.NOTES
  Wymaga Windows + Excel oraz włączonej opcji:
  Plik > Opcje > Centrum zaufania > Ustawienia makr > "Ufaj dostępowi do modelu obiektowego projektu VBA".
  Uruchom na KOPII pliku. Po imporcie zapisz skoroszyt.
#>
param([string]$Workbook = (Join-Path $PSScriptRoot 'PW_integration.xlsm'))
$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP 'pw_vba_import'
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
New-Item $tmp -ItemType Directory | Out-Null
$ansi = [System.Text.Encoding]::GetEncoding(1250)   # VBE oczekuje kodowania ANSI, nie UTF-8
function Read-Src($name) { (Get-Content (Join-Path $PSScriptRoot "src\$name") -Raw -Encoding UTF8) -replace "`r?`n", "`r`n" }

$modules = 'modConfig', 'modViews', 'modBench', 'Module1', 'Module2', 'Module3'
foreach ($m in $modules) { [IO.File]::WriteAllText((Join-Path $tmp "$m.bas"), (Read-Src "$m.bas"), $ansi) }

# komponenty z kodem "w środku": plik źródłowy -> nazwa komponentu w skoroszycie
$inplace = @{ 'ThisWorkbook.cls' = 'ThisWorkbook'; 'Sheet1.cls' = 'Sheet1'; 'Sheet4.cls' = 'Sheet4'; 'frmAttributes.frm' = 'frmAttributes'; 'frmNavigator.frm' = 'frmNavigator' }

# Zdejmij blokady, ktore uniemozliwiaja zapis: znacznik "pobrane z internetu" i atrybut tylko-do-odczytu
$Workbook = (Resolve-Path $Workbook).Path
try { Unblock-File -Path $Workbook } catch {}
try { Set-ItemProperty -Path $Workbook -Name IsReadOnly -Value $false } catch {}
if (Get-Process excel -ErrorAction SilentlyContinue) { Write-Host 'UWAGA: dzialaja procesy Excel - zamknij Excela (Menedzer zadan), jesli plik jest w nim otwarty.' -ForegroundColor Yellow }

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Workbook, 0, $false)
    if ($wb.ReadOnly) { throw 'Plik otwarty tylko do odczytu (jest otwarty w innym Excelu albo w widoku chronionym). Zamknij go i uruchom ponownie.' }
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
    $wb.Close($false)
    Write-Host 'OK: kod VBA zaimportowany.' -ForegroundColor Green
} finally {
    try { if ($wb) { $wb.Close($false) } } catch {}
    $xl.Quit()
}
