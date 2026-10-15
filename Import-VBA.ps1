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

# VBA nie przyjmuje linii dluzszych niz 1023 znaki - taki modul psuje projekt i skoroszyt nie daje sie zapisac
foreach ($f in Get-ChildItem (Join-Path $PSScriptRoot 'src') -File) {
    $n = 0
    foreach ($l in Get-Content $f.FullName -Encoding UTF8) { $n++; if ($l.Length -gt 1000) { throw ('Za dluga linia VBA (' + $l.Length + ' znakow): ' + $f.Name + ':' + $n) } }
}

$modules = 'modConfig', 'modViews', 'modTheme', 'modBench', 'modDownload', 'Module1', 'Module2', 'Module3'
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
    # formularz pobierania: kontrolki tworzone kodem (nie wymaga pliku .frx)
    $fc = $null; try { $fc = $vbp.VBComponents.Item('frmPobierz') } catch {}
    if ($fc) { $vbp.VBComponents.Remove($fc) }
    $fc = $vbp.VBComponents.Add(3)
    $fc.Name = 'frmPobierz'
    $fc.Properties.Item('Caption').Value = 'Download from ProjectWise'
    $fc.Properties.Item('Width').Value = 520
    $fc.Properties.Item('Height').Value = 420
    function Add-Ctl($type, $name, $l, $t, $w, $h, $cap) {
        $c = $fc.Designer.Controls.Add($type, $name)
        $c.Left = $l; $c.Top = $t; $c.Width = $w; $c.Height = $h
        if ($cap) { $c.Caption = $cap }
        return $c
    }
    [void](Add-Ctl 'Forms.Label.1' 'lblInfo' 8 6 490 16 'Paste PW links (pw://...), one per line:')
    $tl = Add-Ctl 'Forms.TextBox.1' 'txtLinks' 8 24 490 250 $null
    $tl.MultiLine = $true; $tl.ScrollBars = 3; $tl.EnterKeyBehavior = $true; $tl.WordWrap = $false
    [void](Add-Ctl 'Forms.Label.1' 'lblFolder' 8 282 490 16 'Target folder (e.g. D:\PW_files):')
    [void](Add-Ctl 'Forms.TextBox.1' 'txtFolder' 8 300 490 20 $null)
    [void](Add-Ctl 'Forms.CommandButton.1' 'btnSel' 8 332 200 28 'Paste links from Dashboard selection')
    [void](Add-Ctl 'Forms.CommandButton.1' 'btnStart' 330 332 80 28 'DOWNLOAD')
    [void](Add-Ctl 'Forms.CommandButton.1' 'btnClose' 418 332 80 28 'Close')
    $fcode = (Read-Src 'frmPobierz.frm') -split "`r`n" | Where-Object { $_ -notmatch '^Attribute VB_' }
    $fc.CodeModule.AddFromString(($fcode -join "`r`n"))
    $build = (Select-String -Path (Join-Path $PSScriptRoot 'src\modConfig.bas') -Pattern 'BUILD_ID As String = "([^"]+)"').Matches[0].Groups[1].Value
    try {
        $wb.Save()
        $savedAs = $Workbook
    } catch {
        Write-Host ('Save() nie zadzialal: ' + $_.Exception.Message) -ForegroundColor Yellow
        Write-Host ('  plik: ' + $wb.FullName + ' | tylko-do-odczytu: ' + $wb.ReadOnly + ' | rozmiar na dysku: ' + (Get-Item $Workbook).Length) -ForegroundColor Yellow
        $alt = Join-Path (Split-Path $Workbook) ((Get-Item $Workbook).BaseName + '_build_' + $build + '.xlsm')
        Write-Host ('Probuje zapisac kopie: ' + $alt) -ForegroundColor Yellow
        $wb.SaveAs($alt, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        $savedAs = $alt
    }
    $wb.Close($false)
    Write-Host ('OK: kod VBA zaimportowany. Build ' + $build + ' -> ' + $savedAs) -ForegroundColor Green
} finally {
    try { if ($wb) { $wb.Close($false) } } catch {}
    $xl.Quit()
}
