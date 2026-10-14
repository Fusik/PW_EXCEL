# Mierzy, co spowalnia pobieranie z ProjectWise. Uruchom:
#   powershell -ExecutionPolicy Bypass -File tools\Measure-PWFetch.ps1 -Link "pw:\\serwer:Datasource\Documents\Projekt\Folder\"
# Uwaga: wariant 1 trwa tyle, co zwykla aktualizacja (kilka minut), calosc ok. 10 min.
param([Parameter(Mandatory)][string]$Link)
$ErrorActionPreference = 'Stop'
$s = $Link.Replace('/', '\').Trim()
if ($s -match '^(?i)pw:') { $s = $s.Substring(3) }
$s = $s.TrimStart('\').Replace(' (Working folder)', '')
$i = $s.IndexOf('\Documents\', [StringComparison]::OrdinalIgnoreCase)
if ($i -lt 0) { throw 'Link bez \Documents\' }
$ds = $s.Substring(0, $i); $path = $s.Substring($i + 11).Trim('\')
Import-Module PWPS_DAB -Force | Out-Null
New-PWLogin -DatasourceName $ds -BentleyIMS -NonAdminLogin | Out-Null
try {
    $f = @(Get-PWFolders -FolderPath $path -ErrorAction SilentlyContinue)
    if ($f.Count -eq 0) { $f = @(Get-PWFolders -FolderPath ('\' + $path) -ErrorAction SilentlyContinue) }
    if ($f.Count -eq 0) { throw 'Nie znaleziono folderu' }
    $id = $f[0].ProjectID
    $tests = [ordered]@{
        '1 GetAttributes + GetVersionsToo (obecnie)' = @{ GetAttributes = $true; GetVersionsToo = $true }
        '2 GetAttributes (bez rewizji)'              = @{ GetAttributes = $true }
        '3 GetVersionsToo (bez atrybutow)'           = @{ GetVersionsToo = $true }
        '4 bez opcji'                                = @{}
        '5 jak 1 + ReturnDataTable'                  = @{ GetAttributes = $true; GetVersionsToo = $true; ReturnDataTable = $true }
    }
    foreach ($k in $tests.Keys) {
        $p = $tests[$k]
        $n = 0
        $t = Measure-Command { $r = @(Get-PWDocumentsBySearch -FolderID $id @p -WarningAction SilentlyContinue -ErrorAction SilentlyContinue); $n = $r.Count }
        '{0,-45} dokumentow: {1,5}  czas: {2:mm\:ss}' -f $k, $n, $t
    }
} finally { Undo-PWLogin | Out-Null }
Read-Host 'Enter, aby zamknac'
