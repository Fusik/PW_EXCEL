Attribute VB_Name = "modBench"
Option Explicit

' =======================================================================
' TEST SZYBKOSCI POBIERANIA Z PW
' Pobiera pierwszy aktywny folder z arkusza Foldery kilkoma sposobami i zapisuje czasy w arkuszu Test_PW.
' Uwaga: pierwszy wariant trwa tyle co zwykla aktualizacja (kilka minut); calosc 1-2 min.
' =======================================================================
Public Sub TestSzybkosciPW()
    Dim wsF As Worksheet, r As Long, lastR As Long, lnk As String, found As String
    Set wsF = ThisWorkbook.Sheets(SH_FOLDERS)
    lastR = wsF.Cells(wsF.Rows.Count, FOLDERS_LINK_COL).End(xlUp).Row
    For r = FOLDERS_FIRST_ROW To lastR
        lnk = Trim(CStr(wsF.Cells(r, FOLDERS_LINK_COL).Value))
        If UCase(Left(lnk, 3)) = "PW:" And UCase(Trim(CStr(wsF.Cells(r, FOLDERS_ACTIVE_COL).Value))) <> "NIE" Then found = lnk: Exit For
    Next r
    If found = "" Then MsgBox "Brak aktywnego linku w arkuszu " & SH_FOLDERS & ".", vbExclamation: Exit Sub
    If ThisWorkbook.Path = "" Then MsgBox "Zapisz najpierw plik na dysku.", vbCritical: Exit Sub
    If MsgBox("Test pobierze folder z pierwszego aktywnego linku kilka razy roznymi sposobami (kilka minut):" & vbCrLf & vbCrLf & found & vbCrLf & vbCrLf & "Kontynuowac?", vbYesNo + vbQuestion, "Test szybkosci PW") <> vbYes Then Exit Sub
    
    Dim dsName As String, fp As String
    If Not ParseFolderLink(found, dsName, fp) Then MsgBox "Nie rozpoznano linku folderu.", vbExclamation: Exit Sub
    
    Dim fso As Object, psPath As String, csvPath As String, psFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    psPath = ThisWorkbook.Path & "\PW_Bench.ps1"
    csvPath = ThisWorkbook.Path & "\PW_Bench.csv"
    On Error Resume Next
    fso.DeleteFile csvPath
    On Error GoTo 0
    Set psFile = fso.CreateTextFile(psPath, True, True)
    psFile.WriteLine "$ErrorActionPreference = 'Stop'"
    psFile.WriteLine "$dsName = '" & Replace(dsName, "'", "''") & "'"
    psFile.WriteLine "$path = '" & Replace(fp, "'", "''") & "'"
    psFile.WriteLine "$loggedIn = $false"
    psFile.WriteLine "$res = New-Object System.Collections.ArrayList; $dt = $null"
    psFile.WriteLine "try {"
    psFile.WriteLine "    Import-Module PWPS_DAB -Force | Out-Null"
    psFile.WriteLine "    Write-Host ('Logowanie do PW (' + $dsName + ')...') -ForegroundColor Cyan"
    psFile.WriteLine "    New-PWLogin -DatasourceName $dsName -BentleyIMS -NonAdminLogin | Out-Null"
    psFile.WriteLine "    $loggedIn = $true"
    psFile.WriteLine "    $f = @(Get-PWFolders -FolderPath $path -ErrorAction SilentlyContinue)"
    psFile.WriteLine "    if ($f.Count -eq 0) { $f = @(Get-PWFolders -FolderPath ('' + $path) -ErrorAction SilentlyContinue) }"
    psFile.WriteLine "    if ($f.Count -eq 0) { throw 'Nie znaleziono folderu w ProjectWise.' }"
    psFile.WriteLine "    $id = $f[0].ProjectID"
    psFile.WriteLine "    $tests = [ordered]@{"
    psFile.WriteLine "        '3 rewizje, bez atrybutow'             = @{ GetVersionsToo = $true }"
    psFile.WriteLine "        '4 bez opcji (sama lista)'             = @{}"
    psFile.WriteLine "        '5 atrybuty + rewizje + ReturnDataTable' = @{ GetAttributes = $true; GetVersionsToo = $true; ReturnDataTable = $true }"
    psFile.WriteLine "        '6 atrybuty + ReturnDataTable'         = @{ GetAttributes = $true; ReturnDataTable = $true }"
    psFile.WriteLine "    }"
    psFile.WriteLine "    foreach ($k in $tests.Keys) {"
    psFile.WriteLine "        $p = $tests[$k]; $n = 0; $err = ''; if ($null -eq $dt) { $dt = $null }"
    psFile.WriteLine "        Write-Host ('Test: ' + $k + ' ...') -ForegroundColor Yellow"
    psFile.WriteLine "        $sw = [System.Diagnostics.Stopwatch]::StartNew()"
    psFile.WriteLine "        try {"
    psFile.WriteLine "            $r = @(Get-PWDocumentsBySearch -FolderID $id @p -WarningAction SilentlyContinue -ErrorAction Stop)"
    psFile.WriteLine "            if ($r.Count -eq 1 -and $r[0] -is [System.Data.DataTable]) { $n = $r[0].Rows.Count; if (-not $dt -and $r[0].Rows.Count -gt 0) { $dt = $r[0] } } else { $n = $r.Count }"
    psFile.WriteLine "        } catch { $err = $_.Exception.Message.Replace('|', '-') }"
    psFile.WriteLine "        $sw.Stop()"
    psFile.WriteLine "        Write-Host ('   dokumentow: ' + $n + '  czas: ' + $sw.Elapsed.ToString('mm\:ss') + $(if ($err) { '  BLAD: ' + $err } else { '' })) -ForegroundColor Green"
    psFile.WriteLine "        [void]$res.Add([PSCustomObject]@{ 'Wariant' = $k; 'Dokumentow' = $n; 'Sekundy' = [math]::Round($sw.Elapsed.TotalSeconds, 1); 'Uwagi' = $err })"
    psFile.WriteLine "    }"
    psFile.WriteLine "    if ($dt) { $row = $dt.Rows[0]; foreach ($c in $dt.Columns) { $v = $row[$c.ColumnName]; if ($v -is [DBNull]) { $v = '' }; if ($v -is [DateTime]) { $v = $v.ToString('yyyy-MM-dd') }; $v = ([string]$v).Replace('|', '-').Replace([string][char]10, ' ').Replace([string][char]13, ''); [void]$res.Add([PSCustomObject]@{ 'Wariant' = ('[KOLUMNA] ' + $c.ColumnName); 'Dokumentow' = 0; 'Sekundy' = 0; 'Uwagi' = ($c.DataType.Name + ' | ' + $v) }) } }"
    psFile.WriteLine "    if ($dt) {"
    psFile.WriteLine "        $fn = [string]$dt.Rows[0]['DocumentFileName']"
    psFile.WriteLine "        $m = 0; $pc = 0; $err2 = ''"
    psFile.WriteLine "        $sw2 = [System.Diagnostics.Stopwatch]::StartNew()"
    psFile.WriteLine "        try {"
    psFile.WriteLine "            $o = @(Get-PWDocumentsBySearch -FolderID $id -FileName $fn -GetAttributes -WarningAction SilentlyContinue -ErrorAction Stop)"
    psFile.WriteLine "            $m = $o.Count"
    psFile.WriteLine "            if ($o.Count -gt 0) { $pc = @($o[0].psobject.Properties).Count }"
    psFile.WriteLine "        } catch { $err2 = $_.Exception.Message.Replace('|', '-') }"
    psFile.WriteLine "        $sw2.Stop()"
    psFile.WriteLine "        [void]$res.Add([PSCustomObject]@{ 'Wariant' = '7 jeden plik po nazwie, z atrybutami'; 'Dokumentow' = $m; 'Sekundy' = [math]::Round($sw2.Elapsed.TotalSeconds, 1); 'Uwagi' = ('wlasciwosci obiektu: ' + $pc + ' ' + $err2) })"
    psFile.WriteLine "    }"
    psFile.WriteLine "    $skip = '^(Verbose|Debug|ErrorAction|WarningAction|InformationAction|ErrorVariable|WarningVariable|InformationVariable|OutVariable|OutBuffer|PipelineVariable)$'"
    psFile.WriteLine "    foreach ($c in @(Get-Command -Module PWPS_DAB | Where-Object { $_.Name -match '^Get-PW' -and $_.Name -match 'Attribute|Environment|DataTable|Document|Search' })) {"
    psFile.WriteLine "        $ps = ($c.Parameters.Keys | Where-Object { $_ -notmatch $skip }) -join ', '"
    psFile.WriteLine "        [void]$res.Add([PSCustomObject]@{ 'Wariant' = ('[CMD] ' + $c.Name); 'Dokumentow' = 0; 'Sekundy' = 0; 'Uwagi' = $ps })"
    psFile.WriteLine "    }"
    psFile.WriteLine "    foreach ($fname in 'Get-PWDocumentsBySearchWithReturnColumns', 'Get-PWDocumentsBySearchExtended') {"
    psFile.WriteLine "        $g = Get-Command $fname -ErrorAction SilentlyContinue"
    psFile.WriteLine "        if (-not $g) { [void]$res.Add([PSCustomObject]@{ 'Wariant' = ('8 ' + $fname); 'Dokumentow' = 0; 'Sekundy' = 0; 'Uwagi' = 'brak cmdletu w tej wersji modulu' }); continue }"
    psFile.WriteLine "        $cp = @($g.Parameters.Keys | Where-Object { $_ -match 'Column' }) | Select-Object -First 1"
    psFile.WriteLine "        $m = 0; $note = ''; $err3 = ''"
    psFile.WriteLine "        $sw3 = [System.Diagnostics.Stopwatch]::StartNew()"
    psFile.WriteLine "        try {"
    psFile.WriteLine "            $a3 = @{ FolderID = $id }"
    psFile.WriteLine "            if ($cp) { $a3[$cp] = @('DocumentName', 'UK_project_Reference', 'Status_Code', 'Revision_number') }"
    psFile.WriteLine "            $o3 = @(& $fname @a3 -WarningAction SilentlyContinue -ErrorAction Stop)"
    psFile.WriteLine "            if ($o3.Count -eq 1 -and $o3[0] -is [System.Data.DataTable]) { $o3 = @($o3[0].Rows) }"
    psFile.WriteLine "            $m = $o3.Count"
    psFile.WriteLine "            if ($m -gt 0) {"
    psFile.WriteLine "                $f3 = $o3[0]"
    psFile.WriteLine "                if ($f3 -is [System.Data.DataRow]) { $nm = @($f3.Table.Columns | ForEach-Object { $_.ColumnName }) } else { $nm = @($f3.psobject.Properties | ForEach-Object { $_.Name }) }"
    psFile.WriteLine "                $note = 'parametr kolumn: ' + $cp + '; zwrocone: ' + (($nm | Select-Object -First 40) -join ', ')"
    psFile.WriteLine "            }"
    psFile.WriteLine "        } catch { $err3 = $_.Exception.Message.Replace('|', '-') }"
    psFile.WriteLine "        $sw3.Stop()"
    psFile.WriteLine "        [void]$res.Add([PSCustomObject]@{ 'Wariant' = ('8 ' + $fname); 'Dokumentow' = $m; 'Sekundy' = [math]::Round($sw3.Elapsed.TotalSeconds, 1); 'Uwagi' = ($note + ' ' + $err3) })"
    psFile.WriteLine "    }"
    psFile.WriteLine "    $res | Export-Csv -Path '" & csvPath & "' -NoTypeInformation -Delimiter '|' -Encoding UTF8"
    psFile.WriteLine "} catch {"
    psFile.WriteLine "    Write-Host ('BLAD: ' + $_.Exception.Message) -ForegroundColor Red"
    psFile.WriteLine "    Read-Host 'Nacisnij Enter, aby zamknac okno'"
    psFile.WriteLine "} finally {"
    psFile.WriteLine "    if ($loggedIn) { try { Undo-PWLogin | Out-Null } catch {} }"
    psFile.WriteLine "}"
    psFile.WriteLine "Start-Sleep -Seconds 1"
    psFile.Close
    
    CreateObject("WScript.Shell").Run "powershell.exe -ExecutionPolicy Bypass -File """ & psPath & """", 1, True
    
    If Not fso.FileExists(csvPath) Then MsgBox "Test nie zapisal wynikow. Sprawdz okno PowerShell.", vbCritical: Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Test_PW")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Test_PW"
    End If
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Wariant", "Dokumentow", "Sekundy", "Uwagi")
    StyleHeader ws.Range("A1:D1")
    
    Dim stream As Object, rowsArr() As String, cols() As String, outR As Long, best As String, bestSec As Double
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile csvPath
    rowsArr = Split(Replace(Replace(stream.ReadText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    stream.Close
    outR = 2
    For r = 1 To UBound(rowsArr)
        If Trim(rowsArr(r)) <> "" Then
            cols = Split(Replace(rowsArr(r), """", ""), "|")
            If UBound(cols) >= 3 Then
                ws.Cells(outR, 1).Value = cols(0)
                ws.Cells(outR, 2).Value = Val(cols(1))
                ws.Cells(outR, 3).Value = Val(cols(2))
                ws.Cells(outR, 4).NumberFormat = "@": ws.Cells(outR, 4).Value = cols(3)
                If UBound(cols) > 3 Then ws.Cells(outR, 4).Value = cols(3) & " | " & cols(4)
                outR = outR + 1
            End If
        End If
    Next r
    ws.Cells(outR + 1, 1).Value = "Folder: " & found
    ws.Cells(outR + 2, 1).Value = "Test: " & Format(Now, "yyyy-mm-dd hh:nn")
    ws.Columns("A:D").AutoFit
    ws.Activate
    On Error Resume Next
    fso.DeleteFile psPath
    fso.DeleteFile csvPath
    On Error GoTo 0
    MsgBox "Gotowe - wyniki w arkuszu Test_PW. Wyslij mi je (zrzut ekranu).", vbInformation
End Sub
