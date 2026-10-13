Attribute VB_Name = "Module3"
' =======================================================================
' ZBADANIE UKRYTYCH ATRYBUTÓW (PROPERTY NAMES)
' Pobiera jeden plik ze wskazanej ścieżki i bada jego surowe atrybuty systemowe,
' aby ułatwić dodawanie nowych kolumn w Sheet_3.
' =======================================================================
Sub ZbadajAtrybutyPW()
    Dim ws2 As Worksheet, wsDiag As Worksheet
    Dim folderPath As String, basePath As String
    Dim psScriptPath As String, csvOutputPath As String
    Dim fso As Object, wsh As Object
    Dim execResult As Long
    Dim r As Long, cols() As String, cleanRow As String
    
    If ThisWorkbook.Path = "" Then
        MsgBox "Zapisz najpierw plik Excel na dysku!", vbCritical
        Exit Sub
    End If
    
    ' Odczytanie testowej ścieżki wpisanej w Sheet_3 (Komórka B2)
    On Error Resume Next
    Set ws2 = ThisWorkbook.Sheets(SH_ATTRS)
    On Error GoTo 0
    
    If ws2 Is Nothing Then
        MsgBox "Brak arkusza '" & SH_ATTRS & "'.", vbExclamation
        Exit Sub
    End If
    
    folderPath = ws2.Range(ATTRS_SAMPLE_FOLDER).Value
    If Trim(folderPath) = "" Then
        MsgBox "Brak folderu w arkuszu " & SH_ATTRS & " (komórka " & ATTRS_SAMPLE_FOLDER & "). Wpisz tam ścieżkę do folderu PW, aby pobrać jego atrybuty.", vbExclamation
        Exit Sub
    End If
    
    basePath = ThisWorkbook.Path
    psScriptPath = basePath & "\PW_Diag.ps1"
    csvOutputPath = basePath & "\PW_Diag.csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim dsName As String, fp As String
    If Not ParseFolderLink(folderPath, dsName, fp) Then
        MsgBox "Nie rozpoznano linku folderu próbnego. Oczekiwany format: pw:\\serwer:Datasource\Documents\Projekt\Folder\ (zobacz arkusz Instrukcja).", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    fso.DeleteFile csvOutputPath ' stary wynik nie moze udawac nowego
    On Error GoTo 0
    
    Dim psFile As Object
    Set psFile = fso.CreateTextFile(psScriptPath, True, True)
    psFile.WriteLine "$ErrorActionPreference = 'Stop'"
    psFile.WriteLine "$dsName = '" & Replace(dsName, "'", "''") & "'"
    psFile.WriteLine "$folderPath = '" & Replace(fp, "'", "''") & "'"
    psFile.WriteLine "$loggedIn = $false"
    psFile.WriteLine "try {"
    psFile.WriteLine "    Import-Module PWPS_DAB -Force | Out-Null"
    psFile.WriteLine "    Write-Host ('Logowanie do PW (' + $dsName + ')...') -ForegroundColor Cyan"
    psFile.WriteLine "    New-PWLogin -DatasourceName $dsName -BentleyIMS -NonAdminLogin | Out-Null"
    psFile.WriteLine "    $loggedIn = $true"
    psFile.WriteLine "    Write-Host ('Szukam folderu: ' + $folderPath) -ForegroundColor Yellow"
    psFile.WriteLine "    $fl = @(Get-PWFolders -FolderPath $folderPath -ErrorAction SilentlyContinue)"
    psFile.WriteLine "    if ($fl.Count -eq 0) { $fl = @(Get-PWFolders -FolderPath ('\' + $folderPath) -ErrorAction SilentlyContinue) }"
    psFile.WriteLine "    if ($fl.Count -eq 0) { $fl = @(Get-PWFolders -FolderName $folderPath.Split('\')[-1] -ErrorAction SilentlyContinue) }"
    psFile.WriteLine "    if ($fl.Count -eq 0) { throw 'Nie znaleziono folderu w ProjectWise. Sprawdz link folderu probnego w arkuszu Atrybuty.' }"
    psFile.WriteLine "    Write-Host 'Pobieram dokumenty z atrybutami (moze chwile potrwac)...' -ForegroundColor Yellow"
    psFile.WriteLine "    $doc = Get-PWDocumentsBySearch -FolderID $fl[0].ProjectID -GetAttributes -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object -First 1"
    psFile.WriteLine "    if (-not $doc) { throw 'Folder nie zawiera dokumentow - wskaz folder z plikami.' }"
    psFile.WriteLine "    $out = New-Object System.Collections.ArrayList"
    psFile.WriteLine "    $seen = @{}"
    psFile.WriteLine "    function Clean($v) { if ($v -is [array]) { $v = ($v -join ', ') }; if ($v -is [DateTime]) { $v = $v.ToString('yyyy-MM-dd') }; if ($v -is [string]) { $v = $v.Replace('|', '-').Replace([string][char]10, ' ').Replace([string][char]13, '') }; return $v }"
    psFile.WriteLine "    foreach ($p in $doc.psobject.Properties) {"
    psFile.WriteLine "        [void]$out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = $p.Name; 'Wartosc' = (Clean $p.Value) })"
    psFile.WriteLine "        $seen[$p.Name] = $true"
    psFile.WriteLine "    }"
    psFile.WriteLine "    foreach ($p in $doc.psobject.Properties) {"
    psFile.WriteLine "        $v = $p.Value"
    psFile.WriteLine "        if ($null -eq $v -or $v -is [string] -or $v -is [ValueType] -or $v -is [array]) { continue }"
    psFile.WriteLine "        if ($v -is [System.Collections.IDictionary]) { $pairs = @($v.Keys | Select-Object -First 200 | ForEach-Object { @{ N = [string]$_; V = $v[$_] } }) }"
    psFile.WriteLine "        else { $pairs = @($v.psobject.Properties | Select-Object -First 200 | ForEach-Object { @{ N = $_.Name; V = $_.Value } }) }"
    psFile.WriteLine "        foreach ($q in $pairs) {"
    psFile.WriteLine "            if (-not $seen.ContainsKey($q.N)) { [void]$out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = $q.N; 'Wartosc' = (Clean $q.V) }); $seen[$q.N] = $true }"
    psFile.WriteLine "        }"
    psFile.WriteLine "    }"
    psFile.WriteLine "    $uName = @($doc.DocumentUpdaterName, $doc.FileUpdaterName, $doc.UpdaterName, $doc.Updater, $doc.CreatorName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1"
    psFile.WriteLine "    [void]$out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = 'Updater'; 'Wartosc' = $(if ($uName) { $uName } else { 'Unknown' }) })"
    psFile.WriteLine "    $uDate = @($doc.DocumentUpdateDate, $doc.FileUpdateDate, $doc.UpdateDate, $doc.Updated, $doc.CreateDate) | Where-Object { $_ } | Select-Object -First 1"
    psFile.WriteLine "    [void]$out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = 'Updated'; 'Wartosc' = (Clean $uDate) })"
    psFile.WriteLine "    $out | Export-Csv -Path '" & csvOutputPath & "' -NoTypeInformation -Delimiter '|' -Encoding UTF8"
    psFile.WriteLine "    Write-Host ('Zapisano ' + $out.Count + ' atrybutow. Plik: ' + $doc.Name) -ForegroundColor Green"
    psFile.WriteLine "} catch {"
    psFile.WriteLine "    Write-Host ('BLAD: ' + $_.Exception.Message) -ForegroundColor Red"
    psFile.WriteLine "    Read-Host 'Nacisnij Enter, aby zamknac okno'"
    psFile.WriteLine "} finally {"
    psFile.WriteLine "    if ($loggedIn) { try { Undo-PWLogin | Out-Null } catch {} }"
    psFile.WriteLine "}"
    psFile.WriteLine "Start-Sleep -Seconds 2"
    psFile.Close
    
    ' Wykonanie skryptu w tle
    Set wsh = CreateObject("WScript.Shell")
    execResult = wsh.Run("powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & """", 1, True)
    
    ' Zrzut pliku wynikowego CSV do nowego arkusza Excel
    If fso.FileExists(csvOutputPath) Then
        On Error Resume Next
        Set wsDiag = ThisWorkbook.Sheets(SH_DIAG)
        If wsDiag Is Nothing Then
            Set wsDiag = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            wsDiag.Name = SH_DIAG
        End If
        On Error GoTo 0
        
        wsDiag.Cells.Clear
        wsDiag.Cells(1, 1).Value = "Nazwa atrybutu (skopiuj do kolumny A w arkuszu " & SH_ATTRS & ")"
        wsDiag.Cells(1, 2).Value = "Przykładowa wartość (z 1 pliku)"
        StyleHeader wsDiag.Range("A1:B1")
        
        Dim stream As Object, csvData As String, rowsArr() As String
        Set stream = CreateObject("ADODB.Stream")
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile csvOutputPath
        If Not stream.EOS Then
            csvData = stream.ReadText
            csvData = Replace(csvData, vbCrLf, vbLf)
            csvData = Replace(csvData, vbCr, vbLf)
            rowsArr = Split(csvData, vbLf)
            
            Dim outR As Long
            outR = 2
            For r = 1 To UBound(rowsArr)
                If Trim(rowsArr(r)) <> "" Then
                    cleanRow = Replace(rowsArr(r), """", "")
                    cols = Split(cleanRow, "|")
                    If UBound(cols) >= 1 Then
                        wsDiag.Cells(outR, 1).Value = Trim(cols(0))
                        wsDiag.Cells(outR, 2).Value = Trim(cols(1))
                        
                        ' Wyróżnienie naszych sztucznych atrybutów
                        If Trim(cols(0)) = "Updater" Or Trim(cols(0)) = "Updated" Then
                            wsDiag.Range(wsDiag.Cells(outR, 1), wsDiag.Cells(outR, 2)).Interior.Color = CLR_CHANGED
                            wsDiag.Range(wsDiag.Cells(outR, 1), wsDiag.Cells(outR, 2)).Font.Bold = True
                        End If
                        
                        outR = outR + 1
                    End If
                End If
            Next r
        End If
        stream.Close
        
        wsDiag.Columns("A:B").AutoFit
        wsDiag.Activate
        MsgBox "Diagnostyka zakończona! Sprawdź arkusz '" & SH_DIAG & "'.", vbInformation
    Else
        MsgBox "Nie udało się wygenerować pliku CSV. Sprawdź okno PowerShell.", vbCritical
    End If
    
    ' Sprzątanie po sobie
    On Error Resume Next
    fso.DeleteFile psScriptPath
    fso.DeleteFile csvOutputPath
    On Error GoTo 0
End Sub

Private Function ParseFolderLink(ByVal link As String, ByRef ds As String, ByRef fp As String) As Boolean
    Dim s As String, p As Long
    s = Replace(Trim(link), "/", "\")
    If UCase(Left(s, 3)) = "PW:" Then s = Mid(s, 4)
    Do While Left(s, 1) = "\"
        s = Mid(s, 2)
    Loop
    s = Replace(s, " (Working folder)", "", 1, -1, vbTextCompare)
    p = InStr(1, s, "\Documents\", vbTextCompare)
    If p > 0 Then
        ds = Left(s, p - 1)
        fp = Mid(s, p + 11)
    Else
        p = InStr(1, s, "\")
        If p = 0 Then Exit Function
        ds = Left(s, p - 1)
        fp = Mid(s, p + 1)
    End If
    Do While Right(fp, 1) = "\"
        fp = Left(fp, Len(fp) - 1)
    Loop
    Do While Left(fp, 1) = "\"
        fp = Mid(fp, 2)
    Loop
    ParseFolderLink = (ds <> "" And fp <> "")
End Function
