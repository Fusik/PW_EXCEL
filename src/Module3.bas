Attribute VB_Name = "Module3"
' =======================================================================
' WYKRYWANIE ATRYBUTOW PW (diagnostyka) + AUTOMAT PO WKLEJENIU LINKU
' DiscoverAttributes: loguje sie do PW, bierze 1 dokument z folderu, zapisuje jego
' wlasciwosci w arkuszu Diagnostyka i DOPISUJE nowe atrybuty do arkusza Atrybuty.
' OnFolderLinkPasted: wywolywane po wklejeniu linku w arkuszu Foldery (ThisWorkbook).
' =======================================================================

' Przycisk "ZBADAJ ATRYBUTY": pierwszy aktywny folder z arkusza Foldery
Sub ZbadajAtrybutyPW()
    Dim wsF As Worksheet, r As Long, lastR As Long, lnk As String
    Set wsF = ThisWorkbook.Sheets(SH_FOLDERS)
    lastR = wsF.Cells(wsF.Rows.Count, FOLDERS_LINK_COL).End(xlUp).Row
    For r = FOLDERS_FIRST_ROW To lastR
        lnk = Trim(CStr(wsF.Cells(r, FOLDERS_LINK_COL).Value))
        If UCase(Left(lnk, 3)) = "PW:" And Not IsNo(wsF.Cells(r, FOLDERS_ACTIVE_COL).Value) Then
            DiscoverAttributes lnk
            Exit Sub
        End If
    Next r
    Notify "No active link in sheet " & SH_FOLDERS & " (column B). Paste a PW folder link."
End Sub

' Wywolywane po wklejeniu/zmianie linku w kolumnie B arkusza Foldery: wykryj atrybuty, potem pobierz dane
Public Sub OnFolderLinkPasted(ByVal Target As Range)
    Static busy As Boolean
    If busy Then Exit Sub
    If Target.Cells.Count > 20 Then Exit Sub        ' ochrona przed kasowaniem/wklejaniem calych kolumn
    If DownloadBusy() Then                          ' trwa pobieranie w tle - nie startuj kolejnego
        Notify "A ProjectWise download is already running. When it finishes, click UPDATE DATA to include the new link."
        Exit Sub
    End If
    busy = True
    Dim c As Range, ok As Boolean, lnk As String
    On Error GoTo Fail
    Application.EnableEvents = False
    For Each c In Target.Cells
        lnk = Trim(CStr(c.Value))
        If UCase(Left(lnk, 3)) = "PW:" And Not IsNo(c.Worksheet.Cells(c.Row, FOLDERS_ACTIVE_COL).Value) Then
            If DiscoverAttributes(lnk, True) Then ok = True
        End If
    Next c
    Application.EnableEvents = True
    If ok And AUTO_DOWNLOAD_AFTER_PASTE Then DownloadFromProjectWise_PWPS
    busy = False
    Exit Sub
Fail:
    Application.EnableEvents = True
    busy = False
    Notify "Error in the paste-link automation: " & Err.Description
End Sub

Public Function DiscoverAttributes(ByVal folderLink As String, Optional ByVal quiet As Boolean = False) As Boolean
    Dim wsDiag As Worksheet, basePath As String
    Dim psScriptPath As String, csvOutputPath As String
    Dim fso As Object, wsh As Object
    Dim r As Long, cols() As String, cleanRow As String
    
    If ThisWorkbook.Path = "" Then
        Notify "Save the Excel workbook to disk first."
        Exit Function
    End If
    
    basePath = ThisWorkbook.Path
    psScriptPath = basePath & "\PW_Diag.ps1"
    csvOutputPath = basePath & "\PW_Diag.csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim dsName As String, fp As String
    If Not ParseFolderLink(folderLink, dsName, fp) Then
        Notify "Folder link not recognised. Expected format: pw:\\server:Datasource\Documents\Project\Folder\ (see sheet Guide)."
        Exit Function
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
    psFile.WriteLine "    if ($fl.Count -eq 0) { $fl = @(Get-PWFolders -FolderPath ('' + $folderPath) -ErrorAction SilentlyContinue) }"
    psFile.WriteLine "    if ($fl.Count -eq 0) { $fl = @(Get-PWFolders -FolderName $folderPath.Split('')[-1] -ErrorAction SilentlyContinue) }"
    psFile.WriteLine "    if ($fl.Count -eq 0) { throw 'Nie znaleziono folderu w ProjectWise. Sprawdz link w arkuszu Foldery.' }"
    psFile.WriteLine "    Write-Host 'Pobieram dokumenty z atrybutami (moze chwile potrwac)...' -ForegroundColor Yellow"
    psFile.WriteLine "    $doc = $null"
    psFile.WriteLine "    $first = ''"
    psFile.WriteLine "    try {"
    psFile.WriteLine "        $r0 = @(Get-PWDocumentsBySearchExtended -FolderPath $folderPath -ReturnDataTable -WarningAction SilentlyContinue -ErrorAction Stop)"
    psFile.WriteLine "        if ($r0.Count -eq 1 -and $r0[0] -is [System.Data.DataTable] -and $r0[0].Rows.Count -gt 0) { $first = [string]$r0[0].Rows[0]['DocumentFileName'] }"
    psFile.WriteLine "    } catch { }"
    psFile.WriteLine "    if ($first) { $doc = Get-PWDocumentsBySearch -FolderID $fl[0].ProjectID -FileName $first -GetAttributes -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object -First 1 }"
    psFile.WriteLine "    if (-not $doc) {"
    psFile.WriteLine "        Write-Host 'Szybka droga nie zadzialala - pobieram caly folder z atrybutami (moze trwac kilka minut)...' -ForegroundColor Yellow"
    psFile.WriteLine "        $doc = Get-PWDocumentsBySearch -FolderID $fl[0].ProjectID -GetAttributes -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object -First 1"
    psFile.WriteLine "    }"
    psFile.WriteLine "    if (-not $doc) { throw 'Folder nie zawiera dokumentow - wklej link do folderu z plikami.' }"
    psFile.WriteLine "    $out = New-Object System.Collections.ArrayList"
    psFile.WriteLine "    $seen = @{}"
    psFile.WriteLine "    function Clean($v) { if ($v -is [array]) { $v = ($v -join ', ') }; if ($v -is [DateTime]) { $v = $v.ToString('yyyy-MM-dd') }; if ($v -is [string]) { $v = $v.Replace('|', '-').Replace([string][char]10, ' ').Replace([string][char]13, '') }; return $v }"
    psFile.WriteLine "    foreach ($p in $doc.psobject.Properties) {"
    psFile.WriteLine "        [void]$out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = $p.Name; 'Wartosc' = (Clean $p.Value); 'Typ' = 'SYSTEM' })"
    psFile.WriteLine "        $seen[$p.Name] = $true"
    psFile.WriteLine "    }"
    psFile.WriteLine "    foreach ($p in $doc.psobject.Properties) {"
    psFile.WriteLine "        if ($p.Name -ne 'CustomAttributes' -or $null -eq $p.Value) { continue }"
    psFile.WriteLine "        $v = $p.Value"
    psFile.WriteLine "        if ($v -is [System.Collections.IDictionary]) { $pairs = @($v.Keys | ForEach-Object { @{ N = [string]$_; V = $v[$_] } }) }"
    psFile.WriteLine "        else { $pairs = @($v.psobject.Properties | ForEach-Object { @{ N = $_.Name; V = $_.Value } }) }"
    psFile.WriteLine "        foreach ($q in $pairs) {"
    psFile.WriteLine "            if (-not $seen.ContainsKey($q.N)) { [void]$out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = $q.N; 'Wartosc' = (Clean $q.V); 'Typ' = 'ATRYBUT' }); $seen[$q.N] = $true }"
    psFile.WriteLine "        }"
    psFile.WriteLine "    }"
    psFile.WriteLine "    $out | Export-Csv -Path '" & csvOutputPath & "' -NoTypeInformation -Delimiter '|' -Encoding UTF8"
    psFile.WriteLine "    Write-Host ('Zapisano ' + $out.Count + ' atrybutow. Plik: ' + $doc.Name) -ForegroundColor Green"
    psFile.WriteLine "} catch {"
    psFile.WriteLine "    Write-Host ('BLAD: ' + $_.Exception.Message) -ForegroundColor Red"
    psFile.WriteLine "    Read-Host 'Nacisnij Enter, aby zamknac okno'"
    psFile.WriteLine "} finally {"
    psFile.WriteLine "    if ($loggedIn) { try { Undo-PWLogin | Out-Null } catch {} }"
    psFile.WriteLine "}"
    psFile.WriteLine "Start-Sleep -Seconds 1"
    psFile.Close
    
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & """", 1, True
    
    If fso.FileExists(csvOutputPath) Then
        On Error Resume Next
        Set wsDiag = ThisWorkbook.Sheets(SH_DIAG)
        If wsDiag Is Nothing Then
            Set wsDiag = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsDiag.Name = SH_DIAG
        End If
        On Error GoTo 0
        
        wsDiag.Cells.Clear
        wsDiag.Range("A1:C1").Value = Array("Nazwa atrybutu", "Przykladowa wartosc (z 1 pliku)", "Typ")
        StyleHeader wsDiag.Range("A1:C1")
        
        Dim stream As Object, csvData As String, rowsArr() As String, outR As Long
        Set stream = CreateObject("ADODB.Stream")
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile csvOutputPath
        If Not stream.EOS Then
            csvData = Replace(Replace(stream.ReadText, vbCrLf, vbLf), vbCr, vbLf)
            rowsArr = Split(csvData, vbLf)
            outR = 2
            For r = 1 To UBound(rowsArr)
                If Trim(rowsArr(r)) <> "" Then
                    cols = Split(Replace(rowsArr(r), """", ""), "|")
                    If UBound(cols) >= 2 Then
                        wsDiag.Cells(outR, 1).Value = Trim(cols(0))
                        wsDiag.Cells(outR, 2).NumberFormat = "@"
                        wsDiag.Cells(outR, 2).Value = Trim(cols(1))
                        wsDiag.Cells(outR, 3).Value = Trim(cols(2))
                        outR = outR + 1
                    End If
                End If
            Next r
        End If
        stream.Close
        wsDiag.Columns("A:C").AutoFit
        
        Dim added As Long
        added = MergeAttributes(wsDiag)
        DiscoverAttributes = True
        If Not quiet Then
            wsDiag.Activate
            Notify "Done. Added " & added & " new attributes to sheet " & SH_ATTRS & "."
        End If
    Else
        Notify "Could not discover attributes. Check the PowerShell window and the folder link."
    End If
    
    On Error Resume Next
    fso.DeleteFile psScriptPath
    fso.DeleteFile csvOutputPath
    On Error GoTo 0
End Function

' Dopisuje do arkusza Atrybuty nazwy z Diagnostyki, ktorych jeszcze tam nie ma (nic nie kasuje, nie rusza Twoich ustawien TAK/NIE)
Private Function MergeAttributes(wsDiag As Worksheet) As Long
    Dim wsA As Worksheet, known As Object, r As Long, lastD As Long, lastA As Long, nm As String, n As Long
    Dim typ As String, pobierac As String
    Set wsA = ThisWorkbook.Sheets(SH_ATTRS)
    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = 1
    lastA = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    If lastA < ATTRS_FIRST_ROW - 1 Then lastA = ATTRS_FIRST_ROW - 1
    For r = ATTRS_FIRST_ROW To lastA
        nm = Trim(CStr(wsA.Cells(r, 1).Value))
        If nm <> "" Then known(nm) = r
    Next r
    lastD = wsDiag.Cells(wsDiag.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastD
        nm = Trim(CStr(wsDiag.Cells(r, 1).Value))
        typ = Trim(CStr(wsDiag.Cells(r, 3).Value))
        If nm <> "" Then
            If known.Exists(nm) Then
                wsA.Cells(known(nm), ATTRS_EXAMPLE_COL).Value = "'" & CStr(wsDiag.Cells(r, 2).Value)   ' odswiez przyklad
            Else
                lastA = lastA + 1
                wsA.Cells(lastA, 1).Value = nm
                ' atrybuty projektu: domyslnie TAK; wlasciwosci systemowe: tylko przydatne, reszta NIE (do wlaczenia recznie)
                If typ = "ATRYBUT" Then
                    pobierac = "YES"
                ElseIf InStr(1, "," & ATTRS_SYSTEM_DEFAULTS & ",", "," & nm & ",", vbTextCompare) > 0 Then
                    pobierac = "YES"
                Else
                    pobierac = "NO"
                End If
                wsA.Cells(lastA, ATTRS_ENABLE_COL).Value = pobierac
                wsA.Cells(lastA, ATTRS_TYPE_COL).Value = typ
                wsA.Cells(lastA, ATTRS_EXAMPLE_COL).Value = "'" & CStr(wsDiag.Cells(r, 2).Value)
                known(nm) = lastA
                n = n + 1
            End If
        End If
    Next r
    MergeAttributes = n
End Function

Public Function ParseFolderLink(ByVal link As String, ByRef ds As String, ByRef fp As String) As Boolean
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

' Dopisuje do arkusza Attributes atrybuty znalezione w podfolderach (plik name|example); nic nie kasuje
Public Sub MergeFoundAttributes(ByVal path As String)
    Dim fso As Object, st As Object, tx As String, ln() As String, i As Long, p As Long, nm As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Sub
    On Error GoTo Done
    Set st = CreateObject("ADODB.Stream"): st.Charset = "utf-8": st.Open: st.LoadFromFile path
    If Not st.EOS Then tx = st.ReadText
    st.Close
    Dim wsA As Worksheet, known As Object, lastA As Long, n As Long
    Set wsA = ThisWorkbook.Sheets(SH_ATTRS)
    Set known = CreateObject("Scripting.Dictionary"): known.CompareMode = 1
    lastA = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    If lastA < ATTRS_FIRST_ROW - 1 Then lastA = ATTRS_FIRST_ROW - 1
    For i = ATTRS_FIRST_ROW To lastA
        If Trim(CStr(wsA.Cells(i, 1).Value)) <> "" Then known(Trim(CStr(wsA.Cells(i, 1).Value))) = 1
    Next i
    ln = Split(Replace(Replace(tx, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(ln)
        p = InStr(ln(i), "|")
        If p > 1 Then
            nm = Trim(Left(ln(i), p - 1))
            If nm <> "" And Not known.exists(nm) Then
                lastA = lastA + 1
                wsA.Cells(lastA, 1).Value = nm
                wsA.Cells(lastA, ATTRS_ENABLE_COL).Value = "YES"
                wsA.Cells(lastA, ATTRS_TYPE_COL).Value = "ATRYBUT"
                wsA.Cells(lastA, ATTRS_EXAMPLE_COL).Value = "'" & Mid(ln(i), p + 1)
                known(nm) = 1: n = n + 1
            End If
        End If
    Next i
    If n > 0 Then Notify n & " new attribute(s) found in subfolders were added to sheet " & SH_ATTRS & " (YES)."
Done:
    On Error Resume Next
    fso.DeleteFile path
End Sub
