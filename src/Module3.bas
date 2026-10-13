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
    Set ws2 = ThisWorkbook.Sheets("Sheet_3")
    On Error GoTo 0
    
    If ws2 Is Nothing Then
        MsgBox "Brak arkusza 'Sheet_3'.", vbExclamation
        Exit Sub
    End If
    
    folderPath = ws2.Cells(2, 2).Value
    If Trim(folderPath) = "" Then
        MsgBox "Brak folderu w Sheet_3 (komórka B2). Wpisz tam ścieżkę do folderu PW, aby pobrać jego atrybuty.", vbExclamation
        Exit Sub
    End If
    
    basePath = ThisWorkbook.Path
    psScriptPath = basePath & "\PW_Diag.ps1"
    csvOutputPath = basePath & "\PW_Diag.csv"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim psFile As Object
    Set psFile = fso.CreateTextFile(psScriptPath, True, True)
    
    ' Budowa skryptu diagnostycznego PowerShell
    psFile.WriteLine "$ErrorActionPreference = 'Stop'"
    psFile.WriteLine "Import-Module PWPS_DAB -Force | Out-Null"
    
    ' --- ZMIANA 1: Dynamiczne wycięcie dsName ze ścieżki ---
    psFile.WriteLine "$rawFolder = '" & folderPath & "'.Trim().Trim('\')"
    psFile.WriteLine "$dsName = $rawFolder.Substring(0, $rawFolder.IndexOf('\'))"
    psFile.WriteLine "$folderPath = $rawFolder.Substring($rawFolder.IndexOf('\') + 1)"
    
    psFile.WriteLine "Write-Host 'Logowanie do PW...' -ForegroundColor Cyan"
    psFile.WriteLine "try { New-PWLogin -DatasourceName $dsName -BentleyIMS -NonAdminLogin | Out-Null } catch { exit 1 }"
    
    psFile.WriteLine "$idx = $folderPath.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase)"
    psFile.WriteLine "if ($idx -ge 0) { $folderPath = $folderPath.Substring($idx + 11) } elseif ($folderPath.StartsWith('Documents\', [System.StringComparison]::OrdinalIgnoreCase)) { $folderPath = $folderPath.Substring(10) }"
    psFile.WriteLine "$folderPath = $folderPath.Trim('\')"
    
    ' --- ZMIANA 2: Szukanie konkretnego pliku zamiast folderu ---
    psFile.WriteLine "Write-Host ('Szukam pliku: ' + $folderPath) -ForegroundColor Yellow"
    psFile.WriteLine "$doc = Get-PWDocumentsBySearch -FullPath $folderPath -GetAttributes -WarningAction SilentlyContinue"
    
    ' --- ZMIANA 3: Usunięcie sprawdzania folderu, sprawdzamy od razu plik ---
    psFile.WriteLine "if ($doc) {"
    psFile.WriteLine "    Write-Host 'Pobieranie 1 pliku do analizy atrybutów...' -ForegroundColor Green"
    psFile.WriteLine "    if ($doc) {"
    psFile.WriteLine "        $props = $doc.psobject.properties"
    psFile.WriteLine "        $out = [System.Collections.Generic.List[psobject]]::new()"
    
    ' Pętla zrzucająca surowe właściwości systemowe z jednego pliku
    psFile.WriteLine "        foreach ($p in $props) {"
    psFile.WriteLine "            $val = $p.Value"
    psFile.WriteLine "            if ($val -is [array]) { $val = ($val -join ', ') }"
    psFile.WriteLine "            $out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = $p.Name; 'Wartosc' = $val })"
    psFile.WriteLine "        }"
    
    ' Ręczne dodanie naszych sztucznych atrybutów Updater i Updated do diagnostyki
    psFile.WriteLine "        $uName = @($doc.DocumentUpdaterName, $doc.FileUpdaterName, $doc.UpdaterName, $doc.Updater, $doc.CreatorName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1"
    psFile.WriteLine "        $out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = 'Updater'; 'Wartosc' = if($uName){$uName}else{'Unknown'} })"
    psFile.WriteLine "        $uDate = @($doc.DocumentUpdateDate, $doc.FileUpdateDate, $doc.UpdateDate, $doc.Updated, $doc.CreateDate) | Where-Object { $_ } | Select-Object -First 1"
    psFile.WriteLine "        $out.Add([PSCustomObject]@{ 'Nazwa_Wewnetrzna' = 'Updated'; 'Wartosc' = if($uDate -is [DateTime]){$uDate.ToString('yyyy-MM-dd')}else{$uDate} })"
    
    psFile.WriteLine "        $out.ToArray() | Export-Csv -Path '" & csvOutputPath & "' -NoTypeInformation -Delimiter '|' -Encoding UTF8"
    psFile.WriteLine "        Write-Host 'Zapisano plik diagnostyczny!' -ForegroundColor Green"
    psFile.WriteLine "    } else {"
    psFile.WriteLine "        Write-Host 'Wybrany folder jest pusty! (Brak plików do zbadania atrybutów)' -ForegroundColor Red"
    psFile.WriteLine "    }"
    psFile.WriteLine "} else {"
    psFile.WriteLine "    Write-Host 'Nie znaleziono folderu w ProjectWise!' -ForegroundColor Red"
    psFile.WriteLine "}"
    psFile.WriteLine "Undo-PWLogin | Out-Null"
    psFile.WriteLine "Start-Sleep -Seconds 3"
    psFile.Close
    
    ' Wykonanie skryptu w tle
    Set wsh = CreateObject("WScript.Shell")
    execResult = wsh.Run("powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & """", 1, True)
    
    ' Zrzut pliku wynikowego CSV do nowego arkusza Excel
    If fso.FileExists(csvOutputPath) Then
        On Error Resume Next
        Set wsDiag = ThisWorkbook.Sheets("PW_Diagnostyka")
        If wsDiag Is Nothing Then
            Set wsDiag = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            wsDiag.Name = "PW_Diagnostyka"
        End If
        On Error GoTo 0
        
        wsDiag.Cells.Clear
        wsDiag.Cells(1, 1).Value = "Wewnętrzna nazwa PowerShell (Kopiuj do A w Sheet_3)"
        wsDiag.Cells(1, 2).Value = "Przykładowa wartość (z 1 pliku)"
        wsDiag.Range("A1:B1").Font.Bold = True
        wsDiag.Range("A1:B1").Interior.Color = RGB(220, 230, 241)
        
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
                            wsDiag.Range(wsDiag.Cells(outR, 1), wsDiag.Cells(outR, 2)).Interior.Color = RGB(255, 242, 204)
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
        MsgBox "Diagnostyka zakończona! Sprawdź arkusz 'PW_Diagnostyka'.", vbInformation
    Else
        MsgBox "Nie udało się wygenerować pliku CSV. Sprawdź okno PowerShell.", vbCritical
    End If
    
    ' Sprzątanie po sobie
    On Error Resume Next
    fso.DeleteFile psScriptPath
    fso.DeleteFile csvOutputPath
    On Error GoTo 0
End Sub
