Attribute VB_Name = "Module1"
Option Explicit                                     ' Wymusza jawną deklarację wszystkich zmiennych w kodzie, co zapobiega literówkom

' =======================================================================
' WERSJA: V1.2 (Z obsługą tabeli PW_DATA i pełnymi komentarzami)
' GLOBALNE ZMIENNE DLA MODUŁU LOGOWANIA
' =======================================================================
Private gBusy As Boolean                            ' True, gdy pobieranie z PW trwa w tle (Excel jest wolny)
Private gStart As Date
Private gTxtPath As String, gStructPath As String, gPsPath As String, gCsvPath As String
Private gNumScopes As Long
Private gScopes As Variant
Private gNextPoll As Date
Private gChainStep As Long                          ' 0 = brak lancucha, 1 = szybka aktualizacja, 2 = skan atrybutow, 3 = dokladne pobranie po nowych atrybutach
Private gScanBusy As Boolean, gScanStart As Date, gScanNext As Date, gScanOut As String
Private wsLog As Worksheet                          ' Zmienna przechowująca obiekt arkusza z logami
Private nextLogRow As Long                          ' Zmienna przechowująca numer kolejnego, pustego wiersza do zapisu w logach
Private logStartRow As Long                         ' Zmienna przechowująca numer wiersza początkowego dla danej sesji uruchomienia

' =======================================================================
' GŁÓWNA PROCEDURA (ORKIESTRATOR)
' Zarządza całym procesem: od przygotowania plików, przez uruchomienie
' skryptu PowerShell, aż po analizę różnicową i zapis logów.
' =======================================================================
Public Function DownloadBusy() As Boolean
    DownloadBusy = gBusy Or gScanBusy
End Function

Sub DownloadFromProjectWise_PWPS()                  ' Publiczny punkt wejścia: chroni stan Excela i raportuje błędy
    If gChainStep = 0 And CHAIN_ATTR_SCAN Then gChainStep = 1       ' UPDATE DATA = lancuch: 1 szybka aktualizacja -> 2 skan atrybutow -> 3 dokladne pobranie (gdy sa nowe atrybuty)
    If gBusy Then
        Notify "A ProjectWise download is already running in the background (PowerShell window). Please wait for it to finish."
        Exit Sub
    End If
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False              ' Wyłączenie odświeżania ekranu - ogromne przyspieszenie zapisów komórka po komórce
    Application.EnableEvents = False                ' Zdarzenia arkuszy (Worksheet_Change itp.) nie mogą spowalniać ani przerywać makra
    Application.Calculation = xlCalculationManual   ' Wstrzymanie przeliczania formuł na czas makra
    RunDownload
Done:
    Application.Calculation = prevCalc              ' Zawsze przywróć stan Excela, także po błędzie
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
Fail:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    gBusy = False: gChainStep = 0
    On Error Resume Next
    If Not wsLog Is Nothing Then AddLogEntry "ERROR", errMsg, "-", "-", "-", "-", CLR_DELETED
    Notify "Macro stopped unexpectedly. " & errMsg
    Resume Done
End Sub

Private Sub RunDownload()                           ' Właściwa logika orkiestratora (dawniej treść DownloadFromProjectWise_PWPS)
    ' --- [NOWE] START STOPERA ---
    Dim startExecTime As Date                       ' Deklaracja zmiennej do przechowywania czasu rozpoczęcia
    startExecTime = Now                             ' Zapisanie aktualnej daty i godziny do zmiennej startExecTime
    
    If ThisWorkbook.Path = "" Then                  ' Sprawdzenie, czy plik Excela został już zapisany na dysku (ścieżka nie jest pusta)
        Notify "Please save the Excel workbook to your drive first!" ' Wyświetlenie błędu, jeśli plik nie jest zapisany
        Exit Sub                                    ' Przerwanie działania makra, jeśli plik nie jest zapisany
    End If                                          ' Zakończenie bloku warunkowego sprawdzania ścieżki
    
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet ' Deklaracja trzech zmiennych przechowujących obiekty arkuszy
    Set ws1 = ThisWorkbook.Sheets(SH_DATA)        ' Przypisanie arkusza "Sheet_1" do zmiennej ws1
    Set ws2 = ThisWorkbook.Sheets(SH_FOLDERS)        ' Przypisanie arkusza "Sheet_2" do zmiennej ws2
    Set ws3 = ThisWorkbook.Sheets(SH_ATTRS)        ' Przypisanie arkusza "Sheet_3" do zmiennej ws3
    
    ' --- [NOWE] KONWERSJA TABELI PW_DATA NA ZAKRES (NA POCZĄTKU) ---
    Dim lo As ListObject                            ' Deklaracja zmiennej dla obiektu tabeli Excel
    On Error Resume Next                            ' Ignorowanie błędu, jeśli tabela nie istnieje w arkuszu 1
    Set lo = ws1.ListObjects("PW_DATA")             ' Próba przypisania istniejącej tabeli PW_DATA do zmiennej
    On Error GoTo 0                                 ' Przywrócenie normalnej obsługi błędów systemowych
    If Not lo Is Nothing Then                       ' Jeśli tabela PW_DATA została poprawnie znaleziona
        lo.Unlist                                   ' Skonwertuj tabelę z powrotem na zwykły, płaski zakres komórek (aby makro mogło swobodnie wstawiać/usuwać wiersze)
    End If                                          ' Zakończenie bloku warunkowego dla tabeli
    ' --------------------------------------------------------------
    
    Dim basePath As String                          ' Deklaracja zmiennej przechowującej ścieżkę głównego folderu
    basePath = ThisWorkbook.Path                    ' Przypisanie ścieżki bieżącego pliku Excela do zmiennej
    
    Dim psArrayString As String                     ' Deklaracja zmiennej przechowującej listę atrybutów jako tekst dla PowerShella
    If Application.WorksheetFunction.CountA(ws3.Range(ws3.Cells(ATTRS_FIRST_ROW, 1), ws3.Cells(ws3.Rows.Count, 1))) = 0 Then
        Dim fr As Long, flnk As String              ' Brak atrybutow - wykryj je z pierwszego aktywnego folderu
        For fr = FOLDERS_FIRST_ROW To ws2.Cells(ws2.Rows.Count, FOLDERS_LINK_COL).End(xlUp).Row
            flnk = Trim(CStr(ws2.Cells(fr, FOLDERS_LINK_COL).Value))
            If UCase(Left(flnk, 3)) = "PW:" And Not IsNo(ws2.Cells(fr, FOLDERS_ACTIVE_COL).Value) Then
                DiscoverAttributes flnk, True
                Exit For
            End If
        Next fr
    End If
    psArrayString = GetAttributesFromSheet3()       ' Wywołanie funkcji pobierającej atrybuty i przypisanie wyniku do zmiennej
    If psArrayString = "" Then Exit Sub             ' Jeśli funkcja nie zwróciła żadnych atrybutów, przerwij działanie makra
    
    InitLogSheet                                    ' Wywołanie procedury inicjalizującej arkusz z logami
    
    Dim txtFoldersPath As String                    ' Deklaracja zmiennej na ścieżkę pliku z folderami wejściowymi
    Dim structureLogPath As String                  ' Deklaracja zmiennej na ścieżkę pliku logu struktury
    txtFoldersPath = basePath & "\PW_Folders.txt"   ' Zbudowanie pełnej ścieżki do pliku tekstowego z listą folderów
    structureLogPath = basePath & "\PW_Structure_Log.txt" ' Zbudowanie pełnej ścieżki do pliku tekstowego ze strukturą plików
    
    Dim numScopes As Long                           ' Deklaracja zmiennej przechowującej liczbę przeszukiwanych folderów
    Dim arrCleanScopes As Variant                   ' Deklaracja tablicy na wyczyszczone ścieżki folderów
    If Not PrepareScopesAndTxtFile(ws2, txtFoldersPath, numScopes, arrCleanScopes) Then Exit Sub ' Wygenerowanie TXT z folderami; jeśli fałsz, wyjdź
    
    AddLogEntry "INFO", "Started scanning ProjectWise via PowerShell...", "-", "-", "-", "-", 0 ' Zapisanie informacji o starcie skanowania w logach
    
    Dim psScriptPath As String                      ' Deklaracja zmiennej dla ścieżki do pliku skryptu PowerShell
    psScriptPath = basePath & "\PW_GetFiles.ps1"    ' Zbudowanie pełnej ścieżki dla tworzonego pliku PS1
    GeneratePowerShellScript psScriptPath, psArrayString, GetCustomAttributesPS() ' Wywołanie procedury generującej fizyczny kod skryptu PS1 na dysku
    
    Dim csvOutputPath As String                     ' Deklaracja zmiennej na ścieżkę pliku wyjściowego CSV
    csvOutputPath = basePath & "\PW_Output.csv"     ' Zbudowanie pełnej ścieżki pliku CSV
    DeleteIfExists csvOutputPath                    ' Stare wyniki NIE mogą zostać użyte, gdy PowerShell się nie powiedzie (inaczej delta liczyłaby się na starych danych)
    DeleteIfExists structureLogPath
    DeleteIfExists csvOutputPath & ".done"
    DeleteIfExists csvOutputPath & ".attrs"
    gStart = startExecTime: gTxtPath = txtFoldersPath: gStructPath = structureLogPath
    gPsPath = psScriptPath: gCsvPath = csvOutputPath: gNumScopes = numScopes: gScopes = arrCleanScopes
    WriteAttrMap txtFoldersPath & ".map"             ' znane podfoldery (juz sprawdzone) - PowerShell ich nie skanuje ponownie
    RunPowerShellScript psScriptPath, txtFoldersPath, csvOutputPath, structureLogPath ' PowerShell startuje w tle - Excel NIE czeka
    gBusy = True
    SchedulePoll
    Application.StatusBar = "Pobieranie z ProjectWise w tle... Excel jest wolny (okno PowerShell jest na pasku zadan)."
End Sub

' =======================================================================
' FAZA 2 (po zakonczeniu PowerShella): analiza roznic, tabela, widok
' =======================================================================
Public Sub FinishDownload()
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    RunFinish
Done:
    gBusy = False
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
Fail:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    gChainStep = 0
    On Error Resume Next
    If Not wsLog Is Nothing Then AddLogEntry "ERROR", errMsg, "-", "-", "-", "-", CLR_DELETED
    Notify "Macro stopped unexpectedly. " & errMsg
    Resume Done
End Sub

Private Sub RunFinish()
    Dim ws1 As Worksheet, lo As ListObject
    Set ws1 = ThisWorkbook.Sheets(SH_DATA)
    Application.StatusBar = "Analiza roznic (delta) w Excelu..."
    
    AddLogEntry "INFO", "Started Delta analysis in Excel...", "-", "-", "-", "-", 0 ' Zapisanie w logach informacji o rozpoczęciu analizy Delta
    
    MergeFoundAttributes gCsvPath & ".attrs"        ' atrybuty znalezione w podfolderach -> arkusz Attributes
    ProcessDeltaAnalysis ws1, gCsvPath, gStructPath, gNumScopes, gScopes ' Wywołanie głównego silnika analizy różnic (CSV vs Excel)
    
    CleanUpFiles gTxtPath, gPsPath, gCsvPath ' Wywołanie procedury z pytaniem o usunięcie plików tymczasowych
    
    ' --- ZAPIS CZASU I DATY W KOMÓRCE A8 ---
    Dim endExecTime As Date                         ' Deklaracja zmiennej do przechowywania czasu zakończenia
    Dim durationStr As String                       ' Deklaracja zmiennej do przechowywania czasu trwania w formacie tekstowym
    endExecTime = Now                               ' Zapisanie aktualnej daty i godziny do zmiennej końcowej
    durationStr = Format(endExecTime - gStart, "hh:mm:ss") ' Obliczenie i sformatowanie różnicy czasu do zmiennej
    
    ws1.Cells(8, 1).Value = "Last Database Update: " & Format(Now, "yyyy-mm-dd HH:mm:ss") & " | Total Duration: " & durationStr ' Wpisanie podsumowania w A8
    ws1.Cells(8, 1).Font.Bold = True                ' Pogrubienie tekstu w komórce A8
    ws1.Cells(8, 1).Font.Color = RGB(0, 102, 204)   ' Ustawienie niebieskiego koloru czcionki w A8
    
    
    On Error Resume Next
    RenderDashboard True                            ' Odswiez Dashboard (strona 1 wynikow)
    On Error GoTo 0
    AddLogEntry "INFO", "OPERATION COMPLETED", "-", "-", "-", "-", 0 ' Zapisanie w logach informacji o pomyślnym zakończeniu makra
    wsLog.Columns("A:G").AutoFit                    ' Automatyczne dopasowanie szerokości kolumn w arkuszu logów
    On Error Resume Next
    ThisWorkbook.Sheets(SH_VIEW).Activate           ' Glowny arkusz pracy to Widok
    RefreshNavigator                                ' drzewo nawigatora od nowa z zaktualizowanej bazy
    On Error GoTo 0
    Select Case gChainStep
        Case 1
            gChainStep = 2
            Notify "Update completed. Next: exact attribute scan in the background..."
            Application.OnTime Now + TimeSerial(0, 0, 2), "'" & ThisWorkbook.Name & "'!StartAttrScan"
        Case 3
            gChainStep = 0
            Notify "Update completed (with the newly found attributes). Data merged."
        Case Else
            Notify "Update completed. Data merged."
    End Select
End Sub
                                             ' Koniec głównej procedury

' =======================================================================
' POBIERANIE W TLE: odpytywanie, czy PowerShell skonczyl (plik .done)
' =======================================================================
Private Sub SchedulePoll()
    gNextPoll = Now + TimeSerial(0, 0, 2)
    Application.OnTime gNextPoll, "'" & ThisWorkbook.Name & "'!PollDownload"
End Sub

Public Sub CancelDownloadPoll()                     ' wywolywane przy zamykaniu skoroszytu
    On Error Resume Next
    If gBusy Then Application.OnTime gNextPoll, "'" & ThisWorkbook.Name & "'!PollDownload", , False
    gBusy = False
End Sub

Public Sub PollDownload()
    If Not gBusy Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(gCsvPath & ".done") Then
        FinishDownload
        Exit Sub
    End If
    Application.StatusBar = "Pobieranie z ProjectWise w tle: " & Format(Now - gStart, "nn:ss") & " - Excel jest wolny, mozesz pracowac."
    If Now - gStart > TimeSerial(0, 0, 15) And Not PsRunning() Then   ' okno PowerShell zamkniete lub skrypt padl bez ukonczenia
        gBusy = False: gChainStep = 0
        Application.StatusBar = False
        Notify "Download aborted (PowerShell window closed or login/script error). Sheet data was not changed."
        Exit Sub
    End If
    SchedulePoll
End Sub

Private Function PsRunning(Optional ByVal scriptName As String = "PW_GetFiles.ps1") As Boolean
    On Error Resume Next
    Dim n As Long
    n = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name='powershell.exe' AND CommandLine LIKE '%" & scriptName & "%'").Count
    If Err.Number <> 0 Then PsRunning = True Else PsRunning = (n > 0)    ' bez WMI zakladamy, ze dziala
End Function

Sub ShowNavigator()                                 ' Procedura wywołująca formularz nawigacyjny (UserForm)
    On Error Resume Next                            ' Ignorowanie ewentualnych błędów (np. gdy formularz nie istnieje)
    Unload frmNavigator                             ' Zamknięcie formularza w tle, jeśli był już otwarty
    frmNavigator.Show vbModeless                    ' Otwarcie formularza w trybie niemodalnym (pozwala na klikanie w Excela)
    On Error GoTo 0                                 ' Przywrócenie standardowej obsługi błędów
End Sub                                             ' Koniec procedury

' =======================================================================
' 1. CZYTANIE ATRYBUTÓW PROSTO Z SHEET_3
' =======================================================================
Private Function GetAttributesFromSheet3() As String ' Deklaracja funkcji zwracającej tekst (string z atrybutami)
    Dim ws3 As Worksheet                            ' Deklaracja zmiennej dla arkusza 3
    Dim lastRow As Long, i As Long                  ' Deklaracja zmiennych przechowujących numery wierszy i iteratora pętli
    Dim attrName As String                          ' Deklaracja zmiennej na nazwę pojedynczego atrybutu
    Dim psArray As String                           ' Deklaracja zmiennej na ostateczny string w formacie tablicy PS
    
    On Error Resume Next                            ' Ignorowanie błędów przy przypisywaniu arkusza
    Set ws3 = ThisWorkbook.Sheets(SH_ATTRS)        ' Próba przypisania arkusza 3 do zmiennej ws3
    On Error GoTo 0                                 ' Przywrócenie standardowej obsługi błędów
    
    If ws3 Is Nothing Then                          ' Sprawdzenie, czy arkusz 3 istnieje w skoroszycie
        Notify "Sheet '" & SH_ATTRS & "' not found!" ' Wyświetlenie błędu o braku arkusza
        GetAttributesFromSheet3 = ""                ' Zwrócenie pustego ciągu znaków jako wynik funkcji
        Exit Function                               ' Przerwanie działania funkcji
    End If                                          ' Koniec warunku
    
    lastRow = ws3.Cells(ws3.rows.count, "A").End(xlUp).Row ' Znalezienie ostatniego wiersza z danymi w kolumnie A
    If lastRow < ATTRS_FIRST_ROW Then                ' Jeśli brak atrybutów poniżej nagłówka (tylko nagłówek) lub mniej
        Notify "Column A in sheet '" & SH_ATTRS & "' is empty. Add attribute names there!" ' Wyświetl ostrzeżenie o braku atrybutów
        GetAttributesFromSheet3 = ""                ' Zwróć pusty ciąg znaków
        Exit Function                               ' Przerwij funkcję
    End If                                          ' Koniec warunku
    
    psArray = "@("                                  ' Inicjalizacja początku definicji tablicy w składni PowerShell
    Dim count As Long                               ' Deklaracja licznika atrybutów
    count = 0                                       ' Wyzerowanie licznika atrybutów
    
    For i = ATTRS_FIRST_ROW To lastRow               ' Pętla po wierszach z atrybutami
        attrName = Trim(ws3.Cells(i, 1).Value)      ' Pobranie nazwy atrybutu z komórki i usunięcie zbędnych spacji
        If IsNo(ws3.Cells(i, ATTRS_ENABLE_COL).Value) Then attrName = "" ' atrybut wylaczony w arkuszu Atrybuty
        If attrName <> "" Then                      ' Sprawdzenie czy komórka nie była pusta
            psArray = psArray & "'" & attrName & "'," ' Doklejenie atrybutu w apostrofach i przecinka do stringu
            count = count + 1                       ' Zwiększenie licznika dodanych atrybutów
        End If                                      ' Koniec warunku
    Next i                                          ' Koniec pętli i przejście do kolejnego wiersza
    
    If count > 0 Then                               ' Jeśli dodano przynajmniej jeden atrybut
        psArray = Left(psArray, Len(psArray) - 1) & ")" ' Ucięcie ostatniego przecinka i dodanie zamykającego nawiasu tablicy
        GetAttributesFromSheet3 = psArray           ' Przypisanie gotowego tekstu jako wynik funkcji
    Else                                            ' W przeciwnym razie (brak poprawnych atrybutów)
        Notify "No valid attributes found in sheet '" & SH_ATTRS & "'!" ' Wyświetl błąd
        GetAttributesFromSheet3 = ""                ' Zwróć pusty ciąg
    End If                                          ' Koniec warunku
End Function                                        ' Koniec funkcji

' =======================================================================
' 2. LOGOWANIE W EXCELU (Przygotowanie Arkusza "Log")
' =======================================================================
' Atrybuty projektu (do ColumnsToReturn): typ ATRYBUT w arkuszu Atrybuty albo pusty typ i nazwa spoza listy wlasciwosci systemowych
Private Function GetCustomAttributesPS() As String
    Dim ws3 As Worksheet, lastRow As Long, i As Long, nm As String, typ As String, res As String
    Set ws3 = ThisWorkbook.Sheets(SH_ATTRS)
    lastRow = ws3.Cells(ws3.rows.count, "A").End(xlUp).Row
    For i = ATTRS_FIRST_ROW To lastRow
        nm = Trim(CStr(ws3.Cells(i, 1).Value))
        typ = UCase(Trim(CStr(ws3.Cells(i, ATTRS_TYPE_COL).Value)))
        If nm <> "" And Not IsNo(ws3.Cells(i, ATTRS_ENABLE_COL).Value) Then
            If typ = "ATRYBUT" Or (typ = "" And InStr(1, "," & ATTRS_SYSTEM_DEFAULTS & ",", "," & nm & ",", vbTextCompare) = 0) Then
                res = res & "'" & Replace(nm, "'", "''") & "',"
            End If
        End If
    Next i
    If res = "" Then GetCustomAttributesPS = "@()" Else GetCustomAttributesPS = "@(" & Left(res, Len(res) - 1) & ")"
End Function

Private Sub InitLogSheet()                          ' Deklaracja procedury przygotowującej arkusz logów
    On Error Resume Next                            ' Ignorowanie błędów przy próbie przypisania arkusza
    Set wsLog = ThisWorkbook.Sheets(SH_LOG)          ' Przypisanie arkusza "Log" do zmiennej wsLog
    On Error GoTo 0                                 ' Przywrócenie standardowej obsługi błędów
    
    If wsLog Is Nothing Then                        ' Jeśli arkusz "Log" nie istnieje w skoroszycie
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)) ' Dodanie nowego arkusza na samym końcu
        wsLog.Name = SH_LOG                          ' Zmiana nazwy nowo dodanego arkusza na "Log"
        wsLog.Cells(1, 1).Value = "Date and Time"   ' Utworzenie nagłówka dla kolumny z datą
        wsLog.Cells(1, 2).Value = "Action"          ' Utworzenie nagłówka dla kolumny akcji
        wsLog.Cells(1, 3).Value = "File Name"       ' Utworzenie nagłówka dla nazwy pliku
        wsLog.Cells(1, 4).Value = "Changed Attribute" ' Utworzenie nagłówka dla zmienionego atrybutu
        wsLog.Cells(1, 5).Value = "OLD Value"       ' Utworzenie nagłówka dla starej wartości
        wsLog.Cells(1, 6).Value = "NEW Value"       ' Utworzenie nagłówka dla nowej wartości
        wsLog.Cells(1, 7).Value = "User"            ' Utworzenie nagłówka dla użytkownika dokonującego zmiany
        StyleHeader wsLog.Range("A1:G1")                ' Jednolity styl nagłówka
        wsLog.Outline.SummaryRow = xlAbove          ' Ustawienie grupowania wierszy tak, aby podsumowanie było nad grupą
    End If                                          ' Koniec warunku tworzenia arkusza
    
    If Trim(CStr(wsLog.Cells(1, 1).Value)) = "" Then   ' Arkusz Log istnieje, ale nie ma nagłówków (np. pusty szablon)
        wsLog.Range("A1:G1").Value = Array("Date and Time", "Action", "File Name", "Changed Attribute", "OLD Value", "NEW Value", "User")
        StyleHeader wsLog.Range("A1:G1")
        wsLog.Range("A:A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
        wsLog.Activate
        ActiveWindow.FreezePanes = False
        wsLog.Range("A2").Select
        ActiveWindow.FreezePanes = True
    End If
    Dim lastLogCell As Range                        ' Deklaracja zmiennej dla komórki końcowej
    Set lastLogCell = wsLog.Columns(1).Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious) ' Szukanie od dołu pierwszej niepustej komórki w A
    
    If Not lastLogCell Is Nothing Then              ' Jeśli znaleziono jakąś zawartość w kolumnie A
        nextLogRow = lastLogCell.Row + 1            ' Ustaw wiersz logowania na pierwszy pusty (poniżej znalezionego)
    Else                                            ' Jeśli kolumna jest pusta
        nextLogRow = 2                              ' Zacznij wpisywanie logów od 2. wiersza
    End If                                          ' Koniec warunku określającego wiersz
    
    logStartRow = nextLogRow                        ' Zapisanie, od którego wiersza zaczyna się obecna sesja (do grupowania)
    wsLog.Cells(logStartRow, 1).Value = Now         ' Wpisanie bieżącej daty do pierwszej komórki nowej sesji
    wsLog.Cells(logStartRow, 2).Value = "START"     ' Wpisanie akcji "START"
    wsLog.Cells(logStartRow, 3).Value = "--- EXPORT BATCH ---" ' Wpisanie informacji o rozpoczęciu partii zadań
    wsLog.Range(wsLog.Cells(logStartRow, 1), wsLog.Cells(logStartRow, 7)).Interior.Color = RGB(153, 204, 255) ' Ustawienie niebieskiego tła dla wiersza startowego
    wsLog.Range(wsLog.Cells(logStartRow, 1), wsLog.Cells(logStartRow, 7)).Font.Bold = True ' Pogrubienie wiersza startowego
    nextLogRow = nextLogRow + 1                     ' Inkrementacja numeru wiersza dla następnych logów
End Sub                                             ' Koniec procedury

Private Sub AddLogEntry(action As String, fName As String, attr As String, oldV As String, newV As String, user As String, bgColor As Long) ' Procedura wprowadzająca konkretny log
    wsLog.Cells(nextLogRow, 1).Value = Now          ' Wpisanie bieżącej daty operacji do logu
    wsLog.Cells(nextLogRow, 2).Value = action       ' Wpisanie nazwy akcji (ADDED, DELETED, CHANGED)
    wsLog.Cells(nextLogRow, 3).Value = fName        ' Wpisanie nazwy edytowanego pliku
    wsLog.Cells(nextLogRow, 4).Value = attr         ' Wpisanie nazwy zmodyfikowanego atrybutu
    
    If oldV <> "-" And oldV <> "Existed" And oldV <> "" Then wsLog.Cells(nextLogRow, 5).Value = "'" & oldV Else wsLog.Cells(nextLogRow, 5).Value = oldV ' Formatowanie starej wartości na tekst, by uniknąć rozpoznania jako formuła Excela
    If newV <> "-" And Left(newV, 7) <> "Deleted" And newV <> "" Then wsLog.Cells(nextLogRow, 6).Value = "'" & newV Else wsLog.Cells(nextLogRow, 6).Value = newV ' Formatowanie nowej wartości na tekst
    
    wsLog.Cells(nextLogRow, 7).Value = user         ' Wpisanie imienia/nazwiska osoby dokonującej aktualizacji
    If bgColor <> 0 Then wsLog.Range(wsLog.Cells(nextLogRow, 1), wsLog.Cells(nextLogRow, 7)).Interior.Color = bgColor ' Nałożenie zdefiniowanego koloru tła na wpis logu
    nextLogRow = nextLogRow + 1                     ' Inkrementacja (przejście do następnego pustego wiersza)
End Sub                                             ' Koniec procedury

' =======================================================================
' 3. ZAKRESY SKANOWANIA
' =======================================================================
Private Function PrepareScopesAndTxtFile(ws2 As Worksheet, txtPath As String, ByRef numScopes As Long, ByRef arrCleanScopes As Variant) As Boolean ' Funkcja parsująca linki z Arkusza 2 i tworząca plik konfiguracyjny TXT
    Dim lastRow As Long, i As Long                  ' Deklaracja zmiennych przechowujących numer wiersza i iterator
    Dim rawLink As String, dsName As String, folderPath As String ' Zmienne na surowy link, nazwę źródła (Datasource) i ścieżkę folderu
    Dim fso As Object, txtFile As Object            ' Zmienne do obsługi systemu plików (File System Object)
    
    Dim lastRowCell As Range                        ' Deklaracja zmiennej zakresu dla kolumny
    Set lastRowCell = ws2.Columns(FOLDERS_LINK_COL).Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious) ' Szukanie ostatniego wiersza w kolumnie B
    If Not lastRowCell Is Nothing Then lastRow = lastRowCell.Row Else lastRow = 1 ' Jeśli znaleziono, ustaw lastRow, w przeciwnym razie ustaw 1
    
    If lastRow < FOLDERS_FIRST_ROW Then                ' Jeśli nie ma linków do przetworzenia
        Notify "No active PW links in sheet '" & SH_FOLDERS & "' (column B)." ' Ostrzeżenie dla użytkownika
        PrepareScopesAndTxtFile = False             ' Zwróć fałsz (przerwij dalsze makro)
        Exit Function                               ' Przerwij funkcję
    End If                                          ' Koniec warunku
    
    Set fso = CreateObject("Scripting.FileSystemObject") ' Zainicjowanie biblioteki FSO
    Set txtFile = fso.CreateTextFile(txtPath, True, False) ' Utworzenie nowego, pustego pliku TXT (True nadpisuje istniejący)
    
    Dim tempArr() As String                         ' Deklaracja wewnętrznej, tymczasowej tablicy tekstowej
    ReDim tempArr(1 To lastRow)                     ' Zmiana wielkości tablicy na podstawie maksymalnej liczby wierszy
    numScopes = 0                                   ' Wyzerowanie licznika wczytanych folderów
    
    For i = FOLDERS_FIRST_ROW To lastRow            ' Pętla po wierszach arkusza Foldery
        rawLink = Trim(ws2.Cells(i, FOLDERS_LINK_COL).Value)
        If IsNo(ws2.Cells(i, FOLDERS_ACTIVE_COL).Value) Then rawLink = "" ' Wiersz wyłączony przez użytkownika (Aktywny = NIE)       ' Pobranie i oczyszczenie linku ze spacji wiodących
        If rawLink <> "" Then                       ' Przetwarzaj tylko, jeśli link nie jest pusty
            Dim normalizedLink As String            ' Zmienna pomocnicza na link ze spójnymi ukośnikami
            normalizedLink = Replace(rawLink, "/", "\") ' Zamiana ukośników w prawo na ukośniki w lewo (format Windows)
            
            If Left(UCase(normalizedLink), 3) = "PW:" Then ' Sprawdzenie, czy link zaczyna się od prefixu "PW:"
                Dim withoutPrefix As String         ' Zmienna na link obcięty o "PW:"
                withoutPrefix = Mid(normalizedLink, 4) ' Zapisanie ciągu znaków od 4 litery w górę
                
                Do While Left(withoutPrefix, 1) = "\" ' Usuwanie nadmiarowych ukośników początkowych w pętli
                    withoutPrefix = Mid(withoutPrefix, 2) ' Obcięcie pierwszego znaku, jeśli jest ukośnikiem
                Loop                                ' Powtarzaj aż ciąg nie będzie zaczynał się od ukośnika
                
                Dim docPos As Long                  ' Pozycja frazy "\Documents\" w linku
                docPos = InStr(1, withoutPrefix, "\Documents\", vbTextCompare) ' Szukanie, gdzie w ścieżce występuje segment dokumentów
                
                If docPos > 0 Then                  ' Jeśli fraza została znaleziona
                    dsName = Left(withoutPrefix, docPos - 1) ' Wyodrębnienie nazwy Datasource jako lewa część przed "\Documents\"
                    folderPath = Mid(withoutPrefix, docPos + 11) ' Wyodrębnienie samej ścieżki folderu (wszystko po)
                Else                                ' Jeśli fraza nie występuje
                    Dim firstSlash As Long          ' Pozycja pierwszego wystąpienia ukośnika
                    firstSlash = InStr(1, withoutPrefix, "\") ' Szukanie pierwszego separatora w ciągu
                    If firstSlash > 0 Then          ' Jeśli występuje
                        dsName = Left(withoutPrefix, firstSlash - 1) ' Nazwa Datasource to tekst przed tym ukośnikiem
                        folderPath = Mid(withoutPrefix, firstSlash + 1) ' Ścieżka folderu to tekst za ukośnikiem
                    Else                            ' Jeśli nie ma w ogóle ukośników (to tylko samo źródło)
                        dsName = withoutPrefix      ' Przypisanie całego wyrażenia do nazwy źródła
                        folderPath = ""             ' Ścieżka folderu jest pusta
                    End If                          ' Koniec sprawdzania ukośników
                End If                              ' Koniec sprawdzania "\Documents\"
                
                Dim cleanRoot As String             ' Zmienna na ostatecznie wyczyszczoną z sufiksów ścieżkę
                cleanRoot = folderPath              ' Przypisanie dotychczas przetworzonej ścieżki
                cleanRoot = Replace(cleanRoot, " (Working folder)", "", 1, -1, vbTextCompare) ' Usunięcie śmieciowego tekstu z niektórych skopiowanych linków PW
                If Right(cleanRoot, 1) = "\" Then cleanRoot = Left(cleanRoot, Len(cleanRoot) - 1) ' Odcięcie ukośnika, jeśli jest na samym końcu
                If Left(cleanRoot, 1) = "\" Then cleanRoot = Mid(cleanRoot, 2) ' Odcięcie ukośnika, jeśli przypadkowo został na przodzie
                
                If dsName <> "" And cleanRoot <> "" Then ' Zapisz do TXT tylko, jeśli wyciągnęto poprawnie i źródło, i ścieżkę
                    txtFile.WriteLine dsName & "|" & cleanRoot ' Zapis do pliku linii o formacie Źródło|Ścieżka_Folderu
                    numScopes = numScopes + 1       ' Zwiększenie licznika przetworzonych lokalizacji
                    tempArr(numScopes) = UCase(Trim(cleanRoot)) ' Zapisanie lokalizacji do pamięci wewnętrznej na potrzebe analizy usunięć (wielkimi literami)
                End If                              ' Koniec warunku weryfikującego kompletność danych
            End If                                  ' Koniec sprawdzania prefiksu PW:
        End If                                      ' Koniec weryfikacji pustego pola
    Next i                                          ' Kontynuacja pętli po wierszach arkusza
    
    txtFile.Close                                   ' Zamknięcie fizycznego pliku TXT zapisywanego na dysku
    arrCleanScopes = tempArr                        ' Przekazanie wypełnionej tablicy z powrotem przez referencję argumentu
    PrepareScopesAndTxtFile = True                  ' Zwrócenie Prawdy - wszystko wygenerowało się pomyślnie
End Function                                        ' Koniec funkcji budującej plik wejściowy

' =======================================================================
' 4. POWERSHELL + GENEROWANIE LOGU STRUKTURY
' =======================================================================
Private Sub GeneratePowerShellScript(psPath As String, psArrayString As String, customArrayString As String) ' Generator fizycznego kodu skryptu PS1
    Dim fso As Object, psFile As Object             ' Zmienne systemowe na obiekt pliku i generator tekstu
    Set fso = CreateObject("Scripting.FileSystemObject") ' Uruchomienie narzędzia systemowego
    Set psFile = fso.CreateTextFile(psPath, True, True) ' Utworzenie/Nadpisanie pliku PS1, w trybie zapisu (Unicode/UTF8)
    
    psFile.WriteLine "param([string]$InputFile, [string]$OutputFile, [string]$LogFile)" ' Zapis kodu PS: Parametry przyjmujące ścieżki plików z makra
    psFile.WriteLine "$ErrorActionPreference = 'Stop'" ' Zapis kodu PS: Konfiguracja, zatrzymaj skrypt jeśli wystąpi krytyczny błąd PowerShell
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz dla czytelności kodu w PS1
    psFile.WriteLine "$globalWatch = [System.Diagnostics.Stopwatch]::StartNew()" ' Zapis kodu PS: Uruchomienie głównego wbudowanego stopera PS
    WritePsPrelude psFile, "update"
    psFile.WriteLine "Say-Head 'Loading ProjectWise module'" ' Zapis kodu PS: Wypisanie informacji w konsoli
    psFile.WriteLine "Import-Module PWPS_DAB -Force -DisableNameChecking 3>$null 4>$null" ' Zapis kodu PS: Wymuszone wczytanie biblioteki integracji z ProjectWise
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz
    psFile.WriteLine "$inputData = @(Get-Content $InputFile | Where-Object { $_.Trim() -ne '' })" ' Zapis kodu PS: Wczytanie zawartości pliku TXT (wiersze konfiguracji)
    psFile.WriteLine "$totalFolders = $inputData.Count" ' Zapis kodu PS: Obliczenie łącznej liczby zdefiniowanych folderów
    psFile.WriteLine "$folderIndex = 0"             ' Zapis kodu PS: Reset wskaźnika przetwarzanego folderu
    psFile.WriteLine "$allVersions = " & IIf(GET_ALL_VERSIONS, "$true", "$false") ' Czy pobierac wszystkie rewizje plikow (modConfig)
    psFile.WriteLine "$loggedDs = $null"
    psFile.WriteLine "$okScopes = New-Object System.Collections.ArrayList" ' Lista folderów, które faktycznie udało się przeskanować (używana później do bezpiecznego oznaczania usunięć)
    psFile.WriteLine "$selectedAttrs = " & psArrayString ' Zapis kodu PS: Wstrzyknięcie z Excela tablicy z atrybutami (np. @('Attr1','Attr2'))
    psFile.WriteLine "function Get-Attr($d, $n) { $p = $d.psobject.Properties[$n]; if ($p) { return $p.Value }; foreach ($q in $d.psobject.Properties) { $v = $q.Value; if ($null -eq $v -or $v -is [string] -or $v -is [ValueType] -or $v -is [array]) { continue }; if ($v -is [System.Collections.IDictionary]) { if ($v.Contains($n)) { return $v[$n] } } else { $pp = $v.psobject.Properties[$n]; if ($pp) { return $pp.Value } } }; return $null }" ' Pomocnik: atrybut moze byc wlasciwoscia dokumentu albo siedziec w zagniezdzonym obiekcie/slowniku
    psFile.WriteLine "$customAttrs = " & customArrayString ' Atrybuty projektu przekazywane do ColumnsToReturn (szybki tryb)
    psFile.WriteLine "$discoverSub = " & IIf(DISCOVER_SUBFOLDER_ATTRS, "$true", "$false")
    psFile.WriteLine "$foundAttrs = [ordered]@{}; $usedFast = $false"
    psFile.WriteLine "$fastFetch = " & IIf(FAST_FETCH, "$true", "$false") ' Szybkie pobieranie (DataTable) z automatycznym wyjsciem awaryjnym na stary tryb
    psFile.WriteLine "$umap = $null; $smap = $null"
    psFile.WriteLine "function Get-DtVal($row, $attr) { $c = $row.Table.Columns; $v = $null; if ($c.Contains($attr)) { $v = $row[$attr] } elseif ($c.Contains('Document' + $attr)) { $v = $row['Document' + $attr] } elseif ($attr -eq 'WorkflowState') { $sid = [string]$row['DocumentStateId']; if ($smap.ContainsKey($sid)) { $v = $smap[$sid] } else { $v = $sid } } elseif ($attr -eq 'CreateDate' -and $c.Contains('DocumentCreatedDate')) { $v = $row['DocumentCreatedDate'] } elseif ($attr -eq 'DocumentCreatorName' -and $c.Contains('DocumentCreatorId')) { $cid = [string]$row['DocumentCreatorId']; if ($umap.ContainsKey($cid)) { $v = $umap[$cid] } else { $v = 'ID ' + $cid } }; if ($v -is [DBNull]) { $v = $null }; if ($v -is [DateTime]) { $v = $v.ToString('yyyy-MM-dd') }; if ($v -is [string]) { $v = $v.Replace('|', '-').Replace([string][char]10, ' ').Replace([string][char]13, '').Trim() }; return $v }"
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz
    psFile.WriteLine "$results = @(foreach ($line in $inputData) {" ' Zapis kodu PS: Inicjalizacja tablicy na wyniki i początek pętli
    psFile.WriteLine " $parts = $line.Split('|')"   ' Zapis kodu PS: Rozbicie wiersza za pomocą separatora "|"
    psFile.WriteLine " if ($parts.Count -lt 2) { continue }" ' Zapis kodu PS: Jeśli nie da się rozbić na źródło i folder, pomiń
    psFile.WriteLine " $dsName = $parts[0].Trim()"  ' Zapis kodu PS: Zmienna PS z wyłuskaną nazwą bazy danych
    psFile.WriteLine " $rawFolder = $parts[1].Trim()" ' Zapis kodu PS: Zmienna PS ze ścieżką folderu
    psFile.WriteLine " $folderIndex++"              ' Zapis kodu PS: Zwiększenie wskaźnika aktualnie przetwarzanego wiersza konfiguracyjnego
    psFile.WriteLine " $folderWatch = [System.Diagnostics.Stopwatch]::StartNew()" ' Zapis kodu PS: Start stopera dla pojedynczego zadania (folderu)
    psFile.WriteLine " Write-Host ('================================') -ForegroundColor DarkCyan" ' Zapis kodu PS: Estetyczny separator w konsoli
    psFile.WriteLine " Write-Host ('Datasource: ' + $dsName) -ForegroundColor Cyan" ' Zapis kodu PS: Wyświetla przetwarzane źródło
    psFile.WriteLine " Write-Host ('Scanning folder AND subfolders: ' + $rawFolder) -ForegroundColor Yellow" ' Zapis kodu PS: Wyświetla nazwę skanowanego folderu
    psFile.WriteLine " try {"                       ' Zapis kodu PS: Rozpoczęcie bloku łapania błędów
    psFile.WriteLine " if ($loggedDs -ne $dsName) { if ($loggedDs) { Undo-PWLogin | Out-Null }; New-PWLogin -DatasourceName $dsName -BentleyIMS -NonAdminLogin | Out-Null; $loggedDs = $dsName }" ' Logowanie tylko gdy zmienia sie datasource (wczesniej: osobne logowanie dla kazdego folderu)
    psFile.WriteLine " } catch {"                   ' Zapis kodu PS: Złapanie błędu logowania
    psFile.WriteLine " Write-Host ('Login ERROR: ' + $_.Exception.Message) -ForegroundColor Red" ' Zapis kodu PS: Wyświetlenie błędu czerwonym kolorem
    psFile.WriteLine " continue"                    ' Zapis kodu PS: Przejście do następnego folderu
    psFile.WriteLine " }"                           ' Zapis kodu PS: Koniec obsługi błędu
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz
    psFile.WriteLine " $currentDs = Get-PWCurrentDatasource" ' Zapis kodu PS: Zweryfikowanie podłączonego środowiska
    psFile.WriteLine " if (-not $currentDs) { Write-Host 'No login!' -ForegroundColor Red; continue }" ' Zapis kodu PS: Zabezpieczenie braku autoryzacji
    psFile.WriteLine " $folderPath = $rawFolder"    ' Zapis kodu PS: Skopiowanie zmiennej ścieżki
    psFile.WriteLine " $idx = $folderPath.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase)" ' Zapis kodu PS: Szukanie indeksu frazy
    psFile.WriteLine " if ($idx -ge 0) { $folderPath = $folderPath.Substring($idx + 11) } elseif ($folderPath.StartsWith('Documents\', [System.StringComparison]::OrdinalIgnoreCase)) { $folderPath = $folderPath.Substring(10) }" ' Zapis kodu PS: Upewnienie się w skrypcie PS że folder omija frazę "Documents"
    psFile.WriteLine " $folderPath = $folderPath.Trim('\')" ' Zapis kodu PS: Usunięcie ewentualnych ukośników bocznych
    psFile.WriteLine " try {"                       ' Zapis kodu PS: Główny blok bezpieczny wyciągania informacji o plikach
    psFile.WriteLine " $pwFoldersList = @(Get-PWFolders -FolderPath $folderPath -ErrorAction SilentlyContinue)" ' Zapis kodu PS: Wyszukiwanie na serwerze obiektu folderu
    psFile.WriteLine " if ($pwFoldersList.Count -eq 0) { $pwFoldersList = @(Get-PWFolders -FolderPath ('\' + $folderPath) -ErrorAction SilentlyContinue) }" ' Zapis kodu PS: Jeśli zły układ ukośników to dodaje na starcie i szuka jeszcze raz
    psFile.WriteLine " if ($pwFoldersList.Count -eq 0) {" ' Zapis kodu PS: Fallback jeśli wciąż nie znaleziono pod bezpośrednim adresem
    psFile.WriteLine " $leafName = $folderPath.Split('\')[-1]" ' Zapis kodu PS: Szuka tylko po samej nazwie ostatniego folderu w drzewie
    psFile.WriteLine " $pwFoldersList = @(Get-PWFolders -FolderName $leafName -ErrorAction SilentlyContinue)" ' Zapis kodu PS: Odpytuje serwer po samej nazwie
    psFile.WriteLine " }"                           ' Zapis kodu PS: Zamknięcie wariantu fallback
    psFile.WriteLine " if ($pwFoldersList.Count -gt 0) {" ' Zapis kodu PS: Jeżeli obiekt folderu został pomyślnie dopasowany
    psFile.WriteLine " [void]$okScopes.Add($rawFolder)" ' Folder znaleziony na serwerze - jego brakujące pliki można uznać za usunięte
    psFile.WriteLine " $pwFolder = $pwFoldersList[0]" ' Zapis kodu PS: Zapisuje id dopasowanego folderu
    psFile.WriteLine " Write-Progress -Id 1 -Activity ""Fetching ALL documents (Recursive) from PW..."" -Status ""Folder $folderIndex of $totalFolders""" ' Zapis kodu PS: Wyświetla pasek pobierania
    
    psFile.WriteLine " $documents = $null"
    psFile.WriteLine " $fetchWatch = [System.Diagnostics.Stopwatch]::StartNew()"
    psFile.WriteLine " if ($fastFetch) {"
    psFile.WriteLine "     try {"
    psFile.WriteLine "         $gv = @{}; if ($allVersions) { $gv['GetVersionsToo'] = $true }"
    psFile.WriteLine "         $cr = @{}; if ($customAttrs.Count -gt 0) { $cr['ColumnsToReturn'] = $customAttrs }"
    psFile.WriteLine "         $r = @(Get-PWDocumentsBySearchExtended -FolderPath $folderPath -PopulatePath @gv @cr -WarningAction SilentlyContinue -ErrorAction Stop)"
    psFile.WriteLine "         if ($r.Count -eq 0) { throw 'brak wynikow' }"
    psFile.WriteLine "         if ($r[0] -is [System.Data.DataTable]) { throw 'zwrocono DataTable zamiast obiektow' }"
    psFile.WriteLine "         if ([string]::IsNullOrWhiteSpace([string]$r[0].FullPath)) { throw 'FullPath pusty mimo PopulatePath' }"
    psFile.WriteLine "         if ([string]::IsNullOrWhiteSpace([string]$r[0].DocumentURN) -and [string]::IsNullOrWhiteSpace([string]$r[0].DocumentGUIDString)) { throw 'brak DocumentURN i GUID' }"
    psFile.WriteLine "         foreach ($d in $r) { if ([string]::IsNullOrWhiteSpace([string]$d.DocumentURN)) { $d | Add-Member -NotePropertyName DocumentURN -NotePropertyValue ('pw://' + $dsName + '/Documents/D{' + ([string]$d.DocumentGUIDString).Trim('{}').ToLower() + '}') -Force } }"
    psFile.WriteLine "         $documents = $r; $usedFast = $true"
    psFile.WriteLine "         Write-Host ('FAST PATH OK: ' + $r.Count + ' documents in ' + $fetchWatch.Elapsed.ToString('mm\:ss')) -ForegroundColor Green"
    psFile.WriteLine "     } catch { Write-Host ('FAST PATH FAILED, using slow mode: ' + $_.Exception.Message) -ForegroundColor Yellow }"
    psFile.WriteLine " }"
    psFile.WriteLine " if (-not $documents) {"
    psFile.WriteLine " $extra = @{}; if ($allVersions -and (Get-Command Get-PWDocumentsBySearch).Parameters.ContainsKey('GetVersionsToo')) { $extra['GetVersionsToo'] = $true }" ' Z GetVersionsToo wyszukiwanie zwraca tez starsze rewizje (A, B, C...)
    psFile.WriteLine " $documents = @(Get-PWDocumentsBySearch -FolderID $pwFolder.ProjectID -GetAttributes @extra -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)"
    psFile.WriteLine " }"
    psFile.WriteLine "if ($discoverSub -and @($documents).Count -gt 0) {"
    psFile.WriteLine "    $dw = [System.Diagnostics.Stopwatch]::StartNew(); $seenF = @{}; $nF = 0; $newF = New-Object System.Collections.ArrayList; $knownF = @{}; if (Test-Path ($InputFile + '.map')) { foreach ($ml in (Get-Content ($InputFile + '.map') -Encoding UTF8)) { $mp = $ml.Split('|'); if ($mp[0]) { $knownF[$mp[0]] = 1 } } }"
    psFile.WriteLine "    Write-Host 'Checking attributes of one file per subfolder...' -ForegroundColor Yellow"
    psFile.WriteLine "    foreach ($d0 in $documents) {"
    psFile.WriteLine "        if ($dw.Elapsed.TotalSeconds -gt 600) { Write-Host 'Subfolder attribute check stopped after 10 min.' -ForegroundColor Yellow; break }"
    psFile.WriteLine "        $fp0 = [string]$d0.FullPath; $ls0 = $fp0.LastIndexOf('\'); if ($ls0 -lt 0) { continue }"
    psFile.WriteLine "        $fk0 = $fp0.Substring(0, $ls0); if ($seenF.ContainsKey($fk0) -or $knownF.ContainsKey($fk0)) { continue }; $seenF[$fk0] = $true; $nF++"
    psFile.WriteLine "        $rel0 = $fk0; $i0 = $rel0.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase); if ($i0 -ge 0) { $rel0 = $rel0.Substring($i0 + 11) } elseif ($rel0.StartsWith('Documents\', [System.StringComparison]::OrdinalIgnoreCase)) { $rel0 = $rel0.Substring(10) }"
    psFile.WriteLine "        if ($nF % 20 -eq 0) { Write-Progress -Id 2 -Activity 'Checking attributes per subfolder' -Status ('Subfolder ' + $nF + ' | ' + $dw.Elapsed.ToString('mm\:ss')) }"
    psFile.WriteLine "        try {"
    psFile.WriteLine "            $fo0 = @(Get-PWFolders -FolderPath $rel0.Trim('\') -ErrorAction SilentlyContinue); if ($fo0.Count -eq 0) { continue }"
    psFile.WriteLine "            $fn0 = [string]$d0.FileName; if (-not $fn0) { $fn0 = [string]$d0.Name }"
    psFile.WriteLine "            $s0 = Get-PWDocumentsBySearch -FolderID $fo0[0].ProjectID -FileName $fn0 -GetAttributes -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object -First 1"
    psFile.WriteLine "            if (-not $s0) { continue }"
    psFile.WriteLine "            $ca0 = $s0.CustomAttributes; if (-not $ca0) { continue }"
    psFile.WriteLine "            if ($ca0 -is [System.Collections.IDictionary]) { $nm0 = @($ca0.Keys | ForEach-Object { [string]$_ }) } else { $nm0 = @($ca0.PSObject.Properties | ForEach-Object { $_.Name }) }"
    psFile.WriteLine "            $ne0 = @(); foreach ($n0 in $nm0) { if (-not $n0) { continue }; if ($ca0 -is [System.Collections.IDictionary]) { $v0 = $ca0[$n0] } else { $v0 = $ca0.PSObject.Properties[$n0].Value }; if ($v0 -is [array]) { $v0 = ($v0 -join ', ') }; $sv0 = ([string]$v0).Replace('|', '-').Replace([string][char]10, ' ').Replace([string][char]13, ''); if ($sv0.Trim() -ne '') { $ne0 += $n0; if (-not $foundAttrs.Contains($n0)) { $foundAttrs[$n0] = $sv0 } } }; [void]$newF.Add('F|' + $fk0 + '|' + ($ne0 -join [string][char]9))"
    psFile.WriteLine "        } catch { }"
    psFile.WriteLine "    }"
    psFile.WriteLine "    Write-Progress -Id 2 -Activity 'Checking attributes per subfolder' -Completed"
    psFile.WriteLine "    $newA = @($foundAttrs.Keys | Where-Object { $selectedAttrs -notcontains $_ -and $customAttrs -notcontains $_ -and $_ -notmatch '^(Name|Description|FileName|Updater|Updated)$' })"
    psFile.WriteLine "    Write-Host ('Subfolders checked: ' + $nF + ' | attributes found: ' + $foundAttrs.Count + ' | new: ' + $newA.Count + ' (' + $dw.Elapsed.ToString('mm\:ss') + ')') -ForegroundColor Green"
    psFile.WriteLine "    if ($newA.Count -gt 0) {"
    psFile.WriteLine "        $selectedAttrs = @($selectedAttrs + $newA); $customAttrs = @($customAttrs + $newA)"
    psFile.WriteLine "        if ($usedFast) { try { $cr = @{ ColumnsToReturn = $customAttrs }; $r2 = @(Get-PWDocumentsBySearchExtended -FolderPath $folderPath -PopulatePath @gv @cr -WarningAction SilentlyContinue -ErrorAction Stop); if ($r2.Count -gt 0 -and $r2[0] -isnot [System.Data.DataTable]) { foreach ($d in $r2) { if ([string]::IsNullOrWhiteSpace([string]$d.DocumentURN)) { $d | Add-Member -NotePropertyName DocumentURN -NotePropertyValue ('pw://' + $dsName + '/Documents/D{' + ([string]$d.DocumentGUIDString).Trim('{}').ToLower() + '}') -Force } }; $documents = $r2 } } catch { Write-Host 'Refetch with new attributes failed - they will be fetched on the next update.' -ForegroundColor Yellow } }"
    psFile.WriteLine "    }"
    psFile.WriteLine "    $ol = New-Object System.Collections.ArrayList; foreach ($k in $foundAttrs.Keys) { [void]$ol.Add('A|' + $k + '|' + $foundAttrs[$k]) }; foreach ($x in $newF) { [void]$ol.Add($x) }; if ($ol.Count -gt 0) { $ol | Out-File -FilePath ($OutputFile + '.attrs') -Encoding UTF8 -Append }"
    psFile.WriteLine "}"
    psFile.WriteLine " $maxSeq = @{}; foreach ($d in $documents) { $u = [string]$d.DocumentURN; $sq = [int]$d.VersionSequence; if (-not $maxSeq.ContainsKey($u) -or $sq -gt $maxSeq[$u]) { $maxSeq[$u] = $sq } }" ' Najwyzsza wersja kazdego URN
    psFile.WriteLine " Write-Host ('Found total documents: ' + $documents.Count + ' (fetch from PW: ' + $fetchWatch.Elapsed.ToString('mm\:ss') + ')') -ForegroundColor Green"
    psFile.WriteLine " $totalDocs = $documents.Count" ' Zapis kodu PS: Utworzenie zmiennej limitu dokumentów w folderze
    psFile.WriteLine " $docIndex = 0"               ' Zapis kodu PS: Licznik przetwarzanych dokumentów na pętli
    psFile.WriteLine " foreach ($doc in $documents) {" ' Zapis kodu PS: Rozpoczęcie iterowania po paczce pobranych dokumentów
    psFile.WriteLine " $docIndex++"                 ' Zapis kodu PS: Podbicie licznika pliku
    
    psFile.WriteLine " if ($docIndex % 500 -eq 0 -or $docIndex -eq $totalDocs) {" ' Zapis kodu PS: Aktualizuj pasek postępu co 500 dokumentów (wydajność renderowania powłoki!)
    psFile.WriteLine " Write-Progress -Id 1 -Activity ""Parsing files | Time elapsed: $($folderWatch.Elapsed.ToString('mm\:ss'))"" -Status ""Folder $folderIndex of $totalFolders | File $docIndex of $totalDocs"" -PercentComplete (($docIndex / $totalDocs) * 100)" ' Zapis kodu PS: Rysuje na nowo postęp %
    psFile.WriteLine " }"                           ' Zapis kodu PS: Koniec warunku dla GUI w terminalu
    psFile.WriteLine " $actualFolder = $rawFolder"  ' Zapis kodu PS: Odniesienie lokalizacji docelowej pliku
    psFile.WriteLine " if ($doc.FullPath) {"        ' Zapis kodu PS: Jeśli plik ma załadowaną ścieżkę absolutną
    psFile.WriteLine " $lastSlash = $doc.FullPath.LastIndexOf('\')" ' Zapis kodu PS: Zlokalizuj ukośnik przed plikiem
    psFile.WriteLine " if ($lastSlash -ge 0) { $actualFolder = $doc.FullPath.Substring(0, $lastSlash) }" ' Zapis kodu PS: Odseparuj nazwę pliku, zostawiając folder
    psFile.WriteLine " $idx2 = $actualFolder.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase)" ' Zapis kodu PS: Standaryzacja usuwająca systemowy prefiks
    psFile.WriteLine " if ($idx2 -ge 0) { $actualFolder = $actualFolder.Substring($idx2 + 11) } elseif ($actualFolder.StartsWith('Documents\', [System.StringComparison]::OrdinalIgnoreCase)) { $actualFolder = $actualFolder.Substring(10) }" ' Zapis kodu PS: Ciąg dalszy normalizacji
    psFile.WriteLine " $actualFolder = $actualFolder.Trim('\')" ' Zapis kodu PS: Oczyszczenie brzegów
    psFile.WriteLine " }"                           ' Zapis kodu PS: Zakończenie pobierania znormalizowanej ścieżki obiektu
    
    psFile.WriteLine " $obj = [ordered]@{}"         ' Zapis kodu PS: Utworzenie pustego, posortowanego słownika do eksportu (kolumny CSV)
    psFile.WriteLine " $obj['File Name'] = $doc.Name" ' Zapis kodu PS: Zapisanie standardowej nazwy pliku do obiektu
    psFile.WriteLine " $obj['PW Source Path'] = $actualFolder" ' Zapis kodu PS: Zapisanie znormalizowanej ścieżki
    psFile.WriteLine " $fClean = $pwFolder.FullPath.Trim('\')" ' Zapis kodu PS: Pobranie struktury serwera folderu
    psFile.WriteLine " if ($fClean.StartsWith('Documents\', [System.StringComparison]::OrdinalIgnoreCase)) { $fClean = $fClean.Substring(10) }" ' Zapis kodu PS: Usunięcie standardowego wpisu
    psFile.WriteLine " $fClean = $fClean.Replace('\', '/')" ' Zapis kodu PS: Zmiana sposobu ułożenia ukośników pod budowę URLa
    psFile.WriteLine " $fUrn = if ($pwFolder.ProjectURN) { $pwFolder.ProjectURN } elseif ($pwFolder.URN) { $pwFolder.URN } else { 'pw://' + $dsName + '/Documents/' + $fClean }" ' Zapis kodu PS: Zbudowanie lub wyciągnięcie unikalnego identyfikatora dla Folderu Głównego
    psFile.WriteLine " $obj['Folder Link'] = $fUrn" ' Zapis kodu PS: Dodanie linku folderowego
    psFile.WriteLine " $docClean = $doc.FullPath.Trim('\')" ' Zapis kodu PS: Czyszczenie ścieżki samego dokumentu
    psFile.WriteLine " if ($docClean.StartsWith('Documents\', [System.StringComparison]::OrdinalIgnoreCase)) { $docClean = $docClean.Substring(10) }" ' Zapis kodu PS: Obcięcie korzenia serwera
    psFile.WriteLine " $docClean = $docClean.Replace('\', '/')" ' Zapis kodu PS: Zamiana na standard sieciowy ukośników
    psFile.WriteLine " $docUrn = if ($doc.DocumentURN) { $doc.DocumentURN } elseif ($doc.URN) { $doc.URN } else { 'pw://' + $dsName + '/Documents/' + $docClean }" ' Zapis kodu PS: Generowanie unikalnego klucza hiperlinka URN konkretnego pliku
    psFile.WriteLine " if ($doc.DocumentURN -and [int]$doc.VersionSequence -lt $maxSeq[[string]$doc.DocumentURN]) { $docUrn = $docUrn + '?version=' + $doc.VersionSequence }" ' Starsza rewizja o tym samym URN dostaje dopisek, by klucz byl unikalny
    psFile.WriteLine " $obj['PW Link'] = $docUrn"   ' Zapis kodu PS: Wstawienie URN
    psFile.WriteLine " $uName = @($doc.DocumentUpdaterName, $doc.FileUpdaterName, $doc.UpdaterName, $doc.Updater, $doc.CreatorName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1" ' Zapis kodu PS: Sprytny dobór osoby edytującej, uwzględniający puste pola zależnie od wersji ProjectWise
    psFile.WriteLine " $obj['Updater'] = if ($uName) { $uName } else { 'Unknown (PW)' }" ' Zapis kodu PS: Wstawienie nazwy autora zmian
    psFile.WriteLine " $uDate = @($doc.DocumentUpdateDate, $doc.FileUpdateDate, $doc.UpdateDate, $doc.Updated, $doc.CreateDate) | Where-Object { $_ } | Select-Object -First 1" ' Zapis kodu PS: Identyczne zachowanie łapiące datę dla kompatybilności API wstecz
    psFile.WriteLine " $obj['Updated'] = if ($uDate -is [DateTime]) { $uDate.ToString('yyyy-MM-dd') } else { $uDate }" ' Zapis kodu PS: Ujednolicenie formatu do ROK-MC-DZIEŃ
    
    psFile.WriteLine " $caLoaded = $false; $ca = $null"     ' atrybuty projektu pobierane leniwie, tylko gdy nie ma ich wsrod zwyklych wlasciwosci (nie kopiujemy wszystkich wlasciwosci dokumentu)
    psFile.WriteLine " foreach ($attr in $selectedAttrs) {" ' Zapis kodu PS: Przechodzenie przez każdą nazwę atrybutu wpisaną w Excelu
    psFile.WriteLine " if ($attr -match '^(Name|Description|FileName|Updater|Updated)$') { continue }" ' Zapis kodu PS: Zablokowanie nadpisywania stałych systemowych, jeśli ktoś wpisał bzdury
    psFile.WriteLine " $val = $null; $pr = $doc.PSObject.Properties[$attr]; if ($pr) { $val = $pr.Value } else { if (-not $caLoaded) { $ca = $doc.CustomAttributes; $caLoaded = $true }; if ($ca) { if ($ca -is [System.Collections.IDictionary]) { if ($ca.Contains($attr)) { $val = $ca[$attr] } } else { $p2 = $ca.PSObject.Properties[$attr]; if ($p2) { $val = $p2.Value } } } }"
    psFile.WriteLine " if ($val -is [DateTime]) { $val = $val.ToString('yyyy-MM-dd') }" ' Zapis kodu PS: Normalizacja, jeśli atrybut PW był typu Data
    psFile.WriteLine " elseif ($val -is [array]) { $val = ($val -join ', ') }" ' Zapis kodu PS: Złączenie atrybutów wielokrotnych (np. wiele tagów) przecinkami w jeden wiersz
    psFile.WriteLine " if ($val -is [string]) { $val = $val.Replace('|', '-').Replace(""`n"", ' ').Replace(""`r"", '').Trim() }" ' Zapis kodu PS: Zabezpieczenie CSV przed zepsuciem (znak potoku to separator, więc jest usuwany z tekstu użytkownika PW)
    psFile.WriteLine " $obj[$attr] = $val"          ' Zapis kodu PS: Wstawienie ostatecznie sformatowanego atrybutu do kolumny
    psFile.WriteLine " }"                           ' Zapis kodu PS: Koniec pętli po właściwościach pojedynczego pliku
    psFile.WriteLine " [PSCustomObject]$obj"        ' Zapis kodu PS: Rzucenie wygenerowanego obiektu pliku na stertę (doda się automatycznie do tablicy wyników pętli na górze)
    psFile.WriteLine " }"                           ' Zapis kodu PS: Zakończenie pętli przechodzącej po dokumentach
    psFile.WriteLine " } else { Write-Host 'FOLDER NOT FOUND OR INVALID!' -ForegroundColor Red }" ' Zapis kodu PS: Błąd w terminalu w przypadku problemów z folderem
    psFile.WriteLine " } catch { Write-Host ('ERROR: ' + $_.Exception.Message) -ForegroundColor Red }" ' Zapis kodu PS: Czerwony błąd systemowy bazy dla pojedynczego wariantu

    psFile.WriteLine " $folderWatch.Stop()"         ' Zapis kodu PS: Zatrzymanie stopera lokalnego zadania
    psFile.WriteLine " Write-Host ('[TIMER] Folder processed in: ' + $folderWatch.Elapsed.ToString('hh\:mm\:ss')) -ForegroundColor Magenta" ' Zapis kodu PS: Wyświetlenie czasu wykonania operacji na jednej gałęzi folderowej
    psFile.WriteLine "})"                           ' Zapis kodu PS: Koniec wielkiej pętli nad folderami docelowymi
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz oddzielający kod pobierania i zapisu pliku
    psFile.WriteLine "if ($loggedDs) { try { Undo-PWLogin | Out-Null } catch {} }"
    psFile.WriteLine "Write-Progress -Id 1 -Activity 'Completed' -Completed" ' Zapis kodu PS: Usunięcie paska pobierania z terminala po zakończeniu
    psFile.WriteLine "Write-Host ('TOTAL DOCUMENTS: ' + $results.Count) -ForegroundColor Cyan" ' Zapis kodu PS: Wypisanie ile całościowo zrzucono rekordów z pamięci serwera do RAM
    
    psFile.WriteLine "if ($results.Count -gt 0) {"  ' Zapis kodu PS: Warunek jeśli cokolwiek ściągnięto do tablicy (wyniki są obecne)
    psFile.WriteLine " $results = $results | Sort-Object -Property 'PW Link' -Unique | Sort-Object 'PW Source Path', 'File Name'" ' Zapis kodu PS: Eliminacja duplikatów ze skanowania nakładających się folderów z użyciem Unikalnych URN oraz posortowanie alfabetycznie
    psFile.WriteLine " $results | Export-Csv -Path $OutputFile -NoTypeInformation -Delimiter '|' -Encoding UTF8" ' Zapis kodu PS: Zrzut danych CSV, z kodowaniem pozwalającym na polskie i chińskie znaki, odseparowany znakami potoku (|)
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz
    
    psFile.WriteLine " Write-Host 'Generating Structure Log...' -ForegroundColor Cyan" ' Zapis kodu PS: Informacja w konsoli o rozpoczęciu generowania mapy folderów (Txt)
    psFile.WriteLine " $logContent = New-Object System.Collections.ArrayList"
    psFile.WriteLine " [void]$logContent.Add('======================================================================')" ' Zapis kodu PS: Dodanie nagłówka separatora logu
    psFile.WriteLine " [void]$logContent.Add('PROJECTWISE STRUCTURE LOG')" ' Zapis kodu PS: Dodanie tytułu
    psFile.WriteLine " [void]$logContent.Add('Generated: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))" ' Zapis kodu PS: Wpisanie sygnatury czasowej wygenerowania
    psFile.WriteLine " [void]$logContent.Add('======================================================================')" ' Zapis kodu PS: Dodanie stopki nagłówka
    psFile.WriteLine " foreach ($ok in $okScopes) { [void]$logContent.Add('[SCOPE_OK] ' + $ok) }" ' Zapis kodu PS: Zapis listy poprawnie przeskanowanych folderów na początku logu
    psFile.WriteLine " $logContent | Out-File -FilePath $LogFile -Encoding UTF8 -Force" ' Zapis kodu PS: Fizyczne wyeksportowanie logu do pliku i nadpisanie w tle starych logów
    psFile.WriteLine "}"                            ' Zapis kodu PS: Zakończenie całego bloku if (obsługa wyników)
    
    psFile.WriteLine "$globalWatch.Stop()"          ' Zapis kodu PS: Zatrzymanie głównego timera terminalowego
    psFile.WriteLine "Write-Host '================================' -ForegroundColor DarkCyan" ' Zapis kodu PS: Dodanie ostatniego wizualnego oddzielenia okna
    psFile.WriteLine "Write-Host ('[TIMER] TOTAL EXECUTION TIME: ' + $globalWatch.Elapsed.ToString('hh\:mm\:ss')) -ForegroundColor Magenta" ' Zapis kodu PS: Informacja w oknie CMD o długości trwania skanowania bazy
    psFile.WriteLine "Set-Content -Path ($OutputFile + '.done') -Value 'done'" ' Znacznik zakonczenia dla Excela (pobieranie w tle)
    psFile.WriteLine "Write-Host 'Done! Window will close automatically in 3 seconds...' -ForegroundColor Green" ' Zapis kodu PS: Informacja o zautomatyzowanym zamknięciu, okno po prostu powiadomiło że zaraz "gaśnie"
    psFile.WriteLine "try { Stop-Transcript | Out-Null } catch {}"
    psFile.WriteLine "Start-Sleep -Seconds 3"      ' Zapis kodu PS: Opóźnienie zamknięcia procesu powłoki, by dało się przeczytać czas
    psFile.Close                                    ' VBA: Ostateczne zapisanie wygenerowanego kodu w pliku .ps1 i zwolnienie uchwytu
End Sub                                             ' Zakończenie modułu kodującego plik

Private Sub RunPowerShellScript(psPath As String, txtPath As String, csvPath As String, logPath As String) ' Procedura do wykonania wygenerowanego wcześniej pliku na silniku Windows PowerShell
    Dim wsh As Object                               ' Inicjalizacja referencji systemowej
    Set wsh = CreateObject("WScript.Shell")         ' Zainicjowanie interfejsu systemu (WScript do powłoki)
    Dim psCommand As String                         ' Zmienna przetrzymująca wywołanie polecenia dla konsoli
    psCommand = "powershell.exe -ExecutionPolicy Bypass -File """ & psPath & """ -InputFile """ & txtPath & """ -OutputFile """ & csvPath & """ -LogFile """ & logPath & """" ' Stworzenie instrukcji wyłączającej zasady bezpieczeństwa lokalnego skryptu oraz podającej mu 3 pliki jako parametry wejściowe/wyjściowe
    wsh.Run psCommand, 7, False                      ' Odpalenie zapytania w widocznym oknie (1) wymuszając oczekiwanie Excela na koniec jego działania (True)
End Sub                                             ' Zakończenie wywołania środowiska OS

' =======================================================================
' 5. DELTA ANALYSIS: ANALIZUJE CSV ORAZ PARSUJE PLIK TXT
' =======================================================================
' =======================================================================
' IMPORT + DELTA (na tablicach w pamieci, zapis zbiorczy w paczkach)
' PW_Data = plaska tabela: jeden wiersz = jeden plik (klucz: PW Link). Kolumny systemowe Status i Change Info oznaczaja zmiany z ostatniej aktualizacji.
' =======================================================================
Private Function ReadUtf8Text(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    If Not st.EOS Then ReadUtf8Text = st.ReadText
    st.Close
End Function

Private Function SplitCsvLine(ByVal s As String) As String()
    Dim p() As String, i As Long, t As String
    p = Split(s, "|")
    For i = 0 To UBound(p)
        t = p(i)
        If Len(t) >= 2 Then
            If Left(t, 1) = """" And Right(t, 1) = """" Then
                t = Mid(t, 2, Len(t) - 2)
                If InStr(t, """""") > 0 Then t = Replace(t, """""", """")
            End If
        End If
        p(i) = t
    Next i
    SplitCsvLine = p
End Function

' Czy wartosc sie realnie zmienila (ta sama data lub liczba zapisana inaczej nie liczy sie jako zmiana)
Private Function ValuesDiffer(ByVal a As String, ByVal b As String) As Boolean
    If StrComp(a, b, vbTextCompare) = 0 Then Exit Function
    ValuesDiffer = True
    On Error Resume Next
    If IsDate(a) And IsDate(b) Then
        If Int(CDate(a)) = Int(CDate(b)) Then ValuesDiffer = False: Exit Function
    End If
    If IsNumeric(a) And IsNumeric(b) Then
        If CDbl(Replace(a, ".", ",")) = CDbl(Replace(b, ".", ",")) Then ValuesDiffer = False
    End If
    On Error GoTo 0
End Function

' Prosty log pliku w kolumnie Change Info: nowy wpis na gorze, starsze pod spodem
Private Sub AddInfo(res() As Variant, ByVal ro As Long, ByVal iCh As Long, ByVal line As String)
    Dim old As String: old = CStr(res(ro, iCh))
    If old <> "" Then line = line & vbLf & old
    res(ro, iCh) = Left(line, 8000)
End Sub

Private Sub QueueLog(lst As Object, ByVal action As String, ByVal fName As String, ByVal attr As String, ByVal oldV As String, ByVal newV As String, ByVal user As String)
    lst.Add Array(Now, action, fName, attr, oldV, newV, user)
End Sub

Private Sub ProcessDeltaAnalysis(ws1 As Worksheet, csvPath As String, txtLogPath As String, numScopes As Long, ByVal arrCleanScopes As Variant)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(csvPath) Then
        AddLogEntry "ERROR", "Missing results CSV file.", "-", "-", "-", "-", CLR_DELETED
        Exit Sub
    End If
    On Error Resume Next
    If ws1.FilterMode Then ws1.ShowAllData
    ws1.Cells.ClearOutline
    ws1.Rows.Hidden = False
    On Error GoTo 0
    
    ' --- foldery faktycznie przeskanowane (tylko pod nimi wolno oznaczac pliki jako usuniete) ---
    Dim dictOk As Object, tx As String, tl() As String, i As Long, ln As String
    Set dictOk = CreateObject("Scripting.Dictionary"): dictOk.CompareMode = 1
    If fso.FileExists(txtLogPath) Then
        tx = Replace(Replace(ReadUtf8Text(txtLogPath), vbCrLf, vbLf), vbCr, vbLf)
        tl = Split(tx, vbLf)
        For i = 0 To UBound(tl)
            ln = Trim(tl(i))
            If InStr(1, ln, "[SCOPE_OK]", vbTextCompare) = 1 Then dictOk(UCase(Trim(Mid(ln, Len("[SCOPE_OK]") + 1)))) = True
        Next i
    End If
    
    ' --- CSV -> tablica ---
    Dim csv As String, lines() As String
    csv = ReadUtf8Text(csvPath)
    If Len(csv) = 0 Then Exit Sub
    lines = Split(Replace(Replace(csv, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim csvH() As String, nCsv As Long
    csvH = SplitCsvLine(lines(0))
    nCsv = UBound(lines)
    
    ' --- istniejaca baza -> tablica ---
    Dim oldC As Long, oldR As Long, oldArr As Variant, hv As Variant
    oldC = ws1.Cells(DATA_HEADER_ROW, ws1.Columns.Count).End(xlToLeft).Column
    If Trim(CStr(ws1.Cells(DATA_HEADER_ROW, oldC).Value)) = "" Then oldC = 0
    oldR = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    Dim m As Long
    If oldC > 0 And oldR > DATA_HEADER_ROW Then
        oldArr = ReadBlock(ws1, DATA_HEADER_ROW + 1, 1, oldR, oldC)
        m = UBound(oldArr, 1)
    End If
    
    ' --- uklad kolumn: 3 pierwsze kolumny danych, potem Change Info (zawsze 4.) i Status (5.), reszta; stare kolumny mapowane po nazwie ---
    Dim hdr() As String, nc As Long, hIdx As Object, nm0 As String
    Set hIdx = CreateObject("Scripting.Dictionary"): hIdx.CompareMode = 1
    Dim cand() As String, nCand As Long
    ReDim cand(1 To oldC + UBound(csvH) + 5)
    Dim oldH() As String, c As Long
    If oldC > 0 Then
        ReDim oldH(1 To oldC)
        hv = ReadBlock(ws1, DATA_HEADER_ROW, 1, DATA_HEADER_ROW, oldC)
        For i = 1 To oldC
            oldH(i) = Trim(CStr(hv(1, i)))
            If oldH(i) <> "" And oldH(i) <> COL_STATUS And oldH(i) <> COL_CHANGES Then nCand = nCand + 1: cand(nCand) = oldH(i)
        Next i
    End If
    For c = 0 To UBound(csvH)
        If csvH(c) <> "Folder Link" And csvH(c) <> COL_STATUS And csvH(c) <> COL_CHANGES Then
            nm0 = csvH(c)
            Dim dupe As Boolean: dupe = False
            For i = 1 To nCand
                If StrComp(cand(i), nm0, vbTextCompare) = 0 Then dupe = True: Exit For
            Next i
            If Not dupe Then nCand = nCand + 1: cand(nCand) = nm0
        End If
    Next c
    ReDim hdr(1 To nCand + 2)
    For i = 1 To nCand
        If i <= 3 Then nc = nc + 1: hdr(nc) = cand(i): hIdx(cand(i)) = nc
    Next i
    nc = nc + 1: hdr(nc) = COL_CHANGES: hIdx(COL_CHANGES) = nc
    nc = nc + 1: hdr(nc) = COL_STATUS: hIdx(COL_STATUS) = nc
    For i = 4 To nCand
        nc = nc + 1: hdr(nc) = cand(i): hIdx(cand(i)) = nc
    Next i
    Dim csvMap() As Long, oldMap() As Long
    ReDim csvMap(0 To UBound(csvH))
    For c = 0 To UBound(csvH)
        If hIdx.exists(csvH(c)) Then csvMap(c) = hIdx(csvH(c)) Else csvMap(c) = 0
    Next c
    If oldC > 0 Then
        ReDim oldMap(1 To oldC)
        For c = 1 To oldC
            If hIdx.exists(oldH(c)) Then oldMap(c) = hIdx(oldH(c)) Else oldMap(c) = 0
        Next c
    End If
    Dim iName As Long, iPath As Long, iLink As Long, iUpd As Long, iSt As Long, iCh As Long
    iName = hIdx("File Name"): iPath = hIdx("PW Source Path"): iLink = hIdx("PW Link"): iUpd = hIdx("Updater")
    iSt = hIdx(COL_STATUS): iCh = hIdx(COL_CHANGES)
    
    ' --- wynik: stare wiersze (bez smieci i dawnych wierszy folderow) + nowe ---
    Dim res() As Variant, cnt As Long, cap As Long, r As Long
    cap = m + nCsv + 1
    ReDim res(1 To cap, 1 To nc)
    Dim dIdx As Object, key As String, pth As String, nm As String
    Dim oNm As Long, oPa As Long
    For c = 1 To oldC
        If StrComp(oldH(c), "File Name", vbTextCompare) = 0 Then oNm = c
        If StrComp(oldH(c), "PW Source Path", vbTextCompare) = 0 Then oPa = c
    Next c
    Set dIdx = CreateObject("Scripting.Dictionary"): dIdx.CompareMode = 1
    For r = 1 To m
        nm = Trim(CStr(oldArr(r, oNm))): pth = Trim(CStr(oldArr(r, oPa)))
        If nm <> "" And pth <> "" Then               ' pusty path = stary wiersz folderu - pomijamy
            cnt = cnt + 1
            For c = 1 To oldC
                If oldMap(c) > 0 Then res(cnt, oldMap(c)) = oldArr(r, c)
            Next c
            If CStr(res(cnt, iSt)) <> "DELETED" Then res(cnt, iSt) = ""
            key = Trim(CStr(res(cnt, iLink)))
            If key = "" Then key = pth & "|" & nm
            dIdx(key) = cnt
        End If
    Next r
    Dim nOld As Long: nOld = cnt
    Dim seen() As Boolean
    ReDim seen(1 To cap)
    
    Dim logs As Object, nAdd As Long, nChg As Long, nChgCells As Long, nDel As Long, nRestored As Long
    Set logs = CreateObject("System.Collections.ArrayList")
    Dim stamp As String: stamp = Format(Now, "yyyy-mm-dd hh:nn")
    Dim currentUpdater As String: currentUpdater = Environ("USERNAME")
    
    Dim f() As String, ro As Long, nv As String, ov As String, t As Long, chg As String, chgTxt As String
    Application.StatusBar = "Merging " & nCsv & " records..."
    For r = 1 To nCsv
        If Len(lines(r)) > 0 Then
            f = SplitCsvLine(lines(r))
            key = ""
            For c = 0 To UBound(csvH)
                If csvH(c) = "PW Link" And c <= UBound(f) Then key = Trim(f(c))
            Next c
            If key = "" Then key = "?" & r
            ro = 0
            If dIdx.exists(key) Then ro = dIdx(key)
            If ro > 0 Then
                seen(ro) = True
                chgTxt = ""
                If CStr(res(ro, iSt)) = "DELETED" Then res(ro, iSt) = "": AddInfo res, ro, iCh, stamp & "  Restored in PW": nRestored = nRestored + 1
                For c = 0 To UBound(csvH)
                    t = csvMap(c)
                    If t > 0 And c <= UBound(f) Then
                        nv = Trim(f(c)): ov = Trim(CStr(res(ro, t)))
                        If t = iLink Or t = iPath Or t = iUpd Then
                            res(ro, t) = nv                 ' pola techniczne - bez oznaczania zmian
                        ElseIf ValuesDiffer(ov, nv) Then
                            res(ro, t) = nv
                            If ov <> "" Then
                                nChgCells = nChgCells + 1
                                chgTxt = chgTxt & csvH(c) & ": " & ov & " -> " & nv & " | "
                                If nChgCells <= 3000 Then QueueLog logs, "CHANGED", CStr(res(ro, iName)), csvH(c), ov, nv, currentUpdater
                            End If
                        End If
                    End If
                Next c
                If chgTxt <> "" Then
                    nChg = nChg + 1
                    res(ro, iSt) = "CHANGED"
                    AddInfo res, ro, iCh, stamp & "  " & chgTxt
                End If
            Else
                cnt = cnt + 1: nAdd = nAdd + 1
                For c = 0 To UBound(csvH)
                    t = csvMap(c)
                    If t > 0 And c <= UBound(f) Then res(cnt, t) = Trim(f(c))
                Next c
                res(cnt, iSt) = "NEW": AddInfo res, cnt, iCh, stamp & "  Added"
                dIdx(key) = cnt: seen(cnt) = True
                If nOld > 0 And nAdd <= 1000 Then QueueLog logs, "ADDED", CStr(res(cnt, iName)), "-", "-", "-", currentUpdater
            End If
        End If
    Next r
    If nAdd > 0 And (nOld = 0 Or nAdd > 1000) Then QueueLog logs, "ADDED", nAdd & " files" & IIf(nOld = 0, " (initial load)", " (bulk)"), "-", "-", "-", currentUpdater
    If nChgCells > 3000 Then QueueLog logs, "CHANGED", "... " & (nChgCells - 3000) & " more changed values not listed", "-", "-", "-", currentUpdater
    
    ' --- usuniete: nie widziane w CSV, ale tylko pod folderami faktycznie przeskanowanymi ---
    Dim scopeCache As Object, underScope As Boolean, s As Long, cs As String, sp As String
    Set scopeCache = CreateObject("Scripting.Dictionary"): scopeCache.CompareMode = 1
    For r = 1 To nOld
        If Not seen(r) Then
            pth = UCase(Replace(Trim(CStr(res(r, iPath))), "/", "\"))
            If scopeCache.exists(pth) Then
                underScope = scopeCache(pth)
            Else
                underScope = False
                For s = 1 To numScopes
                    cs = UCase(Replace(arrCleanScopes(s), "/", "\"))
                    If dictOk.exists(cs) And InStr(1, pth, cs, vbTextCompare) > 0 Then underScope = True: Exit For
                Next s
                scopeCache(pth) = underScope
            End If
            If underScope And CStr(res(r, iSt)) <> "DELETED" Then
                res(r, iSt) = "DELETED": AddInfo res, r, iCh, stamp & "  Deleted from PW"
                nDel = nDel + 1
                If nDel <= 2000 Then QueueLog logs, "DELETED", CStr(res(r, iName)), "-", "Existed", "Deleted from PW", "-"
            End If
        End If
    Next r
    If nDel > 2000 Then QueueLog logs, "DELETED", "... " & (nDel - 2000) & " more deleted files not listed", "-", "-", "-", "-"
    
    Dim sc As Long
    For sc = 1 To numScopes
        If Not dictOk.exists(arrCleanScopes(sc)) Then AddLogEntry "WARNING", "Scope not scanned (skipped deletion check): " & arrCleanScopes(sc), "-", "-", "-", "-", 0
    Next sc
    
    ' --- zapis bazy: naglowki + dane w paczkach ---
    Application.StatusBar = "Writing " & cnt & " rows to " & SH_DATA & "..."
    Dim clearTo As Long
    clearTo = oldR: If DATA_HEADER_ROW + cnt > clearTo Then clearTo = DATA_HEADER_ROW + cnt
    ws1.Range(ws1.Rows(DATA_HEADER_ROW + 1), ws1.Rows(clearTo + 5)).Clear
    Dim hrow() As Variant
    ReDim hrow(1 To 1, 1 To nc)
    For c = 1 To nc: hrow(1, c) = hdr(c): Next c
    ws1.Range(ws1.Cells(DATA_HEADER_ROW, 1), ws1.Cells(DATA_HEADER_ROW, nc)).Value = hrow
    StyleHeader ws1.Range(ws1.Cells(DATA_HEADER_ROW, 1), ws1.Cells(DATA_HEADER_ROW, nc))
    If cnt > 0 Then
        ws1.Range(ws1.Cells(DATA_HEADER_ROW + 1, 1), ws1.Cells(DATA_HEADER_ROW + cnt, nc)).NumberFormat = "@"
        Dim r0 As Long, r1 As Long, blk() As Variant, j As Long
        Const CHUNK As Long = 4000
        For r0 = 1 To cnt Step CHUNK
            r1 = r0 + CHUNK - 1: If r1 > cnt Then r1 = cnt
            ReDim blk(1 To r1 - r0 + 1, 1 To nc)
            For r = r0 To r1
                For j = 1 To nc: blk(r - r0 + 1, j) = res(r, j): Next j
            Next r
            ws1.Range(ws1.Cells(DATA_HEADER_ROW + r0, 1), ws1.Cells(DATA_HEADER_ROW + r1, nc)).Value = blk
            Application.StatusBar = "Writing " & SH_DATA & ": " & r1 & " / " & cnt
        Next r0
    End If
    On Error Resume Next
    ws1.AutoFilterMode = False
    If cnt > 0 Then ws1.Range(ws1.Cells(DATA_HEADER_ROW, 1), ws1.Cells(DATA_HEADER_ROW + cnt, nc)).AutoFilter
    If oldC = 0 Then
        ws1.Columns(iName).ColumnWidth = 48: ws1.Columns(iPath).ColumnWidth = 55: ws1.Columns(iLink).ColumnWidth = 46
        ws1.Columns(iUpd).ColumnWidth = 18
    End If
    ws1.Cells(DATA_HEADER_ROW + 1, 1).Parent.Range(ws1.Cells(DATA_HEADER_ROW + 1, iSt), ws1.Cells(DATA_HEADER_ROW + cnt + 1, iSt)).FormatConditions.Delete
    Dim fcR As Range
    Set fcR = ws1.Range(ws1.Cells(DATA_HEADER_ROW + 1, iSt), ws1.Cells(DATA_HEADER_ROW + cnt + 1, iSt))
    fcR.FormatConditions.Add(xlCellValue, xlEqual, "=""NEW""").Interior.Color = CLR_NEW
    fcR.FormatConditions.Add(xlCellValue, xlEqual, "=""CHANGED""").Interior.Color = CLR_CHANGED
    fcR.FormatConditions.Add(xlCellValue, xlEqual, "=""DELETED""").Interior.Color = CLR_DELETED
    On Error GoTo 0
    SetDataCache res, hdr, cnt, nc
    
    ' --- log: zbiorczo ---
    Dim nl As Long: nl = logs.Count
    If nl > 0 Then
        Dim la() As Variant, k As Long, ent As Variant
        ReDim la(1 To nl, 1 To 7)
        For k = 1 To nl
            ent = logs(k - 1)
            For j = 0 To 6: la(k, j + 1) = ent(j): Next j
            la(k, 5) = "'" & CStr(la(k, 5)): la(k, 6) = "'" & CStr(la(k, 6))
        Next k
        wsLog.Range(wsLog.Cells(nextLogRow, 1), wsLog.Cells(nextLogRow + nl - 1, 7)).Value = la
        Dim a0 As Long, aCol As Long
        For k = 1 To nl                              ' kolory wg akcji (pasy tego samego koloru naraz)
            Select Case CStr(la(k, 2))
                Case "ADDED": aCol = RGB(226, 239, 218)
                Case "CHANGED": aCol = RGB(255, 242, 204)
                Case "DELETED": aCol = CLR_DELETED
                Case Else: aCol = 0
            End Select
            If k = 1 Then a0 = 1
            If k = nl Then
                If aCol <> 0 Then wsLog.Range(wsLog.Cells(nextLogRow + a0 - 1, 1), wsLog.Cells(nextLogRow + k - 1, 7)).Interior.Color = aCol
            ElseIf CStr(la(k + 1, 2)) <> CStr(la(k, 2)) Then
                If aCol <> 0 Then wsLog.Range(wsLog.Cells(nextLogRow + a0 - 1, 1), wsLog.Cells(nextLogRow + k - 1, 7)).Interior.Color = aCol
                a0 = k + 1
            End If
        Next k
        nextLogRow = nextLogRow + nl
    End If
    If nAdd = 0 And nChg = 0 And nDel = 0 Then AddLogEntry "INFO", "No changes compared to the previous download.", "-", "-", "-", "-", 0
    AddLogEntry "INFO", "Files: " & cnt & "  |  new: " & nAdd & "  |  changed: " & nChg & "  |  deleted: " & nDel, "-", "-", "-", "-", 0
    If nextLogRow - 1 > logStartRow Then wsLog.Rows((logStartRow + 1) & ":" & (nextLogRow - 1)).Group
End Sub

Public Sub StyleHeader(rng As Range)                ' Wspólny, profesjonalny styl nagłówków tabel
    With rng
        .Font.Bold = True
        .Font.Color = CLR_WHITE
        .Interior.Color = CLR_HEADER
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
End Sub

Private Sub DeleteIfExists(filePath As String)      ' Bezpieczne usunięcie pliku (brak błędu, gdy go nie ma)
    On Error Resume Next
    Kill filePath
    On Error GoTo 0
End Sub

Private Sub CleanUpFiles(txtPath As String, psPath As String, csvPath As String) ' Funkcja systemowa. Likwidator tymczasowych obiektów po całej brudnej "robocie CSV vs Excel". Generowane pod Powłoką pliki idą w Kosmos.
    Dim fso As Object                               ' Referencje dysku
    Set fso = CreateObject("Scripting.FileSystemObject") ' Uruchom dostęp do operacji usuwań File Sys OS (C++ wewnątrz VBA od Windows)
    
    ' ZGODNIE Z KOREKTĄ UŻYTKOWNIKA [2026-03-19] - Makro musi pytać przed usunięciem czegokolwiek z dysku
    If False Then                                   ' pliki tymczasowe zostaja (bez popupu); nadpisuje je kolejne pobranie ' Spytaj MsgBoxem przed skasowaniem by uratować logi dla np trybu developerskiego lub inspekcji błędów
        On Error Resume Next                        ' Odpalenie systemu ochronnego przed błędem w kasowaniu (np. PowerShell wciąz wisi i powłoka zblokowała plik przed Exelem pod LockFile)
        fso.DeleteFile txtPath                      ' Strzel po poleceniu zniszczenia pliku tekstowego z Drzewem Scopes z Dysku C:\
        fso.DeleteFile psPath                       ' Usuń również plik Generatora skryptu Powershell *.ps1
        fso.DeleteFile csvPath                      ' Wykasuj ogromny plik CSV ze strumienia by zwolnić potworne zapychacze pamięci po całej robocie
        fso.DeleteFile csvPath & ".done"
        On Error GoTo 0                             ' Ściągnij spadochron błędu. Od teraz powrót by wyrzucało makro po normalnemu
    End If                                          ' Zakończenie zgody usera od wariantu MsgBox
End Sub                                             ' Procedura poszła do końca.

' =======================================================================
' KOPIOWANIE NAGŁÓWKÓW
' =======================================================================
Sub CopyHeadersFromSheet3()
    Dim ws1 As Worksheet, ws3 As Worksheet
    Dim r As Long, cCol As Long, lastRow3 As Long
    Set ws1 = ThisWorkbook.Sheets(SH_DATA)
    Set ws3 = ThisWorkbook.Sheets(SH_ATTRS)
    
    Const START_COL_S1 As Long = 6                  ' Atrybuty zaczynają się w kolumnie F (A-E to stałe kolumny systemowe)
    
    lastRow3 = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row ' Czytamy tylko do faktycznie ostatniego atrybutu (wcześniej sztywno do wiersza 100)
    
    cCol = START_COL_S1
    For r = ATTRS_FIRST_ROW To lastRow3
        If Trim(CStr(ws3.Cells(r, 1).Value)) <> "" And Not IsNo(ws3.Cells(r, ATTRS_ENABLE_COL).Value) Then ' Puste wiersze są pomijane - tak samo jak w GetAttributesFromSheet3, więc kolejność kolumn zgadza się z CSV
            ws1.Cells(11, cCol).Value = Trim(CStr(ws3.Cells(r, 1).Value))
            StyleHeader ws1.Cells(11, cCol)
            cCol = cCol + 1
        End If
    Next r
End Sub

' Zapisuje znane podfoldery (arkusz AttrMap) do pliku dla PowerShella: folder|atrybuty
Private Sub WriteAttrMap(ByVal path As String)
    Dim fso As Object, f As Object, ws As Worksheet, r As Long, lastR As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(path, True, True)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_ATTRMAP)
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For r = 2 To lastR
            If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Then f.WriteLine CStr(ws.Cells(r, 1).Value) & "|" & CStr(ws.Cells(r, 2).Value)
        Next r
    End If
    f.Close
End Sub

' =======================================================================
' DOKLADNA ANALIZA ATRYBUTOW (w tle): czyta WSZYSTKIE pliki ze wszystkich folderow z arkusza Folders i zbiera atrybuty,
' ktore maja wartosc. Bez limitu czasu; Excel jest wolny, okno PowerShella pracuje w tle. Wynik: arkusz Attributes + ukryty AttrMap.
' =======================================================================
Public Sub StartAttrScan()
    If gBusy Or gScanBusy Then
        Notify "A ProjectWise job is already running in the background. Wait for it to finish."
        Exit Sub
    End If
    If ThisWorkbook.Path = "" Then Notify "Please save the Excel workbook to your drive first!": Exit Sub
    Dim ws2 As Worksheet, numScopes As Long, arr As Variant, base As String, fso As Object, psFile As Object
    Set ws2 = ThisWorkbook.Sheets(SH_FOLDERS)
    base = ThisWorkbook.Path
    If Not PrepareScopesAndTxtFile(ws2, base & "\PW_Folders.txt", numScopes, arr) Then Exit Sub
    If numScopes = 0 Then Notify "No active folder link in sheet " & SH_FOLDERS & ".": Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    gScanOut = base & "\PW_AttrScan.txt"
    On Error Resume Next
    fso.DeleteFile gScanOut: fso.DeleteFile gScanOut & ".done"
    On Error GoTo 0
    Set psFile = fso.CreateTextFile(base & "\PW_AttrScan.ps1", True, True)
    psFile.WriteLine "param([string]$InputFile, [string]$OutputFile)"
    psFile.WriteLine "$ErrorActionPreference = 'Stop'"
    psFile.WriteLine "$sw = [System.Diagnostics.Stopwatch]::StartNew()"
    WritePsPrelude psFile, "attribute-scan"
    psFile.WriteLine "$sw = [System.Diagnostics.Stopwatch]::StartNew()"
    psFile.WriteLine "Say-Head 'Attribute scan (ProjectWise environments)'"
    psFile.WriteLine "Say 'Step 1/4  Loading ProjectWise module' 'Cyan'"
    psFile.WriteLine "Import-Module PWPS_DAB -Force -DisableNameChecking 3>$null 4>$null"
    psFile.WriteLine "$in = @(Get-Content $InputFile | Where-Object { $_.Trim() -ne '' })"
    psFile.WriteLine "$progFile = $OutputFile + '.progress'"
    psFile.WriteLine "function Prog([string]$t) { try { Set-Content -Path $progFile -Value $t -Encoding UTF8 } catch {} }"
    psFile.WriteLine "function Fmt([double]$s) { if ($s -lt 0) { $s = 0 }; [TimeSpan]::FromSeconds([math]::Round($s)).ToString('hh\:mm\:ss') }"
    psFile.WriteLine "function ItemName($x) {"
    psFile.WriteLine "    if ($null -eq $x) { return $null }"
    psFile.WriteLine "    if ($x -is [string]) { return $x }"
    psFile.WriteLine "    foreach ($pn in 'AttributeName','ColumnName','Name','DisplayName','Property') { if ($x.PSObject.Properties[$pn] -and $x.$pn) { return [string]$x.$pn } }"
    psFile.WriteLine "    return $null"
    psFile.WriteLine "}"
    psFile.WriteLine "$found = [ordered]@{}; $envCount = 0"
    psFile.WriteLine "$dsName = ([string]$in[0]).Split('|')[0].Trim()"
    psFile.WriteLine "Say-Head 'Step 2/4  Logging in'"
    psFile.WriteLine "try { New-PWLogin -DatasourceName $dsName -BentleyIMS -NonAdminLogin | Out-Null } catch { Say ('ERROR: login failed: ' + $_.Exception.Message) 'Red'; Start-Sleep -Seconds 8; exit 1 }"
    psFile.WriteLine "Say 'Step 3/4  Reading the environments and their attributes' 'Yellow'"
    psFile.WriteLine "$fpIn = ([string]$in[0].Split('|')[1]).Trim(); $ix = $fpIn.IndexOf('\Documents\', [System.StringComparison]::OrdinalIgnoreCase); if ($ix -ge 0) { $fpIn = $fpIn.Substring($ix + 11) } elseif ($fpIn.StartsWith('Documents\', [System.StringComparison]::OrdinalIgnoreCase)) { $fpIn = $fpIn.Substring(10) }; $fpIn = $fpIn.Trim('\')"
    psFile.WriteLine "$envNames = [ordered]@{}; $envF = @{}; $envAttrs = @{}; $folderEnv = @{}; $folderObj = @{}"
    psFile.WriteLine "Say '          Probing which environment cmdlets are allowed for this login:' 'Cyan'"
    psFile.WriteLine "foreach ($cn in 'Get-PWEnvironments','Get-PWEnvironmentColumns','Get-PWAttributeSettings','Get-PWTagSetAttributes','Get-PWFolders') {"
    psFile.WriteLine "    $cmd = Get-Command $cn -ErrorAction SilentlyContinue"
    psFile.WriteLine "    if (-not $cmd) { Say ('            ' + $cn + ': not available') 'DarkGray'; continue }"
    psFile.WriteLine "    Say ('            ' + $cn + '  parameters: ' + (($cmd.Parameters.Keys | Where-Object { $_ -notmatch 'Verbose|Debug|Action|Variable|Buffer|Confirm|WhatIf' }) -join ', ')) 'DarkGray'"
    psFile.WriteLine "}"
    psFile.WriteLine "$q = New-Object System.Collections.Queue; $seen = 0; $t0 = Get-Date"
    psFile.WriteLine "try {"
    psFile.WriteLine "    foreach ($f0 in @(Get-PWFolders -FolderPath $fpIn -ErrorAction Stop)) { $q.Enqueue($f0) }"
    psFile.WriteLine "    if ($q.Count -gt 0) { Say ('          Root folder: Environment=' + $q.Peek().Environment + ', SubFolders=' + @($q.Peek().SubFolders).Count) 'DarkGray' }"
    psFile.WriteLine "} catch { Say ('          Get-PWFolders failed: ' + $_.Exception.Message) 'DarkYellow' }"
    psFile.WriteLine "while ($q.Count -gt 0) {"
    psFile.WriteLine "    $f0 = $q.Dequeue(); $seen++; $pp = [string]$f0.FullPath; $folderEnv[$pp] = [string]$f0.Environment; $folderObj[$pp] = $f0"
    psFile.WriteLine "    $ev = [string]$f0.Environment; if ($ev) { if (-not $envNames.Contains($ev)) { $envNames[$ev] = 1; $envF[$ev] = New-Object System.Collections.ArrayList; Say ('          Environment found: ' + $ev + '  (folder ' + $f0.Name + ')') 'Green' }; [void]$envF[$ev].Add($f0) }"
    psFile.WriteLine "    if ($seen % 50 -eq 0) { $m = 'Walking folders: ' + $seen + ' visited | ' + $q.Count + ' queued | environments: ' + $envNames.Count + ' | elapsed ' + (Fmt ((Get-Date) - $t0).TotalSeconds); Write-Progress -Activity 'Attribute scan - step 3/4: walking folders' -Status $m; Prog ('Attribute scan - ' + $m) }"
    psFile.WriteLine "    foreach ($sf in @($f0.SubFolders)) {"
    psFile.WriteLine "        if ($null -eq $sf) { continue }"
    psFile.WriteLine "        if ($sf.PSObject.Properties['Environment']) { $q.Enqueue($sf) }"
    psFile.WriteLine "        elseif ($sf.PSObject.Properties['ProjectID']) { try { foreach ($x in @(Get-PWFolders -FolderID $sf.ProjectID -JustOne -ErrorAction Stop)) { $q.Enqueue($x) } } catch {} }"
    psFile.WriteLine "    }"
    psFile.WriteLine "}"
    psFile.WriteLine "Say ('          Folders visited: ' + $seen) 'DarkGray' "
    psFile.WriteLine "function Add-Attrs($s0, $ev) {"
    psFile.WriteLine "    $ca = $s0.CustomAttributes; if (-not $ca) { return }"
    psFile.WriteLine "    if ($ca -is [System.Collections.IDictionary]) { $names = @($ca.Keys | ForEach-Object { [string]$_ }) } else { $names = @($ca.PSObject.Properties | ForEach-Object { $_.Name }) }"
    psFile.WriteLine "    foreach ($nm in $names) {"
    psFile.WriteLine "        if (-not $nm) { continue }"
    psFile.WriteLine "        if ($ca -is [System.Collections.IDictionary]) { $v = $ca[$nm] } else { $v = $ca.PSObject.Properties[$nm].Value }"
    psFile.WriteLine "        if ($v -is [array]) { $v = ($v -join ', ') }"
    psFile.WriteLine "        $sv = ([string]$v).Replace('|', '-').Replace([string][char]10, ' ').Replace([string][char]13, '')"
    psFile.WriteLine "        if ($sv.Trim() -ne '') { $envAttrs[$ev][$nm] = 1; if (-not $found.Contains($nm)) { $found[$nm] = $sv } }"
    psFile.WriteLine "    }"
    psFile.WriteLine "}"
    psFile.WriteLine "$n = $envNames.Count"
    psFile.WriteLine "Say ('Listing file names of the whole scope (fast request, about 2 min) - one file will be read per environment (' + $n + ')') 'Yellow'"
    psFile.WriteLine "$sb = { param($ds, $fp); Import-Module PWPS_DAB -Force -DisableNameChecking 3>$null 4>$null; New-PWLogin -DatasourceName $ds -BentleyIMS -NonAdminLogin | Out-Null; $r = @(Get-PWDocumentsBySearchExtended -FolderPath $fp -PopulatePath -WarningAction SilentlyContinue -ErrorAction Stop); try { Undo-PWLogin | Out-Null } catch {}; ,$r }"
    psFile.WriteLine "$tl = Get-Date; $docs = $null; $lastSay = 0"
    psFile.WriteLine "try {"
    psFile.WriteLine "    $ps = [powershell]::Create(); [void]$ps.AddScript($sb).AddArgument($dsName).AddArgument($fpIn)"
    psFile.WriteLine "    $ar = $ps.BeginInvoke()"
    psFile.WriteLine "    while (-not $ar.IsCompleted) {"
    psFile.WriteLine "        Start-Sleep -Seconds 2"
    psFile.WriteLine "        $el = ((Get-Date) - $tl).TotalSeconds; $msg = 'Listing files from PW: elapsed ' + (Fmt $el)"
    psFile.WriteLine "        Write-Progress -Activity 'Attribute scan - listing files' -Status $msg"
    psFile.WriteLine "        Prog ('Attribute scan - ' + $msg)"
    psFile.WriteLine "        if ($el - $lastSay -ge 60) { Say ('          ... still listing (' + $msg + ')') 'DarkGray'; $lastSay = $el }"
    psFile.WriteLine "    }"
    psFile.WriteLine "    $res = $ps.EndInvoke($ar)"
    psFile.WriteLine "    foreach ($er in $ps.Streams.Error) { Say ('          background request error: ' + $er) 'Red' }"
    psFile.WriteLine "    if ($res -and $res.Count -gt 0 -and $null -ne $res[0]) { $docs = @($res[0]) }"
    psFile.WriteLine "    $ps.Dispose()"
    psFile.WriteLine "} catch { Say ('          background request failed: ' + $_.Exception.Message) 'Red' }"
    psFile.WriteLine "if ($null -eq $docs) {"
    psFile.WriteLine "    Say '          Falling back to a direct request in this window' 'DarkYellow'"
    psFile.WriteLine "    try { $docs = @(Get-PWDocumentsBySearchExtended -FolderPath $fpIn -PopulatePath -WarningAction SilentlyContinue -ErrorAction Stop) } catch { Say ('          ERROR: ' + $_.Exception.Message) 'Red'; $docs = @() }"
    psFile.WriteLine "}"
    psFile.WriteLine "Say ('Listed ' + $docs.Count + ' files in ' + (Fmt ((Get-Date) - $tl).TotalSeconds)) 'Green'"
    psFile.WriteLine "$pick = @{}"
    psFile.WriteLine "foreach ($d in $docs) {"
    psFile.WriteLine "    $fp0 = [string]$d.FullPath; $ls = $fp0.LastIndexOf('\'); if ($ls -lt 0) { continue }"
    psFile.WriteLine "    $fk = $fp0.Substring(0, $ls); $ev = $folderEnv[$fk]"
    psFile.WriteLine "    if ($ev -and -not $pick.ContainsKey($ev)) { $pick[$ev] = @($d, $fk) }"
    psFile.WriteLine "}"
    psFile.WriteLine "Say ('Environments with at least one file: ' + $pick.Count + ' / ' + $n) 'Cyan'"
    psFile.WriteLine "$i = 0; $t1 = Get-Date"
    psFile.WriteLine "foreach ($ev in @($envNames.Keys)) {"
    psFile.WriteLine "    $i++; $envAttrs[$ev] = @{}"
    psFile.WriteLine "    $ela = ((Get-Date) - $t1).TotalSeconds; $eta = if ($i -gt 1) { $ela / ($i - 1) * ($n - $i + 1) } else { 0 }"
    psFile.WriteLine "    $stat = 'Environment ' + $i + ' / ' + $n + ' [' + $ev + '] | attributes: ' + $found.Count + ' | elapsed ' + (Fmt $ela) + ' | remaining ~' + (Fmt $eta)"
    psFile.WriteLine "    Write-Progress -Activity 'Attribute scan - reading one file per environment' -Status $stat -PercentComplete ([int](100 * $i / [math]::Max(1, $n)))"
    psFile.WriteLine "    Prog ('Attribute scan - ' + $stat)"
    psFile.WriteLine "    if (-not $pick.ContainsKey($ev)) { Say ('          [' + $i + '/' + $n + '] ' + $ev + ': no files in this environment') 'DarkYellow'; continue }"
    psFile.WriteLine "    $d = $pick[$ev][0]; $fo = $folderObj[$pick[$ev][1]]; $fn = [string]$d.FileName; if (-not $fn) { $fn = [string]$d.Name }"
    psFile.WriteLine "    try { $s0 = Get-PWDocumentsBySearch -FolderID $fo.ProjectID -FileName $fn -GetAttributes -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object -First 1; if ($s0) { Add-Attrs $s0 $ev } } catch { Say ('          read failed: ' + $_.Exception.Message) 'DarkYellow' }"
    psFile.WriteLine "    Say ('          [' + $i + '/' + $n + '] ' + $ev + ': ' + $fn + ' -> ' + $envAttrs[$ev].Count + ' attributes with values') 'DarkGray'"
    psFile.WriteLine "    $envCount++"
    psFile.WriteLine "}"
    psFile.WriteLine "if ($found.Count -eq 0) { Say 'No attributes read. Send me the lines above (parameters / properties / messages) and I will adapt the scan.' 'Red' }"
    psFile.WriteLine "try { Undo-PWLogin | Out-Null } catch {}"
    psFile.WriteLine "Write-Progress -Activity 'Attribute scan' -Completed"
    psFile.WriteLine "Say-Head 'Step 4/4  Saving the result'"
    psFile.WriteLine "$ol = New-Object System.Collections.ArrayList"
    psFile.WriteLine "foreach ($k in $found.Keys) { [void]$ol.Add('A|' + $k + '|' + $found[$k]) }"
    psFile.WriteLine "foreach ($ev in $envF.Keys) { $al = (@($envAttrs[$ev].Keys) -join [string][char]9); foreach ($fo in $envF[$ev]) { [void]$ol.Add('F|' + [string]$fo.FullPath + '|' + $al) } }"
    psFile.WriteLine "$ol | Out-File -FilePath $OutputFile -Encoding UTF8 -Force"
    psFile.WriteLine "Set-Content -Path ($OutputFile + '.done') -Value 'done'"
    psFile.WriteLine "Say-Head 'SUMMARY'"
    psFile.WriteLine "Write-Host ('   Environments read      : ' + $envCount) -ForegroundColor White"
    psFile.WriteLine "Write-Host ('   Attributes found       : ' + $found.Count) -ForegroundColor White"
    psFile.WriteLine "Write-Host ('   Total time             : ' + (Fmt $sw.Elapsed.TotalSeconds)) -ForegroundColor White"
    psFile.WriteLine "Write-Host ('   Log file               : ' + $psLogFile) -ForegroundColor DarkGray"
    psFile.WriteLine "Say 'Done. This window closes in 5 seconds.' 'Green'"
    psFile.WriteLine "try { Stop-Transcript | Out-Null } catch {}"
    psFile.WriteLine "Start-Sleep -Seconds 5"
    psFile.Close
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    wsh.Run "powershell.exe -ExecutionPolicy Bypass -File """ & base & "\PW_AttrScan.ps1"" -InputFile """ & base & "\PW_Folders.txt"" -OutputFile """ & gScanOut & """", 7, False
    gScanBusy = True: gScanStart = Now
    SchedScanPoll
    Notify "Attribute scan started in the background (PowerShell window on the taskbar). You can keep working; the sheet updates when it finishes."
End Sub

Private Sub SchedScanPoll()
    gScanNext = Now + TimeSerial(0, 0, 3)
    Application.OnTime gScanNext, "'" & ThisWorkbook.Name & "'!PollAttrScan"
End Sub

Public Sub PollAttrScan()
    If Not gScanBusy Then Exit Sub
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(gScanOut & ".done") Then
        gScanBusy = False
        Dim nNew As Long
        nNew = MergeFoundAttributes(gScanOut, True)
        On Error Resume Next
        fso.DeleteFile gScanOut & ".done"
        On Error GoTo 0
        If nNew > 0 Then
            gChainStep = 3
            Notify "Attribute scan finished (" & Format(Now - gScanStart, "hh:mm:ss") & "): " & nNew & " new attribute(s). Starting the exact download in the background..."
            Application.OnTime Now + TimeSerial(0, 0, 2), "'" & ThisWorkbook.Name & "'!DownloadFromProjectWise_PWPS"
        Else
            gChainStep = 0
            Notify "Attribute scan finished (" & Format(Now - gScanStart, "hh:mm:ss") & "): no new attributes."
        End If
        Exit Sub
    End If
    If Now - gScanStart > TimeSerial(0, 0, 20) And Not PsRunning("PW_AttrScan.ps1") Then
        gScanBusy = False: gChainStep = 0
        Application.StatusBar = False
        Notify "Attribute scan aborted (PowerShell window closed or error). Check the log with OPEN PS LOG."
        Exit Sub
    End If
    Dim pt As String
    On Error Resume Next
    If fso.FileExists(gScanOut & ".progress") Then pt = Trim(fso.OpenTextFile(gScanOut & ".progress", 1, False, -1).ReadLine)
    On Error GoTo 0
    If pt = "" Then pt = "Attribute scan running in the background"
    Application.StatusBar = pt & "  (total " & Format(Now - gScanStart, "hh:mm:ss") & ", click OPEN PS LOG for details)"
    SchedScanPoll
End Sub

Private Sub WritePsPrelude(psFile As Object, ByVal job As String)  ' wspolny naglowek skryptow PS: log do pliku (transcript), tryb okna, helpery
    psFile.WriteLine "$jobName = '" & job & "'"
    psFile.WriteLine "$psLogDir = Join-Path (Split-Path -Parent $OutputFile) 'PW_Logs'"
    psFile.WriteLine "if (-not (Test-Path $psLogDir)) { New-Item -ItemType Directory -Path $psLogDir -Force | Out-Null }"
    psFile.WriteLine "$psLogFile = Join-Path $psLogDir ((Get-Date).ToString('yyyy-MM-dd_HHmmss') + '_' + $jobName + '.log')"
    psFile.WriteLine "try { Start-Transcript -Path $psLogFile -Force | Out-Null } catch {}"
    psFile.WriteLine "Set-Content -Path (Join-Path $psLogDir 'last.txt') -Value $psLogFile -Encoding UTF8"
    psFile.WriteLine "try { Add-Type -Namespace Q -Name K -MemberDefinition '[DllImport(""kernel32.dll"")] public static extern IntPtr GetStdHandle(int h); [DllImport(""kernel32.dll"")] public static extern bool GetConsoleMode(IntPtr h, out uint m); [DllImport(""kernel32.dll"")] public static extern bool SetConsoleMode(IntPtr h, uint m);'; $h = [Q.K]::GetStdHandle(-10); $m = 0; if ([Q.K]::GetConsoleMode($h, [ref]$m)) { [void][Q.K]::SetConsoleMode($h, (($m -bor 0x80) -band (-bnot 0x40))) } } catch {}"
    psFile.WriteLine "try { $Host.UI.RawUI.WindowTitle = 'PW Tracker - ' + $jobName; $Host.UI.RawUI.BackgroundColor = 'Black'; $Host.UI.RawUI.ForegroundColor = 'Gray'; Clear-Host } catch {}"
    psFile.WriteLine "function Say([string]$m, [string]$c = 'Gray') { Write-Host ((Get-Date).ToString('HH:mm:ss') + '  ') -NoNewline -ForegroundColor DarkGray; Write-Host $m -ForegroundColor $c }"
    psFile.WriteLine "function Say-Head([string]$t) { Write-Host ''; Write-Host ('  ' + $t) -ForegroundColor Cyan; Write-Host ('  ' + ('-' * 66)) -ForegroundColor DarkCyan }"
    psFile.WriteLine "Write-Host ''; Write-Host '  ================================================================' -ForegroundColor DarkCyan"
    psFile.WriteLine "Write-Host '   PROJECTWISE INTEGRATION' -ForegroundColor White; Write-Host ('   Job: ' + $jobName.ToUpper() + '   |   Started: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) -ForegroundColor Gray"
    psFile.WriteLine "Write-Host ('   Log: ' + $psLogFile) -ForegroundColor DarkGray; Write-Host '   You can keep working in Excel - this window runs in the background.' -ForegroundColor DarkGray"
    psFile.WriteLine "Write-Host '  ================================================================' -ForegroundColor DarkCyan"
End Sub

' Otwiera najnowszy log PowerShella (folder PW_Logs obok skoroszytu) w Notatniku
Public Sub OpenLastPsLog()
    Dim fso As Object, f As Object, best As Object, dir As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    dir = ThisWorkbook.Path & "\PW_Logs"
    If Not fso.FolderExists(dir) Then Notify "No PowerShell logs yet (folder PW_Logs is created by the first update).": Exit Sub
    For Each f In fso.GetFolder(dir).Files
        If LCase(fso.GetExtensionName(f.Name)) = "log" Then
            If best Is Nothing Then
                Set best = f
            ElseIf f.DateLastModified > best.DateLastModified Then
                Set best = f
            End If
        End If
    Next f
    If best Is Nothing Then Notify "No PowerShell logs yet.": Exit Sub
    CreateObject("WScript.Shell").Run "notepad.exe """ & best.Path & """", 1, False
End Sub
