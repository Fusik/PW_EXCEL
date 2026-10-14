Attribute VB_Name = "Module1"
Option Explicit                                     ' Wymusza jawną deklarację wszystkich zmiennych w kodzie, co zapobiega literówkom

' =======================================================================
' WERSJA: V1.2 (Z obsługą tabeli PW_DATA i pełnymi komentarzami)
' GLOBALNE ZMIENNE DLA MODUŁU LOGOWANIA
' =======================================================================
Private wsLog As Worksheet                          ' Zmienna przechowująca obiekt arkusza z logami
Private nextLogRow As Long                          ' Zmienna przechowująca numer kolejnego, pustego wiersza do zapisu w logach
Private logStartRow As Long                         ' Zmienna przechowująca numer wiersza początkowego dla danej sesji uruchomienia

' =======================================================================
' GŁÓWNA PROCEDURA (ORKIESTRATOR)
' Zarządza całym procesem: od przygotowania plików, przez uruchomienie
' skryptu PowerShell, aż po analizę różnicową i zapis logów.
' =======================================================================
Sub DownloadFromProjectWise_PWPS()                  ' Publiczny punkt wejścia: chroni stan Excela i raportuje błędy
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
    On Error Resume Next
    If Not wsLog Is Nothing Then AddLogEntry "ERROR", errMsg, "-", "-", "-", "-", CLR_DELETED
    MsgBox "Macro stopped unexpectedly." & vbCrLf & errMsg, vbCritical
    Resume Done
End Sub

Private Sub RunDownload()                           ' Właściwa logika orkiestratora (dawniej treść DownloadFromProjectWise_PWPS)
    ' --- [NOWE] START STOPERA ---
    Dim startExecTime As Date                       ' Deklaracja zmiennej do przechowywania czasu rozpoczęcia
    startExecTime = Now                             ' Zapisanie aktualnej daty i godziny do zmiennej startExecTime
    
    If ThisWorkbook.Path = "" Then                  ' Sprawdzenie, czy plik Excela został już zapisany na dysku (ścieżka nie jest pusta)
        MsgBox "Please save the Excel workbook to your drive first!", vbCritical ' Wyświetlenie błędu, jeśli plik nie jest zapisany
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
    psArrayString = GetAttributesFromSheet3()       ' Wywołanie funkcji pobierającej atrybuty i przypisanie wyniku do zmiennej
    If psArrayString = "" Then Exit Sub             ' Jeśli funkcja nie zwróciła żadnych atrybutów, przerwij działanie makra
    
    Dim lastColHeaders As Long                      ' Deklaracja zmiennej dla numeru ostatniej kolumny z nagłówkiem
    lastColHeaders = ws1.Cells(11, ws1.Columns.count).End(xlToLeft).Column ' Znalezienie ostatniej niepustej kolumny w 11 wierszu arkusza 1
    If lastColHeaders < 1 Then lastColHeaders = 1   ' Zabezpieczenie przed błędem, gdyby arkusz był całkowicie pusty
    
    ws1.Range(ws1.Cells(11, 1), ws1.Cells(11, lastColHeaders)).ClearContents ' Wyczyszczenie dotychczasowych nagłówków w wierszu 11
    ws1.Cells(11, 1).Value = "File Name"            ' Ustawienie stałego nagłówka kolumny 1
    ws1.Cells(11, 2).Value = "PW Source Path"       ' Ustawienie stałego nagłówka kolumny 2
    ws1.Cells(11, 3).Value = "PW Link"              ' Ustawienie stałego nagłówka kolumny 3
    ws1.Cells(11, 4).Value = "Updater"              ' Ustawienie stałego nagłówka kolumny 4
    ws1.Cells(11, 5).Value = "Updated"              ' Ustawienie stałego nagłówka kolumny 5
    StyleHeader ws1.Range("A11:E11")                ' Jednolity styl nagłówków
    
    CopyHeadersFromSheet3                           ' Wywołanie procedury kopiującej nagłówki atrybutów z arkusza 3
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
    GeneratePowerShellScript psScriptPath, psArrayString ' Wywołanie procedury generującej fizyczny kod skryptu PS1 na dysku
    
    Dim csvOutputPath As String                     ' Deklaracja zmiennej na ścieżkę pliku wyjściowego CSV
    csvOutputPath = basePath & "\PW_Output.csv"     ' Zbudowanie pełnej ścieżki pliku CSV
    DeleteIfExists csvOutputPath                    ' Stare wyniki NIE mogą zostać użyte, gdy PowerShell się nie powiedzie (inaczej delta liczyłaby się na starych danych)
    DeleteIfExists structureLogPath
    RunPowerShellScript psScriptPath, txtFoldersPath, csvOutputPath, structureLogPath ' Uruchomienie skryptu PowerShell w tle
    
    AddLogEntry "INFO", "Started Delta analysis in Excel...", "-", "-", "-", "-", 0 ' Zapisanie w logach informacji o rozpoczęciu analizy Delta
    
    ProcessDeltaAnalysis ws1, csvOutputPath, structureLogPath, numScopes, arrCleanScopes ' Wywołanie głównego silnika analizy różnic (CSV vs Excel)
    
    CleanUpFiles txtFoldersPath, psScriptPath, csvOutputPath ' Wywołanie procedury z pytaniem o usunięcie plików tymczasowych
    
    ' --- ZAPIS CZASU I DATY W KOMÓRCE A8 ---
    Dim endExecTime As Date                         ' Deklaracja zmiennej do przechowywania czasu zakończenia
    Dim durationStr As String                       ' Deklaracja zmiennej do przechowywania czasu trwania w formacie tekstowym
    endExecTime = Now                               ' Zapisanie aktualnej daty i godziny do zmiennej końcowej
    durationStr = Format(endExecTime - startExecTime, "hh:mm:ss") ' Obliczenie i sformatowanie różnicy czasu do zmiennej
    
    ws1.Cells(8, 1).Value = "Last Database Update: " & Format(Now, "yyyy-mm-dd HH:mm:ss") & " | Total Duration: " & durationStr ' Wpisanie podsumowania w A8
    ws1.Cells(8, 1).Font.Bold = True                ' Pogrubienie tekstu w komórce A8
    ws1.Cells(8, 1).Font.Color = RGB(0, 102, 204)   ' Ustawienie niebieskiego koloru czcionki w A8
    
    ' --- [NOWE] TWORZENIE TABELI PW_DATA (NA KOŃCU) ---
    Dim finalLastRow As Long                        ' Deklaracja zmiennej dla ostatniego wiersza wygenerowanych danych
    Dim finalLastCol As Long                        ' Deklaracja zmiennej dla ostatniej kolumny wygenerowanych danych
    Dim tableRange As Range                         ' Deklaracja obiektu zakresu pod nową tabelę
    
    finalLastRow = ws1.Cells(ws1.rows.count, 1).End(xlUp).Row ' Znalezienie ostatecznego ostatniego wiersza po wszystkich operacjach makra
    finalLastCol = ws1.Cells(11, ws1.Columns.count).End(xlToLeft).Column ' Znalezienie ostatecznej ostatniej użytej kolumny z nagłówkiem
    
    If finalLastRow >= 11 Then                      ' Upewnienie się, że mamy przynajmniej nagłówki (nie tworzy tabeli na zupełnie pustym arkuszu)
        Set tableRange = ws1.Range(ws1.Cells(11, 1), ws1.Cells(finalLastRow, finalLastCol)) ' Określenie pełnych granic danych pod nową tabelę
        Set lo = ws1.ListObjects.Add(xlSrcRange, tableRange, , xlYes) ' Utworzenie dynamicznej tabeli ListObject w zaznaczonym obszarze uwzględniając nagłówki
        lo.Name = "PW_DATA"                         ' Nadanie nowej tabeli nazwy PW_DATA do użycia przy kolejnym odpaleniu makra
        lo.TableStyle = "TableStyleLight1"          ' Ustawienie jasnego stylu dla tabeli (najczystszy biały/szary wbudowany profil)
        lo.ShowTableStyleRowStripes = False         ' Usunięcie poziomych pasków (zebry) dla idealnie białego i czystego wyglądu tabeli zgodnego z poleceniem
    End If                                          ' Zakończenie bloku tworzenia tabeli końcowej
    ' --------------------------------------------------------------
    
    AddLogEntry "INFO", "OPERATION COMPLETED", "-", "-", "-", "-", 0 ' Zapisanie w logach informacji o pomyślnym zakończeniu makra
    wsLog.Columns("A:G").AutoFit                    ' Automatyczne dopasowanie szerokości kolumn w arkuszu logów
    ws1.Activate                                    ' Aktywowanie arkusza 1 na sam koniec działania makra
    MsgBox "Update completed! Data merged. Tabela PW_DATA została odświeżona.", vbInformation ' Wyświetlenie okna z komunikatem sukcesu
End Sub                                             ' Koniec głównej procedury

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
        MsgBox "Sheet '" & SH_ATTRS & "' not found!", vbCritical ' Wyświetlenie błędu o braku arkusza
        GetAttributesFromSheet3 = ""                ' Zwrócenie pustego ciągu znaków jako wynik funkcji
        Exit Function                               ' Przerwanie działania funkcji
    End If                                          ' Koniec warunku
    
    lastRow = ws3.Cells(ws3.rows.count, "A").End(xlUp).Row ' Znalezienie ostatniego wiersza z danymi w kolumnie A
    If lastRow < ATTRS_FIRST_ROW Then                ' Jeśli brak atrybutów poniżej nagłówka (tylko nagłówek) lub mniej
        MsgBox "Column A in sheet '" & SH_ATTRS & "' is empty. Add attribute names there!", vbExclamation ' Wyświetl ostrzeżenie o braku atrybutów
        GetAttributesFromSheet3 = ""                ' Zwróć pusty ciąg znaków
        Exit Function                               ' Przerwij funkcję
    End If                                          ' Koniec warunku
    
    psArray = "@("                                  ' Inicjalizacja początku definicji tablicy w składni PowerShell
    Dim count As Long                               ' Deklaracja licznika atrybutów
    count = 0                                       ' Wyzerowanie licznika atrybutów
    
    For i = ATTRS_FIRST_ROW To lastRow               ' Pętla po wierszach z atrybutami
        attrName = Trim(ws3.Cells(i, 1).Value)      ' Pobranie nazwy atrybutu z komórki i usunięcie zbędnych spacji
        If attrName <> "" Then                      ' Sprawdzenie czy komórka nie była pusta
            psArray = psArray & "'" & attrName & "'," ' Doklejenie atrybutu w apostrofach i przecinka do stringu
            count = count + 1                       ' Zwiększenie licznika dodanych atrybutów
        End If                                      ' Koniec warunku
    Next i                                          ' Koniec pętli i przejście do kolejnego wiersza
    
    If count > 0 Then                               ' Jeśli dodano przynajmniej jeden atrybut
        psArray = Left(psArray, Len(psArray) - 1) & ")" ' Ucięcie ostatniego przecinka i dodanie zamykającego nawiasu tablicy
        GetAttributesFromSheet3 = psArray           ' Przypisanie gotowego tekstu jako wynik funkcji
    Else                                            ' W przeciwnym razie (brak poprawnych atrybutów)
        MsgBox "No valid attributes found in sheet '" & SH_ATTRS & "'!", vbExclamation ' Wyświetl błąd
        GetAttributesFromSheet3 = ""                ' Zwróć pusty ciąg
    End If                                          ' Koniec warunku
End Function                                        ' Koniec funkcji

' =======================================================================
' 2. LOGOWANIE W EXCELU (Przygotowanie Arkusza "Log")
' =======================================================================
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
        MsgBox "No active PW links in sheet '" & SH_FOLDERS & "' (column B).", vbExclamation ' Ostrzeżenie dla użytkownika
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
        If UCase(Trim(CStr(ws2.Cells(i, FOLDERS_ACTIVE_COL).Value))) = "NIE" Then rawLink = "" ' Wiersz wyłączony przez użytkownika (Aktywny = NIE)       ' Pobranie i oczyszczenie linku ze spacji wiodących
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
Private Sub GeneratePowerShellScript(psPath As String, psArrayString As String) ' Generator fizycznego kodu skryptu PS1
    Dim fso As Object, psFile As Object             ' Zmienne systemowe na obiekt pliku i generator tekstu
    Set fso = CreateObject("Scripting.FileSystemObject") ' Uruchomienie narzędzia systemowego
    Set psFile = fso.CreateTextFile(psPath, True, True) ' Utworzenie/Nadpisanie pliku PS1, w trybie zapisu (Unicode/UTF8)
    
    psFile.WriteLine "param([string]$InputFile, [string]$OutputFile, [string]$LogFile)" ' Zapis kodu PS: Parametry przyjmujące ścieżki plików z makra
    psFile.WriteLine "$ErrorActionPreference = 'Stop'" ' Zapis kodu PS: Konfiguracja, zatrzymaj skrypt jeśli wystąpi krytyczny błąd PowerShell
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz dla czytelności kodu w PS1
    psFile.WriteLine "$globalWatch = [System.Diagnostics.Stopwatch]::StartNew()" ' Zapis kodu PS: Uruchomienie głównego wbudowanego stopera PS
    psFile.WriteLine "Write-Host 'Loading PWPS_DAB...' -ForegroundColor Cyan" ' Zapis kodu PS: Wypisanie informacji w konsoli
    psFile.WriteLine "Import-Module PWPS_DAB -Force" ' Zapis kodu PS: Wymuszone wczytanie biblioteki integracji z ProjectWise
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz
    psFile.WriteLine "$inputData = @(Get-Content $InputFile | Where-Object { $_.Trim() -ne '' })" ' Zapis kodu PS: Wczytanie zawartości pliku TXT (wiersze konfiguracji)
    psFile.WriteLine "$totalFolders = $inputData.Count" ' Zapis kodu PS: Obliczenie łącznej liczby zdefiniowanych folderów
    psFile.WriteLine "$folderIndex = 0"             ' Zapis kodu PS: Reset wskaźnika przetwarzanego folderu
    psFile.WriteLine "$allVersions = " & IIf(GET_ALL_VERSIONS, "$true", "$false") ' Czy pobierac wszystkie rewizje plikow (modConfig)
    psFile.WriteLine "$okScopes = New-Object System.Collections.ArrayList" ' Lista folderów, które faktycznie udało się przeskanować (używana później do bezpiecznego oznaczania usunięć)
    psFile.WriteLine "$selectedAttrs = " & psArrayString ' Zapis kodu PS: Wstrzyknięcie z Excela tablicy z atrybutami (np. @('Attr1','Attr2'))
    psFile.WriteLine "function Get-Attr($d, $n) { $p = $d.psobject.Properties[$n]; if ($p) { return $p.Value }; foreach ($q in $d.psobject.Properties) { $v = $q.Value; if ($null -eq $v -or $v -is [string] -or $v -is [ValueType] -or $v -is [array]) { continue }; if ($v -is [System.Collections.IDictionary]) { if ($v.Contains($n)) { return $v[$n] } } else { $pp = $v.psobject.Properties[$n]; if ($pp) { return $pp.Value } } }; return $null }" ' Pomocnik: atrybut moze byc wlasciwoscia dokumentu albo siedziec w zagniezdzonym obiekcie/slowniku
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
    psFile.WriteLine " New-PWLogin -DatasourceName $dsName -BentleyIMS -NonAdminLogin -Verbose | Out-Null" ' Zapis kodu PS: Procedura logowania do serwera Benteley
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
    
    psFile.WriteLine " $extra = @{}; if ($allVersions -and (Get-Command Get-PWDocumentsBySearch).Parameters.ContainsKey('GetVersionsToo')) { $extra['GetVersionsToo'] = $true }" ' Z GetVersionsToo wyszukiwanie zwraca tez starsze rewizje (A, B, C...)
    psFile.WriteLine " $documents = @(Get-PWDocumentsBySearch -FolderID $pwFolder.ProjectID -GetAttributes @extra -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)"
    psFile.WriteLine " $maxSeq = @{}; foreach ($d in $documents) { $u = [string]$d.DocumentURN; $sq = [int]$d.VersionSequence; if (-not $maxSeq.ContainsKey($u) -or $sq -gt $maxSeq[$u]) { $maxSeq[$u] = $sq } }" ' Najwyzsza wersja kazdego URN
    psFile.WriteLine " Write-Host ('Found total documents: ' + $documents.Count) -ForegroundColor Green" ' Zapis kodu PS: Wypisanie ilości odnalezionych rekordów w terminalu
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
    
    psFile.WriteLine " foreach ($attr in $selectedAttrs) {" ' Zapis kodu PS: Przechodzenie przez każdą nazwę atrybutu wpisaną w Excelu
    psFile.WriteLine " if ($attr -match '^(Name|Description|FileName|Updater|Updated)$') { continue }" ' Zapis kodu PS: Zablokowanie nadpisywania stałych systemowych, jeśli ktoś wpisał bzdury
    psFile.WriteLine " $val = Get-Attr $doc $attr"
    psFile.WriteLine " if ($val -is [DateTime]) { $val = $val.ToString('yyyy-MM-dd') }" ' Zapis kodu PS: Normalizacja, jeśli atrybut PW był typu Data
    psFile.WriteLine " elseif ($val -is [array]) { $val = ($val -join ', ') }" ' Zapis kodu PS: Złączenie atrybutów wielokrotnych (np. wiele tagów) przecinkami w jeden wiersz
    psFile.WriteLine " if ($val -is [string]) { $val = $val.Replace('|', '-').Replace(""`n"", ' ').Replace(""`r"", '').Trim() }" ' Zapis kodu PS: Zabezpieczenie CSV przed zepsuciem (znak potoku to separator, więc jest usuwany z tekstu użytkownika PW)
    psFile.WriteLine " $obj[$attr] = $val"          ' Zapis kodu PS: Wstawienie ostatecznie sformatowanego atrybutu do kolumny
    psFile.WriteLine " }"                           ' Zapis kodu PS: Koniec pętli po właściwościach pojedynczego pliku
    psFile.WriteLine " [PSCustomObject]$obj"        ' Zapis kodu PS: Rzucenie wygenerowanego obiektu pliku na stertę (doda się automatycznie do tablicy wyników pętli na górze)
    psFile.WriteLine " }"                           ' Zapis kodu PS: Zakończenie pętli przechodzącej po dokumentach
    psFile.WriteLine " } else { Write-Host 'FOLDER NOT FOUND OR INVALID!' -ForegroundColor Red }" ' Zapis kodu PS: Błąd w terminalu w przypadku problemów z folderem
    psFile.WriteLine " } catch { Write-Host ('ERROR: ' + $_.Exception.Message) -ForegroundColor Red }" ' Zapis kodu PS: Czerwony błąd systemowy bazy dla pojedynczego wariantu
    psFile.WriteLine " Undo-PWLogin | Out-Null"    ' Zapis kodu PS: Odłączenie od serwera Bentley dla danej iteracji folderowej, aby zachować poprawne sesje licencyjne
    psFile.WriteLine " $folderWatch.Stop()"         ' Zapis kodu PS: Zatrzymanie stopera lokalnego zadania
    psFile.WriteLine " Write-Host ('[TIMER] Folder processed in: ' + $folderWatch.Elapsed.ToString('hh\:mm\:ss')) -ForegroundColor Magenta" ' Zapis kodu PS: Wyświetlenie czasu wykonania operacji na jednej gałęzi folderowej
    psFile.WriteLine "})"                           ' Zapis kodu PS: Koniec wielkiej pętli nad folderami docelowymi
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz oddzielający kod pobierania i zapisu pliku
    psFile.WriteLine "Write-Progress -Id 1 -Activity 'Completed' -Completed" ' Zapis kodu PS: Usunięcie paska pobierania z terminala po zakończeniu
    psFile.WriteLine "Write-Host ('TOTAL DOCUMENTS: ' + $results.Count) -ForegroundColor Cyan" ' Zapis kodu PS: Wypisanie ile całościowo zrzucono rekordów z pamięci serwera do RAM
    
    psFile.WriteLine "if ($results.Count -gt 0) {"  ' Zapis kodu PS: Warunek jeśli cokolwiek ściągnięto do tablicy (wyniki są obecne)
    psFile.WriteLine " $results = $results | Sort-Object -Property 'PW Link' -Unique | Sort-Object 'PW Source Path', 'File Name'" ' Zapis kodu PS: Eliminacja duplikatów ze skanowania nakładających się folderów z użyciem Unikalnych URN oraz posortowanie alfabetycznie
    psFile.WriteLine " $results | Export-Csv -Path $OutputFile -NoTypeInformation -Delimiter '|' -Encoding UTF8" ' Zapis kodu PS: Zrzut danych CSV, z kodowaniem pozwalającym na polskie i chińskie znaki, odseparowany znakami potoku (|)
    psFile.WriteLine ""                             ' Zapis kodu PS: Pusty wiersz
    
    psFile.WriteLine " Write-Host 'Generating Structure Log...' -ForegroundColor Cyan" ' Zapis kodu PS: Informacja w konsoli o rozpoczęciu generowania mapy folderów (Txt)
    psFile.WriteLine " $logContent = @()"           ' Zapis kodu PS: Deklaracja nowej tablicy na wpisy logu TXT
    psFile.WriteLine " $logContent += '======================================================================'" ' Zapis kodu PS: Dodanie nagłówka separatora logu
    psFile.WriteLine " $logContent += 'PROJECTWISE STRUCTURE LOG'" ' Zapis kodu PS: Dodanie tytułu
    psFile.WriteLine " $logContent += 'Generated: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')" ' Zapis kodu PS: Wpisanie sygnatury czasowej wygenerowania
    psFile.WriteLine " $logContent += '======================================================================'" ' Zapis kodu PS: Dodanie stopki nagłówka
    psFile.WriteLine " foreach ($ok in $okScopes) { $logContent += ('[SCOPE_OK] ' + $ok) }" ' Zapis kodu PS: Zapis listy poprawnie przeskanowanych folderów na początku logu
    psFile.WriteLine " $grouped = $results | Group-Object 'PW Source Path' | Sort-Object Name" ' Zapis kodu PS: Pogrupowanie bazy w pamięci operacyjnej po identycznych ścieżkach do renderu graficznego
    psFile.WriteLine " foreach ($group in $grouped) {" ' Zapis kodu PS: Iterowanie dla każdej grupy (folderu)
    psFile.WriteLine " $logContent += ''"           ' Zapis kodu PS: Dodanie pustej linii jako bufor przed nową grupą
    psFile.WriteLine " $logContent += ('[FOLDER] ' + $group.Name)" ' Zapis kodu PS: Wypisanie w logu tekstowym nazwy rodzica/folderu
    psFile.WriteLine " foreach ($item in $group.Group) {" ' Zapis kodu PS: Przelatuj przez obiekty wewnątrz grupy
    psFile.WriteLine " $logContent += (' |-- [FILE] ' + $item.'File Name')" ' Zapis kodu PS: Formatowanie typu drzewo rysujące gałąź do pliku
    psFile.WriteLine " $logContent += ('     [LINK] ' + $item.'PW Link')" ' Zapis kodu PS: Umiejscowienie pod tym ukrytego unikalnego linku URN do systemu identyfikacji śmieci w Excelu
    psFile.WriteLine " }"                           ' Zapis kodu PS: Koniec wariantu grupowego
    psFile.WriteLine " }"                           ' Zapis kodu PS: Zakończenie przebudowy logu tekstowego
    psFile.WriteLine " $logContent | Out-File -FilePath $LogFile -Encoding UTF8 -Force" ' Zapis kodu PS: Fizyczne wyeksportowanie logu do pliku i nadpisanie w tle starych logów
    psFile.WriteLine "}"                            ' Zapis kodu PS: Zakończenie całego bloku if (obsługa wyników)
    
    psFile.WriteLine "$globalWatch.Stop()"          ' Zapis kodu PS: Zatrzymanie głównego timera terminalowego
    psFile.WriteLine "Write-Host '================================' -ForegroundColor DarkCyan" ' Zapis kodu PS: Dodanie ostatniego wizualnego oddzielenia okna
    psFile.WriteLine "Write-Host ('[TIMER] TOTAL EXECUTION TIME: ' + $globalWatch.Elapsed.ToString('hh\:mm\:ss')) -ForegroundColor Magenta" ' Zapis kodu PS: Informacja w oknie CMD o długości trwania skanowania bazy
    psFile.WriteLine "Write-Host 'Done! Window will close automatically in 10 seconds...' -ForegroundColor Green" ' Zapis kodu PS: Informacja o zautomatyzowanym zamknięciu, okno po prostu powiadomiło że zaraz "gaśnie"
    psFile.WriteLine "Start-Sleep -Seconds 10"      ' Zapis kodu PS: Opóźnienie zamknięcia procesu powłoki, by dało się przeczytać czas
    psFile.Close                                    ' VBA: Ostateczne zapisanie wygenerowanego kodu w pliku .ps1 i zwolnienie uchwytu
End Sub                                             ' Zakończenie modułu kodującego plik

Private Sub RunPowerShellScript(psPath As String, txtPath As String, csvPath As String, logPath As String) ' Procedura do wykonania wygenerowanego wcześniej pliku na silniku Windows PowerShell
    Dim wsh As Object                               ' Inicjalizacja referencji systemowej
    Set wsh = CreateObject("WScript.Shell")         ' Zainicjowanie interfejsu systemu (WScript do powłoki)
    Dim psCommand As String                         ' Zmienna przetrzymująca wywołanie polecenia dla konsoli
    psCommand = "powershell.exe -ExecutionPolicy Bypass -File """ & psPath & """ -InputFile """ & txtPath & """ -OutputFile """ & csvPath & """ -LogFile """ & logPath & """" ' Stworzenie instrukcji wyłączającej zasady bezpieczeństwa lokalnego skryptu oraz podającej mu 3 pliki jako parametry wejściowe/wyjściowe
    wsh.Run psCommand, 1, True                      ' Odpalenie zapytania w widocznym oknie (1) wymuszając oczekiwanie Excela na koniec jego działania (True)
End Sub                                             ' Zakończenie wywołania środowiska OS

' =======================================================================
' 5. DELTA ANALYSIS: ANALIZUJE CSV ORAZ PARSUJE PLIK TXT
' =======================================================================
Private Sub ProcessDeltaAnalysis(ws1 As Worksheet, csvPath As String, txtLogPath As String, numScopes As Long, ByVal arrCleanScopes As Variant) ' Algorytm łączący bazę wyciągniętą i przetworzoną (CSV) z tym, co aktualnie znajduje się w pliku Excela, ignorując pliki które zostały skasowane poza Excelem w chmurze
    Dim fso As Object                               ' Referencja FSO
    Set fso = CreateObject("Scripting.FileSystemObject") ' Uruchomienie obsługi plików
    
    If Not fso.FileExists(csvPath) Then             ' Upewnienie się, że PowerShell pomyślnie utworzył bazę wynikową
        AddLogEntry "ERROR", "Missing results CSV file.", "-", "-", "-", "-", CLR_DELETED ' Jeśli nie, poinformowanie w logach na czerwono
        Exit Sub                                    ' Wyrzucenie błędu w formie cichej (odłączenie procesu), co nie wyczyści reszty danych
    End If                                          ' Zakończenie weryfikacji powiązania plików bazy
    
    On Error Resume Next                            ' Zabezpieczenie przed usuwaniem filtrów i wierszy
    If ws1.FilterMode Then ws1.ShowAllData          ' Wyłączenie filtra w Excelu, aby macro mogło modyfikować również wyfiltrowane obszary (rozwiązywanie bugów wydajności)
    ws1.Cells.ClearOutline                          ' Zdjęcie jakiegokolwiek wcześniej zapisanego poziomu grupowania i zwijania (uszkodzenia wizualne)
    ws1.rows.Hidden = False                         ' Odkrycie całego arkusza, aby makro nie pomyliło numerów R
    On Error GoTo 0                                 ' Utrzymanie normalnego profilowania awarii
    
    RemoveEmptyRows ws1                             ' Odpalenie procedury czyszczenia wierszy (która ma zapytanie MsgBox z reguł usera)
    
    Dim dictOkScopes As Object                      ' Foldery faktycznie przeskanowane przez PowerShell (bez nich nie wolno oznaczać plików jako usunięte)
    Set dictOkScopes = CreateObject("Scripting.Dictionary")
    dictOkScopes.CompareMode = 1
    Dim dictTxtLinks As Object                      ' Obiekt szybkiego wyszukiwania słownikowego (hashowania)
    Set dictTxtLinks = CreateObject("Scripting.Dictionary") ' Inicjalizacja obiektu Dictionary
    dictTxtLinks.CompareMode = 1                    ' Ustawienie porównywania w trybie Text (case-insensitive, nie rozróżniając wielkości liter)
    
    If fso.FileExists(txtLogPath) Then              ' Upewnienie się, że plik struktury logu .txt został utworzony
        Dim streamTxt As Object                     ' Deklaracja interfejsu połączenia danych
        Set streamTxt = CreateObject("ADODB.Stream") ' Wykorzystanie ADODB z powodu błędów VBA z odczytem innych kodowań z UTF8
        streamTxt.Charset = "utf-8"                 ' Ustawienie poprawnego trybu wymiany znakowej
        streamTxt.Open                              ' Odblokowanie i otwarcie kanału z plikiem TXT
        streamTxt.LoadFromFile txtLogPath           ' Przerzucenie logu tekstowego w całości z pliku z dysku prosto do obwodu
        
        If Not streamTxt.EOS Then                   ' Sprawdzenie czy strumień (EOF - End Of Stream/File) jest pełny (plik nie był pusty)
            Dim txtData As String                   ' Przydzielenie lokalizacji w pamięci tekstowej na zbiór
            txtData = streamTxt.ReadText            ' Przypisanie masywnych paczek tekstu
            txtData = Replace(txtData, vbCrLf, vbLf) ' Oczyszczanie niechcianych znaków końca linii Mac / Win / Unix, zamiana wszystkiego na Linuxa do pętli Split
            txtData = Replace(txtData, vbCr, vbLf)  ' Zamiana dodatkowych znaków karetki (Carriage Return) jeśli zostały
            
            Dim txtRows() As String                 ' Utworzenie tablicy z rozdzielonego strumienia (na konkretne linie pliku)
            txtRows = Split(txtData, vbLf)          ' Rozłamanie tekstu poprzez znak końca Lf
            Dim tr As Long                          ' Deklaracja iteratora dla logów drzewa struktury
            
            For tr = 0 To UBound(txtRows)           ' Początek przelotu po całym pliku Txt
                Dim tLine As String                 ' Bufor do obsługi bieżącego wiersza
                tLine = Trim(txtRows(tr))           ' Ściągnięcie wiersza bez spacji bocznych z Array'a do RAM'u
                If InStr(1, tLine, "[SCOPE_OK]", vbTextCompare) = 1 Then
                    dictOkScopes(UCase(Trim(Mid(tLine, Len("[SCOPE_OK]") + 1)))) = True
                ElseIf InStr(1, tLine, "[LINK]", vbTextCompare) = 1 Then ' Wykrycie tagu URN dla identyfikacji śmieci na podstawie poprzedniego generowania w PowerShell (gdzie dodawano prefiks linku w konsoli pod $item)
                    Dim extLink As String           ' Deklaracja zmiennej która utrzyma unikalny URL do systemu Bentley
                    extLink = Trim(Replace(tLine, "[LINK]", "", 1, 1, vbTextCompare)) ' Pocięcie wiersza poprzez zdjęcie tagu by zostało samo połączenie w stringu
                    If extLink <> "" Then dictTxtLinks(extLink) = True ' Zbudowanie szybkiej ścieżki i powiązania (Hash), weryfikacja ułatwia sprawdzenie (istnieje True/False w < 1 ms!)
                End If                              ' Zamknięcie analizy i parowania Txt do Słownika
            Next tr                                 ' Skok po tablicach i następny wiersz logów
        End If                                      ' Zakończenie bufora pustego
        streamTxt.Close                             ' Bezpieczne zniszczenie rury plikowej
    End If                                          ' Kończenie procedur bezpieczeństwa sprawdzania struktury bazy .txt
    
    Dim stream As Object                            ' Analogiczna procedura jak przy Txt dla obsługi pliku z wynikami CSV. Deklaracja interfejsu (ponieważ CSV również nie ma 100% wsparcia dla "open" z znakami PL - polskimi)
    Set stream = CreateObject("ADODB.Stream")       ' Wwołanie procedury COM streamu w Windows (ADO Database format)
    stream.Charset = "utf-8"                        ' Ponownie wstawienie dekodera znakowego dla pliku (tym razem bazowego CSV z parametrami dokumentów i wersjami dat)
    stream.Open                                     ' Otwarcie linii strumania
    stream.LoadFromFile csvPath                     ' Ładowanie wszystkich wyników (może być setki Megabajtów pliku!)
    
    If stream.EOS Then                              ' Mechanizm zapobiegania awarii - weryfikacja, że PowerShell faktycznie dorzucił dane po "Export-CSV" w PS
        stream.Close                                ' Zamknięcie pliku
        Exit Sub                                    ' Powrót ze struktury - zakończenie ciche, plik bazy był nieprawidłowy, ale nie jest błędem do "zawieszenia Excela"
    End If                                          ' Zakończenie zapobiegania awarii
    
    Dim csvData As String                           ' Rezerwowanie pojemnego RAMu String pod dane
    csvData = stream.ReadText                       ' Odczyt i przetransferowanie 100% bazy w RAM stringu Excela
    stream.Close                                    ' Natychmiastowe zniszczenie połączenia streamu CSV aby inny proces, lub wirus plikowy go nie chwycił w zawieszeniu (System OS Lock)
    csvData = Replace(csvData, vbCrLf, vbLf)        ' Tak samo standardyzacja znaków do tablicy. Część z tego musi być zmieniana na w locie, aby kodowanie Unix (Powershell wewnętrznie) poprawnie dzieliło bazy po wierszach
    csvData = Replace(csvData, vbCr, vbLf)          ' Korekcja łamaczy linii w starej architekturze Macintosh
    
    Dim rows() As String                            ' Array zawierający całe strukturalne rzędy wierszy do sprawdzenia i przetworzenia na Excelu
    rows = Split(csvData, vbLf)                     ' Ostateczne posiekanie tekstu plikowego CSV na wiersz za wierszem Array dla VBA
    If UBound(rows) < 0 Then Exit Sub               ' Mechanizm chroniący przed "pętlą w nieskończoność" gdy plik jest zniszczony, puste bloki nie posiadają zadeklarowanej góry
    
    Dim csvHeaders() As String                      ' Inicjalizacja bloku na listę rzędu numer zero z pliku (Nagłówki Pliku Excela/CSV z właściwości atrybutów)
    csvHeaders = Split(Replace(rows(0), """", ""), "|") ' Zdjęcie obustronnych cudzysłowów które dodał automatycznie po chamsku PowerShell przy formatowaniu CSV, podzielenie znaku potoku
    
    Dim colDict As Object                           ' Deklaracja Hash Dictionary dla łatwej identyfikacji nagłówków Atrybutów do ich położeń w arkuszach Excel (Kolumnie 1,2,3... etc).
    Set colDict = CreateObject("Scripting.Dictionary") ' Uruchomienie Hasha atrybutów
    colDict.CompareMode = 1                         ' Ustawiony format nierozróżniający (A / a nie powoduje zawieszenia i tworzenia dwóch wpisów do nagłówków kolumn).
    
    Dim lastCol As Long, c As Long                  ' Deklaracja "ostatniej skanowanej kolumny" i iteratora C
    lastCol = ws1.Cells(11, ws1.Columns.count).End(xlToLeft).Column ' Skan w Excelu po 11 linii od prawej części arkusza. Znajduje nagłówki by oszacować pozycję dopisywania (Append).
    If lastCol < 3 Then lastCol = 3                 ' Twardy mechanizm chroniący szkielet arkusza 1. Nawet jeśli makro jest puste to trzymaj szerokość na poziomie 3 bazowych kolumn (File Name, Path, Link).
    
    For c = 1 To lastCol                            ' Pętla ładująca wszystkie już widoczne w excelu nazwy kolumn z Arkusza 1, nagłówek "11" do Hash table (słownika).
        If ws1.Cells(11, c).Value <> "" Then colDict(ws1.Cells(11, c).Value) = c ' Zmapuj to, gdzie jest jaki element bazy, by wiedzieć gdzie go zaktualizować (Dla atrybutu "Opis" powiedz że znajduje w kolumnie "C")
    Next c                                          ' Pętla następnego wiersza
    
    Dim csvColMapping() As Long                     ' Array zadeklarowany, by wskazywać które elementy kolumny wejściowej pliku CSV celują w nagłówki Arkusza1 z Excela (mapowanie).
    ReDim csvColMapping(UBound(csvHeaders))         ' Dopasowanie tablic do struktury nagłówkowej
    Dim urnIndex As Long, pathIndex As Long, updaterIndex As Long, folderLinkIndex As Long ' Przetrzymanie specjalnych indeksów atrybutów dla przyspieszenia skryptu (by nie musiał ich wyszukiwać w pętli 50 000 razy)
    updaterIndex = -1: folderLinkIndex = -1         ' Stan początkowy niezdefiniowany (na wypadek ich braku) - oznacza brak wsparcia i wstrzymuje skrypt przed awarią "Index -1 is OUT OF BOUNDS"
    
    For c = 0 To UBound(csvHeaders)                 ' Przeleć całą kolumnę "zero" CSV (czyli nazwy cech/atrybutów)
        Dim hName As String                         ' Cache nazwy bieżącej główki (by ułatwić użycie w procedurach wyszukiwarki słownikowej)
        hName = csvHeaders(c)                       ' Odczytaj zawartość z array wiersza kolumny
        If hName = "PW Link" Then urnIndex = c      ' Wyszukaj flagowy techniczny link na pozycję URN
        If hName = "PW Source Path" Then pathIndex = c ' Wyszukaj systemowy atrybut ścieżki i zapamiętaj która to komórka kolumny!
        If hName = "Updater" Then updaterIndex = c  ' Podłącz właściwość autora do bazy ścieżek
        If hName = "Folder Link" Then folderLinkIndex = c ' Podłącz lokalizator URN na obiekt katalogu z Bentley!
        
        If hName <> "Folder Link" Then              ' Ważne, pomiń wykreowanie widocznej kolumny (folder link ma być wpisany jako sam hyperlink bezpośrednio na tekst pod spodem!) Omiń jego mapowanie po stronie użytkowej widoków arkusza Excel
            If colDict.exists(hName) Then           ' Jeśli ta cecha z CSV już znajduje sie w Hash zdefiniowanym u góry (krok Excela)...
                csvColMapping(c) = colDict(hName)   ' Ustal, że ta kolumna na 100% trafia na pozycję Excel (jest odwzorowana i znana bazie, bezpieczny append nadpisania z aktualizacją wersji)
            Else                                    ' A co jeśli nazwa Atrybutu wpisana przez wklejony PowerShell nie istnieje na szczycie tabel Arkusza 1?
                lastCol = lastCol + 1               ' Powieksz ostatnie użycie struktury w Excelu (nowa docelowa Kolumna X)
                ws1.Cells(11, lastCol).Value = hName ' Automatycznie dogeneruj w głowie/nagłówku arkusza nową etykietę i nazwij tym co dostarczył PowerShell w pliku!
                StyleHeader ws1.Cells(11, lastCol)          ' Jednolity styl nagłówka ' Automatycznie pogrub i wystylizuj tak, aby zrównało się z resztą formatu
                colDict(hName) = lastCol            ' Dołącz świeżo utworzone powiązanie do struktury szybkiej by skrypt go nie wygenerował na następnej pętli!
                csvColMapping(c) = lastCol          ' Wróć na standardowe tory powiązania wektorowego do nadpisywań i wpisów
            End If                                  ' Koniec mechanizmu samouzdrawiającego mapowania kolumn
        Else                                        ' Drugi warunek Jeśli w bazie wykazanej to jest folder Link
            csvColMapping(c) = 0                    ' Ignoruj kompletnie w widocznym Excel (nie dodawaj wpisów numerycznych 1,2,3). Służy tylko jako link "open" z Hyperlink obiektu w VBA!
        End If                                      ' Zakończenie sprawdzania nagłówków!
    Next c                                          ' Wymuś sprawdzenie wszystkich parametrów właściwości z nagłówka CSV dla kolumny ZERO.
    
    Dim urnCol As Long, pathCol As Long             ' Ponowne wydzielenie zmiennych na pamiętanie Excelowych ułożeń atrybutów "PW Link" i "PW Ścieżki" bo w mapowaniu są dla nas świętością lokalizującą w bazie
    urnCol = colDict("PW Link")                     ' Przypisanie miejsca przebywania "PW LINK" w kolumnach 1 Arkusza na konkretną stałą
    pathCol = colDict("PW Source Path")             ' Zapis i zapamiętanie ścieżki w kolumnach
    
    Dim changesCount As Long                        ' Numeracja użyta dla wysłania w eter podsumowań powiadamiających do arkusza "Logs" "Znaleziono XYX zmian" lub "0 - Baza była uaktualniona w pełni w tym pliku"
    changesCount = 0                                ' Reset
    
    Dim dictExcelURNs As Object, dictExcelFolders As Object ' Utworzenie dwóch gigantycznych Hash Object Map, do skanowania arkusza z prędkością mili-sekund zamiast godzinnego pętlowania
    Set dictExcelURNs = CreateObject("Scripting.Dictionary") ' Start Słownika pamięciowego numer 1 pod parametry URN obiektowych identyfikacji systemów wewnątrz bazy (By znaleźć duplikat w 0ms! URN to unikalny identyfikator Bentley z bazy SQL na każdy wrzucony pikokształt pliku z datami w nazwie)
    dictExcelURNs.CompareMode = 1                   ' Wyłączenie funkcji weryfikacji dużych liter w hashowaniu
    Set dictExcelFolders = CreateObject("Scripting.Dictionary") ' Start Słownika drugiego na parametry Grupowania/Nagłówków Zielonych dla widoków folderów
    dictExcelFolders.CompareMode = 1                ' Identycznie wyłączenie rozróżniania Case na literach by nie utworzyło 2 osobnych ścieżek "Docs/Polska" vs "docs/polska" bo system operacyjny tego nie zrobi.
    
    Dim maxR As Long, rEx As Long                   ' Deklaracja szukania w Excelu ostatnich maksymalnych wierszy, oraz interatory dla Rzędów (Excel).
    Dim maxRCell As Range                           ' Interfejs Obiektowy
    Set maxRCell = ws1.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious) ' Algorytm szukający absolutnie na krańcach "Sheet_1"
    If Not maxRCell Is Nothing Then maxR = maxRCell.Row Else maxR = 11 ' Gdy skraj wierszowy został zidentyfikowany to przypisz go, albo start z wiersza 11 gdy kartka biała i zepsuta.
    If maxR < 11 Then maxR = 11                     ' Zabezpieczenie na poziomie pętlowym w przypadku braku odczytania góry
    
    For rEx = 12 To maxR                            ' Główna pętla skanująca to co posiada już wewnątrz tabel stary EXCEL ("Baza"). Zaczynająca od wiersza 12
        Dim ePath As String, eName As String, eURN As String ' Wydzielenie odczytów dla Excela ze starą zawartością tabel - ePath / eName.
        eName = Trim(ws1.Cells(rEx, 1).Value)       ' Pobranie komórki do pamięci bez odstępów
        ePath = Trim(ws1.Cells(rEx, pathCol).Value) ' Pobranie i sformatowanie dla ścieżek skopiowanych u dołu
        eURN = Trim(CStr(ws1.Cells(rEx, urnCol).Value)) ' Ostatnia faza to unikalny obiekt identyfikacji jako Text na pewniaka CStr konwersja!
        
        If ePath = "" And eName <> "" Then          ' Gdy mamy wpis nazwy ale brakuje ścieżki w kolumnie systemowej (PW SOURCE PATH)... Mamy doczynienia z tzw Nagłówkiem Zielonym "FOLDEREM"! Excel jest formatowany w ten sposób by dla oka pokazywał Grupy ścieżek
            dictExcelFolders(UCase(eName)) = rEx    ' Zapamiętaj dla szybkiego indeksowania do mapy (Hash) położenie tego foldera, jako wielkie litery "nazwy i położenia ścieżki (ponieważ on ma ją wypisaną z lewej strony pliku na zielono)". Zapisz na jakim wierszu to występuje do Hash Tab!
        ElseIf ePath <> "" And eName <> "" Then     ' Drugi stan zapytania if. Zwykły prawidłowo wypełniony szkielet - jest Ścieżka pobrana oraz ma nazwę, to po prostu wstawiony do raportu stary PLIK z ProjectWise z CSV wyciągu wstecznego bazy
            If eURN <> "" Then dictExcelURNs(eURN) = rEx ' Tak jak folder. Zapisz identyfikację absolutną pliku "URN", aby za pomocą zera czasu znaleźć duplikat Excel vs CSV wewnątrz. Hash table trzyma tylko 1 pozycję
        End If                                      ' Zakończ obsługę logiki rozpoznającej co w trawie Excelu (arkuszu) starej i wyłączonej (niedostępnej) piszczy!
    Next rEx                                        ' Przerysuj iterację Excela od zera dla kolejnego starego wiersza
    
    Dim r As Long                                   ' Start nowej operacji. Odczytywanie wierszy dla samej wielkiej tabeli CSV pobranej z Powłoki systemowej Powershell na żywca! Właściwy tryb różnic "Delta Process"
    For r = 1 To UBound(rows)                       ' Skok przez każdy pobrany w pliku tekstowym CSV element (Z pominięciem R = 0 na starcie pętli co oznacza że pominiemy kolumnę etykiet - pierwszą nagłówków z pliku)
        If Trim(rows(r)) <> "" Then                 ' Jakakolwiek operacja tylko gdy linia pętli pliku nie stwarza pustego wiersza.
            ProcessSingleCSVRow ws1, rows(r), csvHeaders, csvColMapping, colDict, urnCol, pathCol, urnIndex, pathIndex, updaterIndex, folderLinkIndex, dictExcelURNs, dictExcelFolders, changesCount ' Zrzut 15 argumentów do dużej funkcji logicznej. Będzie rozbierać dany ciąg wyrazowy z jednego wiersza i przetwarzać
        End If                                      ' Zakończenie wariantu zabezpieczającego (If) przed pustką (Error w bazie lub błąd przy wycięciu spacji w plikach CSV przy EOF - ostatnim rzędzie z enterem nadanym przez PowerShell bez sensu)
    Next r                                          ' Koniec pętli dla CSV dla pojedynczej linii (rzędu wyników pobranego prosto z bazy internetowej cloud w sekundach). Uruchom kod nad linijką nowo pobraną dla CSV nr "R".
    
    Dim sc As Long
    For sc = 1 To numScopes                         ' Ostrzeżenie o folderach, których nie udało się przeskanować (login/ścieżka) - ich pliki NIE zostaną oznaczone jako usunięte
        If Not dictOkScopes.exists(arrCleanScopes(sc)) Then AddLogEntry "WARNING", "Scope not scanned (skipped deletion check): " & arrCleanScopes(sc), "-", "-", "-", "-", RGB(255, 192, 0)
    Next sc
    MarkDeletedItems ws1, urnCol, pathCol, lastCol, numScopes, arrCleanScopes, dictTxtLinks, dictOkScopes, changesCount ' Mechanizm sprawdzający Excela od tyłu! Zauważmy że wszystko "DODANO LUB EDYTOWANO (DELTA PROCESS) a pliki co już w bazie Excela, ale zniknęły w CSV (bo z serwera skasowane) nadal w excelu sobie siedzą nieruszane na błękitnym/białym podkładzie! Zaznacz te wiersze jako martwe "DELETED - kolor wściekła czerwień i skreślenie". Mechanizm potrzebuje Map TXT i CSV dla porównań
    
    If changesCount = 0 Then                        ' Całe makro sprawdziło tysiące pikseli dla bazy i CSV, zmian na plikach jest literalnie "zero". Nie musisz logować nic prócz faktu braku jakiejkolwiek zmiany (Dysk twardy nie piszczy że zapisał za dużo zmian a excel oszczędza na limicie 1.000.000 wierszy w pliku XLSX dla raportów i wyciągów ProjectWise)
        AddLogEntry "INFO", "No changes compared to the previous download.", "-", "-", "-", "-", 0 ' Logowanie, wpis bez koloru z samą informacją ("Zero" w kolorach)
    End If                                          ' Końcowa akcja pętli. Jeśli pliki były modyfikowane, wygenerowały log w pętlach i ten wariant ominął kod o 0 wpisach!
    
    If nextLogRow - 1 > logStartRow Then            ' Konwersja formatująca LOG Txt i LOG Excel (aby się uczytelniło wizualnie pod "plusika" i nie zasypywało setką wierszy do scroolowania, log jest pod grupami ukryty pod plusem w Excel Sheet "LOG")
        wsLog.rows((logStartRow + 1) & ":" & (nextLogRow - 1)).Group ' Automatycznie zaznacz wygenerowane zdarzenia z bieżącej daty pracy programu (START) do samego jego zamknięcia przez polecenie LOG.
    End If                                          ' Zakończenie stylizowania logowania Excel. Plik jest czysty po usunięciu starych danych po Delta, sformatowany i zapisany, a po zmianie dodano log
End Sub                                             ' Pełne zamknięcie operacji przetwarzania "Delty" w sercu weryfikacji skryptu w tle

Private Sub RemoveEmptyRows(ws1 As Worksheet)       ' Reguła poprawkowa usera z dnia [2026-03-19] o zapytaniach przed usunięciem wierszy Excela
    Dim maxR As Long, r As Long                     ' Ograniczniki wierszowe
    Dim maxRCell As Range                           ' Definicja zasięgów
    Set maxRCell = ws1.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious) ' Znajdź kres komórek dla Excela pod rzędy 11 i koniec
    If Not maxRCell Is Nothing Then maxR = maxRCell.Row Else maxR = 11 ' Ochrona gdy jest wyczyszczone i na gładko zepsute ("Nothing")
    If maxR < 12 Then Exit Sub                      ' Przerywamy brak możliwości posiadania starych wpisów w wierszu startowym nagłówka tabel nr. 11
    
    Dim emptyCount As Long                          ' Najpierw policz puste wiersze - pytanie ma sens tylko, gdy faktycznie jakieś są
    For r = 12 To maxR
        If Trim(CStr(ws1.Cells(r, 1).Value)) = "" And Trim(CStr(ws1.Cells(r, 3).Value)) = "" Then emptyCount = emptyCount + 1
    Next r
    If emptyCount = 0 Then Exit Sub
    
    ' ZGODNIE Z KOREKTĄ UŻYTKOWNIKA - Makro musi pytać przed usunięciem wierszy
    If MsgBox("Znaleziono " & emptyCount & " pustych wierszy do potencjalnego usunięcia. Czy chcesz kontynuować i je usunąć z arkusza?", vbYesNo + vbQuestion, "Usuwanie wierszy") = vbYes Then ' Okienko potwierdzające operację modyfikacji układu (kasacji rzędów z Arkusz1)
        For r = maxR To 12 Step -1                  ' Przejście w pętli P-U-L-S-U-J-Ą-C-E-J-! Od tyłu pliku, do góry. Przez usuwanie (Delete), dolne wiersze podjeżdżają w górę i omijane są numery R w tradycyjnej pętli. Usuwając od tyłu omijamy zepsute usunięcia.
            If Trim(CStr(ws1.Cells(r, 1).Value)) = "" And Trim(CStr(ws1.Cells(r, 3).Value)) = "" Then ' Detekcja kompletnej pustki. Brak wpisanego atrybutu pliku jak "Opis", nie uratuje go jeśli w nazwie ani unikalnym URN w kolumnie 3 nic mu nie piszczy - co znaczy że wiersz "umarł i jest ułomny" z błędem w klejeniu ręcznym EXCEL
                ws1.rows(r).Delete                  ' Dokonanie faktycznej egzekucji polecenia poprzez obiekt systemowy VBA i skasowanie z powierzchni Excela ułomnego wpisu pustego wiersza. (Arkusz automatycznie dosuwa wiersze w górę z powrotem by zapełnić lukę po rzędzie 'r').
            End If                                  ' Koniec identyfikacji błędów i wierszy zniszczonych przy ręcznej edycji exela!
        Next r                                      ' Koło napędowe pętli tyłowej - przejdź na niższy (czyli wyższy wizualnie) rząd Exela by go przetworzyć przez wariant if Empty Delete = true
    End If                                          ' Zakończenie obróbki wiadomości ostrzegawczej (nie skasuje jeśli zablokujesz okienkiem na ekranie!)
End Sub                                             ' Wyjście zwrotne po zniszczeniu pustaków

Private Sub ProcessSingleCSVRow(ws1 As Worksheet, rawRow As String, csvHeaders() As String, csvColMapping() As Long, colDict As Object, urnCol As Long, pathCol As Long, urnIndex As Long, pathIndex As Long, updaterIndex As Long, folderLinkIndex As Long, ByRef dictExcelURNs As Object, ByRef dictExcelFolders As Object, ByRef changesCount As Long) ' Procedura serce Delty, procesująca tylko jednego pacjenta/wiersza
    Dim cols() As String                            ' Array służący jako podziały na dane wejściowe atrybutowe, 1,2,3... do pociętego wiersza z potokiem "|"
    cols = Split(Replace(rawRow, """", ""), "|")    ' Demontaż znaków zbędnych jak podwójne cytaty (dodawane przy eksporcie stringów CSV na siłę) a na końcu cięcie według "Potokowych Pionowych kresek CSV"
    If UBound(cols) < urnIndex Then Exit Sub        ' Zabawa w piaskownicy! URN jest ostatnim lub dalekim elementem tablicy jeśli podział z Split wywalił się błędem np na 1 słowie (UBOUND(COLS) mniejszy niz INDEX np. z URN), nie wchodź by nie wywalać z Out Of Bounds index error, bo plik CSV nie zawierał wszystkich informacji i makro się potnie błędami
    
    Dim actualFolder As String, urn As String, currentUpdater As String, fName As String, folderUrn As String ' Poszatkowanie zebranych informacji z pliku zasilającego dla makra w odpowiednie miejsce u lokalnych zmiennych
    actualFolder = Trim(cols(pathIndex))            ' Złap adres źródłowy do odniesień dla zielonego okienka
    urn = Trim(cols(urnIndex))                      ' Unikalny token na poszukiwanie Hash w Hash Dict.
    fName = Trim(cols(0))                           ' Element 0 tablicy wejściowej zawsze wskazywał w PS skrypcie wpis $obj['File Name'] który zrzuci się od razu do zmiennej.
    
    If updaterIndex >= 0 And UBound(cols) >= updaterIndex Then ' Weryfikacja braku minus jedynek nadanych po nierozpoznaniu Nagłówków w procedurach głównych mapujących tablice u góry
        currentUpdater = Trim(cols(updaterIndex))   ' Znajdź tego kto zrobił pomyłkę lub zaktualizował atrybut po indeksie wyliczonym dla Excela
        If currentUpdater = "" Then currentUpdater = "Unknown" ' Zamaskowanie braku odpowiedzi o Usera np. przez usunięte konta Windows / Bentley IMS, zamiast pustego wpisu i problemów formatowania
    Else                                            ' Jeśli kolumna UPDATER w pliku po prostu nie istniała dla tej bazy od strony ustawień PS (A zdarza się!)
        currentUpdater = "Unknown"                  ' Oznacz nieokreślonym flagiem bez wywoływania craszy z -1
    End If                                          ' Zamknięcie odgałęzień łapania winowajcy edycji pliku
    
    If folderLinkIndex >= 0 And UBound(cols) >= folderLinkIndex Then ' Taki sam test logiczny -1 na upewnienie się co do istnienia ukrytego pola generowania URL w CSV
        folderUrn = Trim(cols(folderLinkIndex))     ' Zapisanie adresu generowania z pamięci komórki na zmienną
    Else                                            ' Nie ma linku URL... co zrobić w pętli?
        folderUrn = ""                              ' Upewnić że bufor jest pusty i nie wstawia linku do skasowanego folderu w ProjectWise po starej linii CSV (Bezpieczeństwo sesji)
    End If                                          ' Zamknięcie
    
    Dim headerRow As Long                           ' System lokalizowania "głównej ścieżki zielonego paska na EXCELU" z użyciem Słownika (Wielkie Litery) w pamięci Cache RAM komputera dla każdego pojedynczego pliku z wiersza
    headerRow = FindOrCreateFolderHeader(ws1, actualFolder, folderUrn, urnCol, colDict.count, dictExcelFolders) ' Odpalenie procedury odpytującej lub kreującej (i aktualizującej pozycje dla) folderów bazowych by ten plik mógł pod niego wejść niżej do EXCELA z wynikiem numeru do dodania (np Rząd=341!)
    
    Dim foundRow As Long                            ' Numer przypięcia "istniejącego już starego elementu pliku w Excelu" odznaczony jako Zero by wymusić ewentualne generowania nowych elementów od zera
    foundRow = 0                                    ' Wymuszenie twardego resetu stanu istnienia pliku na "Nie istnieje w starej karcie - Oznacz zero - wygeneruj dla niego z CSV nowy Wpis" (Wstawiane gdy stary plik z PW zniknął pod URN lub po starej nazwie - bo URN się zgubił gdy plik był wycięty w nowe miejsce a PW to zapomniał skatalogować w SQL)
    
    If urn <> "" And dictExcelURNs.exists(urn) Then ' Detekcja: Jeśli po adresie identyfikacji unikalnego elementu w Hash Txt jest to plik w starej liście "Excel" - on się nazywa "Istniejący/Aktualizowany Element! Mamy jego numer rzędu z pętli górnej, wystarczy użyć dla Excel"
        foundRow = dictExcelURNs(urn)               ' Odkop numer rzędu z Excela i nadpisz Zmienną foundRow pod ten wykopany z Hasha numer w 1 ms czasu. Np znalazł na Wierszu=912 starego Exela!.
    End If                                          ' Zakończenie wyciągania wariantu
    
    If foundRow = 0 Then                            ' Trudny błąd... Nie ma po URN identyfikacji pliku w systemie a ma folder i może być ukryty po "Samej starych zmianach nazwy w excelu zepsutym ręcznie". Należy sprawdzić ręcznie
        Dim srchRow As Long                         ' Deklaracja pętli awaryjnej (Lokalne Do-While na pod-ścieżce po tym jak folder został wydobyty np z Rząd = 341. Pętla zacznie od 342 szukać starych duplikatów po po prostu "File Name").
        srchRow = headerRow + 1                     ' Wskocz dokładnie 1 niżej od Zielonego foldera głównego znalezionego u góry
        Do While Trim(ws1.Cells(srchRow, 1).Value) <> "" ' Kręć tak długo aż kolumna nazwy pliku nie będzie pusta w arkuszu
            If Trim(ws1.Cells(srchRow, pathCol).Value) = "" Then Exit Do ' Wyskocz z błędu gdy pod nazwą znajduje się nagle pusty path... Oznacza że uderzyliśmy w Kolejny inny Folder Zielony - ten plik w ogóle tutaj fizycznie się nie znajdował! Przerywaj poszukiwanie po "FileName" w dół by nie zmarnować godzin.
            If StrComp(Trim(ws1.Cells(srchRow, 1).Value), fName, vbTextCompare) = 0 And Trim(CStr(ws1.Cells(srchRow, urnCol).Value)) = "" Then ' po nazwie tylko wiersze bez linku (rewizje o tej samej nazwie inaczej nadpisywalyby sie) ' Ale gdy trafiliśmy na starą "Pustą/Niezgodną z URN nazwę i jest tak samo równa (Case insensitive poprzez StrComp)" co wygenerowany fName do pliku... TO JEST TEN ZGUBIONY DUPLIKAT NAZWY I JEST ON PRZETWARZANY
                foundRow = srchRow                  ' Przypisz mu wiersz Excelowy bo to on (Mamy go!). I zamknij operowanie logiki do "0" by nadpisać. Zamiast tworzyć powielany wiersz jako "New Line"
                Exit Do                             ' Ucieczka z pętli po dobrym wyszukaniu - Oszczędność prądu.
            End If                                  ' Zakończenie wariantowania.
            srchRow = srchRow + 1                   ' Posuw pętli While w głąb (przejście o 1 wiersz na dół dla poszukiwań "starej nazwy - zguby").
        Loop                                        ' Koniec okrążenia - wróć i pytaj warunki
    End If                                          ' Zakończenie bloku poszukiwania awaryjnego z zepsutym/pominiętym URN kluczem identyfikacji.
    
    If foundRow > 0 Then                            ' Decydujące dróg rozejście: Znaleziono numer stary wiersz > 0 czyli jest to plik w bazie Excela ze starym wpisem!
        UpdateExistingFile ws1, foundRow, cols, csvHeaders, csvColMapping, colDict, urnCol, pathCol, currentUpdater, changesCount ' Wymuś wykonanie Sub-procedury UPDATE z licznikiem i odcieniowaniem żółtym podświetlania zmian (oraz rzuć notki w komórkę na modyfikatora/datę zmiany - wersjonowanie wsteczne w chmurce EXCEL!).
        If urn <> "" Then dictExcelURNs(urn) = foundRow ' Uzupełnij ewentualny nowy zepsuty Hash (z trybu awaryjnego po starych wyszukanych po 'Nazwie' po starym pustym pliku z błędem w bazie). Hash się zaktualizował jako foundRow! Excel znów po mapowaniu słownikowym zna ten rekord do duplikatów przy MarkDelete i innych
    Else                                            ' Odgałęzienie binarne - Wiersz "foundRow" po wszystkich walkach z Hash i Pętlą ręczną wyszedł zero. Co oznacza - Nowy plik ProjectWise wykazany w CSV ze skanowania Powłoki!.
        Dim newRow As Long                          ' Zmienna pomocnicza przyjmująca ZWRÓCONY z "AddNewFile(..)" numerek nowego utworzonego świeżego rzędu podpiętego gdzieś pośrodku EXCEL'a przez .Insert Shift w arkusz
        newRow = AddNewFile(ws1, headerRow, cols, csvColMapping, colDict.count, urnCol, pathCol, currentUpdater, dictExcelURNs, dictExcelFolders, changesCount) ' Zlecenie "INSERT SHIFT" i wypełnienia całej reszty białego i nowego wiersza Exela atrybutami, oraz dodanie go w statystyki, plus updateowanie logiki hash poprzez byReference zmiennej Słowników bo rzędy wszystkie "zostały zepchnięte i ich Hash indeks padł (np z rzędu 510 zjechał w Exel na 511)". Ta funkcja zmusza Hash aktualizację w locie by nie stracić 511.
        If urn <> "" Then dictExcelURNs(urn) = newRow ' Wrzuć "ten nowy wiersz jako nowo-powiązany klucz URN do zaktualizowanego Słownika" - tak aby "MarkDelete(usuwające)" oraz inne w Excel pętli omijały i nie wywalały go jako Delete na koniec kodu w markDelete
    End If                                          ' Koniec wariantu i całkowitej sub-obróbki jednego "rzędu pliku z bazy CSV" i powrót do wielkiej pętli nad CSV elementami.
End Sub                                             ' Wyjście zwrotne
    
Private Function FindOrCreateFolderHeader(ws1 As Worksheet, actualFolder As String, folderUrn As String, urnCol As Long, totalCols As Long, ByRef dictExcelFolders As Object) As Long ' Znalezienie Nagłówka / Tła na zielono (Foldera Grupującego pliki w pliku)
    Dim headerRow As Long                           ' Zmienna deklaracyjna
    Dim searchKey As String                         ' Klucz szybkiego dostępu (Ujednolicony i bez spacji do hash - Upper Case by zrównać litery bo to folder operacyjny)
    searchKey = UCase(Trim(actualFolder))           ' Zamiana literał na ogromne pod odczyt szybkiego wariantu dictionary (Dlatego w Hash deklarowano ignore case by nie kłóciło się ze zmienną na serwerze i ujednoliconą pod Windows Search)
    
    If dictExcelFolders.exists(searchKey) Then      ' Jeśli słownik starych Excel zielonych nagłówków rozpoznał ten wpis na liście z "CSV actualFolder"...
        headerRow = dictExcelFolders(searchKey)     ' Masz jego numerek prosto z Cache Hash Excela! Podstaw pod Zmienną (Np. Numer: 520 rządek pod Folderek A\B\Polska)!
        If folderUrn <> "" And InStr(1, folderUrn, "pw://", vbTextCompare) > 0 And urnCol > 0 Then ' Uaktualnienie gdy brakuje zepsutego/niewyczyszczonego linku [ Open Folder ] jako klikalnego url na arkuszu! (Uzupełnia przy aktualizacji brakujący detal do klikania w komórkę na zielono - i otwiera ją na platformie BENTLEY z poziomu EXEL!)
            If ws1.Cells(headerRow, urnCol).Hyperlinks.count = 0 Then ' Test awaryjny - hiperłącze padło z formatowania użytkownika? Dodaj! (Tylko jeśli zero = Hyperlink zepsuty/niewidoczny)
                ws1.Hyperlinks.Add Anchor:=ws1.Cells(headerRow, urnCol), Address:=folderUrn, TextToDisplay:="Open folder in PW" ' Narzędzie osadzające skrót pod zielony nagłówek z HyperLink VBA (Wymuszenie wektora na komórkę)
                ws1.Cells(headerRow, urnCol).Font.Bold = True ' Pogrubienie Hyperłącza pod stylizacje - żeby ładniej wyglądało i pasowało.
            End If                                  ' Koniec sprawdzania czy brakowało i wrzucania hiperłącza
        End If                                      ' Zakończenie pełnej pętli modyfikacji brakujących łącz na starych zielonych Excelowych Headerach
    Else                                            ' Odnoga if'a z Hash Search -> Jeśli zielony FOLDER nagłówek "nie istnieje (zwraca z hash pustkę dla ActualFolder z CSV.. znaczy się - że w bazie ProjectWise narodził i zrodził się zupełnie nowy system folderowy!)"
        Dim lastR As Long                           ' Szukanie ostatniego pustego dachu na dole Excela dla zbudowania nowego drzewka - i wstawienia zielonego paska na dole by dokleić mu potem wpisy puste białego!
        Dim lastRCell As Range                      ' Obiekt Range w celu interakcji po stronie EXCEL Arkusza i "wyprowadzenia ataku na pozycję od tyłu na zlokalizowanie podkładu"
        Set lastRCell = ws1.Columns(1).Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious) ' Jak zwykle, znalezienie najbardziej skrajnego zapisu w Arkusz by wiedzieć, gdzie doczepić Nowy Element Bazy (Doczep do końca pociągu nowy wóz Zielony pod Grupę folderową z ProjectWise!)
        If Not lastRCell Is Nothing Then lastR = lastRCell.Row Else lastR = 11 ' Gdy skraj wierszowy został zidentyfikowany to przypisz go. Fallback na nagłówkowy 11 w przypadku braku lub destrukcji całego szkieletu Excelem by nie wywalać z błędu -1 lub "0 is invalid range!"
        If lastR < 11 Then lastR = 11               ' Twardy stop dla zepsutego Nothing gdy plik jest ułomny (od Rząd = 11 ładuj)
        
        headerRow = lastR + 1                       ' Zjedź o jedną dół od końca i utwórz "TUTAJ BĘDZIE RZĄD NA FOLDER ZIELONY" jako pozycja dla nowego Elementu Grupy w excel
        ws1.Cells(headerRow, 1).Value = actualFolder ' Zapisz Nazwę jego Path w kolumnę pierwszą na świeżym białym dole!
        ws1.Range(ws1.Cells(headerRow, 1), ws1.Cells(headerRow, totalCols)).Interior.Color = CLR_FOLDER ' Oznacz pozycję do końca nagłówka kolorów RGB jako JASNY ZIELONY Z FORMATOWANIA
        ws1.Cells(headerRow, 1).Font.Bold = True    ' Pogrub czcionkę na Nazwie ścieżki i podziału pliku
        
        If folderUrn <> "" And InStr(1, folderUrn, "pw://", vbTextCompare) > 0 And urnCol > 0 Then ' Dorzuć nowy świeży klikalny URL Linker (Zabezpieczenie przy braku URL jak zwykle - sprawdzanie PW:// aby upewnić z bazą) do ProjectWise po instalacji pod zielony blok (Nowo narodził się folder a więc dorób jego odpowiednik hiperłącze w EXCEL!)
            ws1.Hyperlinks.Add Anchor:=ws1.Cells(headerRow, urnCol), Address:=folderUrn, TextToDisplay:="Open folder in PW" ' Instrukcja dodania dla nowej komórki hyperlinka wizualnego (z tekstem '[ OPEN FOLDER ]' by nie zagracać pliku setką długich szarych krzaków pod URN)
            ws1.Cells(headerRow, urnCol).Font.Bold = True ' Pogrub klikalny dodany element hiperłącza by było ładnie i logicznie odróżnialne
        End If                                      ' Zakończenie procesu dokładania klikacza pod Folder!
        dictExcelFolders(searchKey) = headerRow     ' Aktualizacja na boku pamięci Hash dla "Dictionary Zielonego Folderu" o najświeższy nowo wytworzony wpis na pozycję HeaderRow pod plikiem Excel (Aby przy następnych plikach CSV należących od nowa do 'TEGO SAMEGO ZIELONEGO NAGŁÓWKA NA DOLE - znalazło po nim z łatwością ten sam klucz i dorzuciło już jako 'Istniejący Zielony - Insert pod Spód białym' a nie by mnożył zielone pod sobą - Bo hash został teraz zaaktualizowany w locie!)
    End If                                          ' Końcowa akcja pętli kreowania w głowach / grupowaniu zielonych wpisów odznaczających hierarchie drzew z PW na płaskim modelu tabeli Excel CSV
    FindOrCreateFolderHeader = headerRow            ' Zwrot dla zmiennej na funkcję, o wskazaniu ostatecznego rzędu EXCEL (np 351 lub 951 jako nowo dodany dla pustego pliku na dole!), gdzie ten docelowo "Wyszukany/Skonstruowany przed chwila" Folder ma Rząd. Pozwala to operacji pod spodem do rzucenia pod ten 951 lub 351 pliku (Biały wiersz pod spód z informacjami) w kodach "AddNewFile / Update").
End Function                                        ' Koniec budowy główki Zielonej pod pliki

Private Sub UpdateExistingFile(ws1 As Worksheet, targetRow As Long, cols() As String, csvHeaders() As String, csvColMapping() As Long, colDict As Object, urnCol As Long, pathCol As Long, currentUpdater As String, ByRef changesCount As Long) ' Kod dokonujący Modyfikacji i aktualizacji Pliku do wersji Wierszów (Update Data from delta na starym rekordzie i żółci komórki na różnice!).
    Dim lastCol As Long: lastCol = colDict.count    ' Ściągnięcie ostatniej mapowanej struktury ze spłaszczonego Hash słownika atrybutów dla Excela jako LastCol (By usunąć np kolory czerwone z Delete z poprzednich odpaleń aż po brzeg mapy)
    
    If ws1.Cells(targetRow, 1).Interior.Color = CLR_DELETED Then ' Jeśli plik w starym excelu posiadał kolor RGB czerwień i był Strikethrough(Skreślony) - z procedury MarkDeleted z poprzedniego dnia (Oznaczony do kosza bo wczoraj go ktoś wywalił w ProjectWise a dzisiaj wrócił!) to makro zdejmie ten kolor na gładko
        ws1.Range(ws1.Cells(targetRow, 1), ws1.Cells(targetRow, lastCol)).Interior.Color = xlNone ' Zresetuj podłoże na brak koloru (Normalny stary biały rządek powrócił po śmietniku chmury ProjectWise jak zmartwychwstały zza Grobu i w locie CSV znów mu odczyta parametry)
        ws1.Range(ws1.Cells(targetRow, 1), ws1.Cells(targetRow, lastCol)).Font.Strikethrough = False ' Ściągnij błąd przekreślenia pliku jak wyżej z czcionki na false i odśwież
    End If                                          ' Koniec wariantu de-red zombiaczego pliku powracającego na wykaz bazy (Update przy zmartwychwstaniu)
    
    Dim maxC As Long, c As Long                     ' Ograniczniki w kolumnach CSV
    maxC = UBound(cols)                             ' Zmienna z max kolumn jakie w tablicach ma CSV ściągnięte od powłoki w linii podziału Potoku (Splita na wierszu wczytywanego Stringa wyżej)
    If maxC > UBound(csvHeaders) Then maxC = UBound(csvHeaders) ' Normalizacja zapobiegająca błędowi uBound Exception gdy PowerShell wrzucił o jedną/dwie nadmiarowe pałeczki '|' potoku za dużo na końcówkę i tablica cols ma 15 pól, a nagłówków było CSV_HEAD 13 - Utnij to nadmiarowe puste i sprawdzaj do max 13 (Bezpieczeństwo na krawędzi blatu).
    
    For c = 0 To maxC                               ' Kręć się po kolumnach i porównuj komórka CSV <-> Excel komórka na poszukiwania zmian "Żółci - Delta"! Od kolumny zera (np nazwy w pliku)
        Dim targetCol As Long                       ' Oznaczenie kolumny Excel pod która wskazuje Hash Array od kolumny C w array z pliku CSV
        targetCol = csvColMapping(c)                ' Pobranie z mapy np dla C = 4 (Oznacza Datę w CSV kolumnie) Maping wskazuje że Data na Excel "TargetCol" to Kolumna 'F'. Wpisz pod to numer Excel kolumny np 6 i uderz w Excel target! To działa po to bo Hash map pozwala Excelowi mieć 'Opis' na kolumnie 'Z' a w bazie z CSV ma na pozycji '2'. To się samo znajduje i przekleja prawidłowo by użytkownik na 1 arkuszu mógł przesuwać na boki sobie sam atrybuty wedle kaprysu na myszce po Arkuszu! To kluczowa sprawa Makra by nie nadpisać mu nazwy na Opisie i nie robić hardcode! Target col to ratuje przed tragedią
        If targetCol > 0 Then                       ' Jesli mapowanie wydało cokolwiek (Czyli nie równa się zero, puste = 'False Map' - zignoruj i nie wrzucaj zmian)
            Dim oldVal As String, newVal As String  ' Rezerwacja pamięci w pętli. Bufor starego wpisu (Odczyt EXCEL) oraz nowego String od powłoki (z podziału rzędu pliku CSV na zmienne tablic cols)
            oldVal = Trim(CStr(ws1.Cells(targetRow, targetCol).Value)) ' Odczyt na czysto wycięty i sformatowany do testu ze starego pliku Excela
            newVal = Trim(cols(c))                  ' Nowy string wejściowy jako New Value
            
            If targetCol <> urnCol And targetCol <> pathCol And csvHeaders(c) <> "Updater" Then ' Wyłącz techniczne stałe plikowe, URN i ścieżkę jako "zmiany atrybutów żółte", updater - one są stałymi ukrytymi danymi do silników powiązań bazy i nie interesują użytkownika na warianty "zmiany wersji i żółte pola modyfikacji".
                Dim isDiff As Boolean               ' Logiczna brama decyzyjna do wyzwalacza zmiany flagi (Na start False=Nie Różni Się). Skrypt omija fałszywe powiadomienia (false positives). Poniżej jest skomplikowany system detekcji różnicy w datach - Np "Excel myśli że 20-03 = 20-March-2023 i zapisze jako Double a CSV da ci tekst String = 20.03.2023". Dla VBA to są DWA RÓŻNE WPISY = Rzuci żółcią że zaktualizował ale dla usera nic się tam nie zmieniło i uzna program za uszkodzony gówniany system Excel! Brama IS DIFF to filtruje
                isDiff = False                      ' Start domyślnej blokady na false.
                
                If StrComp(oldVal, newVal, vbTextCompare) <> 0 Then ' Weryfikacja czystego StrComp z ignorowaniem 'A' na 'a' po Text Compare dla zepsutych przez uzytkownika wielkosci. Jesli jednak się RÓŻNI po podstawowym - przepusc do dalszej analizy false positive isDiff = True
                    isDiff = True                   ' Odbezpiecz flagę zmiany na 'Prawda = Odnotowano Różnice!'
                    Dim dOld As Date, dNew As Date  ' Rozpoczęcie filtra Dat False-positives. (Blokuje rzucanie zmianą gdy 'Data = data').
                    On Error Resume Next            ' Blok awaryjny dla błędów gdy nie potrafi sformatować bzdurnych starych wpisów EXCEL (np CData z tekstu 'Brak') - Ominie i nie rzuci awarią do loga i nie zamknie okna!
                    If IsDate(oldVal) And IsDate(newVal) Then ' Czy stary Excel wpis i nowy wpis jest datą z systemu - Sprawdź w 1 milisekunde w locie po komórkach pod IsDate z C++! Zaczyna się walka o to by Excel nie kolorował pustych dat na źle.
                        dOld = CDate(oldVal)        ' Sformatuj awaryjnie stary string Excel na prawdziwy obiekt Double Data z CDate(String)
                        dNew = CDate(newVal)        ' Odkoduj nowy na tak samo CDate Date obiekt. Teraz mają tą samą strukturę. Data to double po kodowaniu Windows
                        If Err.Number = 0 Then      ' Przetestuj czy przy formatowaniu (powyżej od On Error Resume) padł na pysk CDate z winy np zepsutej litery w dacie '21-Marzcck=2012'. Jeśli przeszedł = Number Zero
                            If Int(dOld) = Int(dNew) Then isDiff = False ' Główny myk - Zrównaj do czystej daty całkowitej Double pod Int, czyli np bez ułamków na godzinach jeśli CSV podał "20-03-12 15:40" a excel zapamiętał w komórkach "20-03-12". Wyłącz po Int = Int fałszywą odnotowaną zmianę i uratuj plik przed zżółceniem! Ustaw z powrotem False flagę bo obie te daty pomimo innych dopisek Stringów mają ten sam numer dni całkowity Int!
                        End If                      ' Zamknięcie testu zero po błędach CDate na liczbach
                    End If                          ' Zakończenie filtrowania problemu systematycznego CDate pod EXCEL - ominięto żółty błąd aktualizacji po niczym dla oka
                    Err.Clear                       ' Oczyszczenie brudów awaryjnych "On Error Resume"
                    If isDiff And IsNumeric(oldVal) And IsNumeric(newVal) Then ' Jeśli flaga 'isDiff' to nadal zmieniona flaga = True i jest to zmiana. A jednocześnie oldVal oraz NewVal to Double Numeryczny ułamek - Testuj ten fałszywy wpis po drugie na format liczb by excel nie wyrzucił że 0.20 to inna liczba niż 0,20 !! (To ten sam numer i go wygasimy poniżej). Sprawdzam czy Old i New to liczby?
                        If CDbl(Replace(oldVal, ".", ",")) = CDbl(Replace(newVal, ".", ",")) Then isDiff = False ' Wyrównaj zamieniając ewentualną złą kropkę CSV na stary styl Excel Comma przecinka ułamków i wepchnij do formatu CDbl - ułamka, jeśli są w konwersji CDbl równe to zgaś i wyrównaj ten brud do false! Excel nie podświetli po raz setny tego jako zmianę! Bo z poziomu liczby double to to samo ułamek pomimo kropki
                    End If                          ' Zakończ obsługę CDbl (false pos. na numerach)
                    On Error GoTo 0                 ' Odbezpiecz system awaryjny po skończonym filtrowaniu fałszywych alarmów (Delta isDiff logic) - Przywróć błędy i zgłoszenia makra na twardo
                End If                              ' Koniec testu i powrót głównego węzła zmian
                
                If isDiff Then                      ' Ostatecznie: Pętla oceniła po odfiltrowaniu zepsutych komórek z Excel (Daty / Kropki Przecinki CDbl / Typ Text Spacji Case itd) "PRAWDA JEST! IsDiff. To NOWA ZMIANA ATRYBUTU DELTA!". Podświetl, wrzuć wersje do historii i zaprezentuj żółć!
                    If oldVal = "" Then             ' Odrzucenie systemu notek w komentarzu jeśli stary pole w excel było puste od początku... Excel by się pluł o dopięcie tam pustki. Wrzuca tam po prostu "Czysty wpis bez podświetleń na żółto przy pierwszym ściąganiu z bazy od białej/pustej starej planszy Excel" (W ten sposób tylko nowe puste ściągania nic nie zżółknią).
                        ws1.Cells(targetRow, targetCol).Value = newVal ' Po prostu zaktualizuj plik od PW i nadpisz po cichu pustą kratę bez rzucania zmian po oczach bo jej "Stara wersja" nie istniała do historii. Zignoruj AddLog do historii atrybutu "Old->New"
                    Else                            ' Mamy do czynienia ze zmianą nad wpisem który "ISTNIAŁ" jako oldVal - Oznacz jako Różnicę i zaktualizuj Historię Komórki Komentarzem żółtym - Excel Comment (Wersjonowanie)!
                        Dim cmtText As String       ' Definicja do wpisu Komórkowego (prawy przycisk na Excel z notatką komórkową widoczną z góry czerwoną strzałką myszki ukrytą).
                        cmtText = Format(Now, "yyyy-mm-dd HH:mm") & " " & csvHeaders(c) & "=" & oldVal & " (User: " & currentUpdater & ")" ' Zbuduj z parametrów ciąg historyczny do Excel (np: 2026-03-24 Nazwa=Projekt_Budowlany (User: Jan Kowalski)). Pokazuje dokładnie "Do jakiej starej wersji zmieniono od kogo i kiedy!"
                        
                        Dim cmt As Comment          ' Utwórz wirtualny obiekt do sprawdzenia - Co jeśli komórka Excel miała już "Komentarz pod sobą stary z poprzedniego tygodnia (Kolejna 5 zmiana pliku PW!)"
                        Set cmt = ws1.Cells(targetRow, targetCol).Comment ' Sprawdź obiekt Komórki pod względem doczepionego na gładko Comment. Wrzuć go do "cmt" i sprawdz jego obecność w Pamięci Ram.
                        If cmt Is Nothing Then      ' Zero powiązań z "Cmt" (Nic się nie przyczepiło). Pierwsza historyczna zmiana od stworzenia pliku w chmurze - dodajemy nowy dymek komórkowy Comment!
                            ws1.Cells(targetRow, targetCol).AddComment cmtText ' Przypnij od razu sformatowany cmtText i dodaj "Bąbelek" dla docelowego miejsca nowej zmiany.
                        Else                        ' Nie zero "Cmt", tam już istnieje dymek i zarys historyczny - Należy go odczytać i zrobić po chamsku "Append (Znak i dodanie z nową linią na spód)" by rozszerzyć listę do dołu!
                            cmt.Text Text:=cmt.Text & vbCrLf & cmtText ' Wywołaj polecenie wektora modyfikacji Text:= pobierając najpierw ze wskaźnika swój własny cmt.Text oraz doczep do tego vbCrLf (Linijka w Dół ENTER systemowy) oraz doczep aktualną zmianę cmtText zbudowaną wyżej. (System dodawania historii zmian atrybutów działa z Excel bez spowalniania jako C++ backend)!
                        End If                      ' Zamknij dodawania komentarza wstecznego dla Wersjonowania Excela Delty ProjectWise
                        
                        ws1.Cells(targetRow, targetCol).Interior.Color = CLR_CHANGED ' Zmień tło całej widocznej kratki komórki na ostrzegawczy Żółty na arkuszu Excel żeby Użytkownik go widział jako "Różnica - DELTA CHANGED!".
                        ws1.Cells(targetRow, targetCol).Value = newVal ' Aktualizuj również jej fizyczny nowy środek po CSV paczce, by po otwarciu miała najświeższe PW dane!
                        
                        If targetCol = 1 Then       ' Warunkowanie małego logu u dołu. Jeśli zaktualizował samą Naze pliku "Name" (Kolumna Excel = 1 to nazwy). Zgłoś nazwę
                            AddLogEntry "CHANGED", newVal, csvHeaders(c), oldVal, newVal, currentUpdater, RGB(255, 242, 204) ' Rzuć wpis do Arkusza Zmian "Log" rzędu 'CHANGED' z NewVal by nazwa pliku nie rzucała starym i się myliła z nową modyfikacją (Log jest pomarańczowy RGB). Zapis "Nowa -> Stara"
                        Else                        ' Jeśli to modyfikacja 'Opisu' albo innego Atrybutu niż główna nazwa "Kolumna Excel inna niż 1"
                            AddLogEntry "CHANGED", ws1.Cells(targetRow, 1).Value, csvHeaders(c), oldVal, newVal, currentUpdater, RGB(255, 242, 204) ' Rzuć staryą zaktualizowaną w systemie nazwę do logowania
                        End If                      ' Koniec generowania z logiem atrybutów Update do Arkusza.
                        changesCount = changesCount + 1 ' Dodano jedną Zmianę - Wrzuc na Zmienną sumującą by podliczyć statystyki dla MsgBox i logu powiadomień po przeleceniu całego pliku.
                    End If                          ' Zakończ główny system sprawdzający puste zmiany Notek wersji dla tego Atrybutu.
                End If                              ' Zakończenie IsDiff filtru - po odrzutach nic się nie wydarzy po nim w pętli. Kod działa gładko.
            Else                                    ' Przypadek drugi od "If<> urnCol/Path itp" z góry, gdy Target Kolumny = To URN Link "PW" (Klikalny Adres). Ten wpis systemowy od platformy musi być nadpisywany z cicha.
                ws1.Cells(targetRow, targetCol).Value = newVal ' Brutalne proste "Aktualizuj wpis tekstu pod komórkę linka pod spód bez zmian na Żółto bo to systemowy link od powłoki, wklej go na biało!"
                If targetCol = urnCol And newVal <> "" And InStr(1, newVal, "pw://", vbTextCompare) > 0 Then ' Ale dla URN Link musi być coś na "Klin" by znów był hiperłączem otwierającym system jak po usunięciach user format - Wymuś sprawdzenie hiperłącza z tekstem pw i uruchom wstawkę awaryjną na obiekt
                    If ws1.Cells(targetRow, targetCol).Hyperlinks.count = 0 Then ' Gdy ten systemowy wektor "Urn" był na czysto (Zgubił "Niebieski Link Klikalny W Excel") - Bo użytkownik wkleił bez formatowania plik z góry? Zabezpiecz plik w locie.
                        ws1.Hyperlinks.Add Anchor:=ws1.Cells(targetRow, targetCol), Address:=newVal, TextToDisplay:=newVal ' Doczep na nowo hiperłącze pod obiekt by Excel potraktował systemowy podkład jako otwierany po "Klik" do otworzenia platformy
                    Else                            ' Miał Hiperłącze, ale adres url może być inny od tego co miał po nowym eksporcie CSV bazy plik zmienił go!
                        ws1.Cells(targetRow, targetCol).Hyperlinks(1).Address = newVal ' Aktualizuj środek/Cel starego doczepionego Hyperłącza - nadpisanie systemowe pod maską w 0 ms!
                        ws1.Cells(targetRow, targetCol).Hyperlinks(1).TextToDisplay = newVal ' Aktualizuj wygląd pod spodem (System to uaktualni dla nowego adresu nowej chmury dla pewności na czysto z CSV!)
                    End If                          ' Zakończenie nadpisywania URN bez żółtych flag - z weryfikacją poprawności "Click Url!" w komórce EXCEL.
                End If                              ' Zamykanie sprawdzania Target Link / URN wariantów dla atrybutów systemowych (ukrytych przed wersją dla "Delty").
            End If                                  ' Zamykanie IF Target Col operacji i bloków filtrowania atrybutów kolumny z wiersza
        End If                                      ' Zakończenie "Zabezpieczenia" (Jeśli równe Zero = fałszywy map Hash table "Klucz pusty/Zepsuty z mapowania od początku") - Zignoruj i przesuń po pętli.
    Next c                                          ' Skacz pętlą For do kolejnej C "Kolumny z CSV z atrybutem i ładuj jej właściwości do Excel komórek po wierszu!"
End Sub                                             ' Kaskada po wykonaniu aktualizacji Zółci (Diff Delt) i ukrytych parametrów dla Pół EXCEL z CSV wierszy i zwolnij zasób wyjścia do Pętli Wyższej nad rzędami.

Private Function AddNewFile(ws1 As Worksheet, headerRow As Long, cols() As String, csvColMapping() As Long, totalCols As Long, urnCol As Long, pathCol As Long, currentUpdater As String, ByRef dictExcelURNs As Object, ByRef dictExcelFolders As Object, ByRef changesCount As Long) As Long ' Dodawanie Elementów Plików Pustych do zielonego "Koszaru (Folder Header)" pod nowo wpisane struktury z powłoki powiązania bazy CSV
    Dim insertRow As Long, c As Long                ' Deklaracja zmiennej do trzymania wskaźnika Wstawianego elementu na pustym i C (Kolumn z Atrybutami - For C pętle dla ładowania New File)
    insertRow = headerRow + 1                       ' Zjedź jeden niżej (Oczko Puste Białego Tła pod zielony nagłówek folderu - Tu wstawiamy plik by znalazł się 'W nim / Pod Nim'). Do loop wciśnie na spód ten wiersz po niższym pętlach o ile inne już nie zapełniały jego wnętrzności!
    
    Do While ws1.Cells(insertRow, 1).Value <> "" And Trim(ws1.Cells(insertRow, pathCol).Value) <> "" ' Zabezpieczenie przed "Przebiciem Excela"! Kręć pętlą dopóki ta jedna kratka (headerRow+1) w dół pod Folderem Zielonym JEST ZAJĘTA (Ma inną starą nazwę i inną ścieżkę!). Znaczy to że folder jest pełen plików starych, omiń na spód szukając białego pola. Przebij go w głąb jego "Wnęki na białe wiersze Excel z listą" aż nie natrafisz na pole puste lub koniec pliku na wpis
        insertRow = insertRow + 1                   ' Dokładaj plus jedynkę i zjeżdżaj o "Zajęty stary plik z bazy". Tak długo i przesuwając na nowe czyste miejsce do włożenia "Spodem Pudełka zielonego koszyka Listy" i pętluj aż warunek wyda "Empty Cell!". Włóż ten Insert Row w ramy pustego pudła plikowego - "Tędy Wejdzie!"
    Loop                                            ' Zamknij wyszukiwarkę dna grupy folderów
    
    ws1.rows(insertRow).Insert Shift:=xlDown        ' Mechanizm Excel "Shift Down!" - Zrób miejsce po pustym, zepchnij o jedną linię CAŁĄ RESZTĘ ARKUSZA Z INNYMI ZIELONYMI FOLDERAMI DO TYŁU! (Zrób insert Pusty biały u dołu kosza) a stara lista zachowa swoją pozycję pod zielonymi folderami na dół bo została ucięta .Insert'em EXCEL a nie pętlą nadpisującą makra.
    
    Dim k As Variant                                ' Krytyczna część dla Insert Shift! (Ponieważ Excel został "rozerwany" na dwie częsci i zepchnięty, to co było np Rzędem = 100 spadło u dołu i ma Rząd = 101 po dodaniu 1 insert Shift na pozycji np 50!) Hash jest Zrujnowany jeśli użył pamięci pod "Rządek starego Dict= 100 dla URN linku starego pliku pod folderem X, a po zejściu ShiftxlDown on znalazł się niżej, jak go makro Update-Existing wywoła by go zedytować na 100 wywali mu na pliku nad nim bo to 101!!!". Napraw Słowniki po zepchnięciu Shift DOWN
    For Each k In dictExcelURNs.Keys                ' Przeleć pętlą po wszystkich ukrytych kluczach po nazwach Hash pamięci. Zaktualizuj ich "Value" na prawą kolumnę czyli numerki.
        If dictExcelURNs(k) >= insertRow Then dictExcelURNs(k) = dictExcelURNs(k) + 1 ' Jeśli to powiązanie URN w hash miało NUMER rzędu równy bądź WYŻSZY(niżej w Excel od Insert = Został on Zepchnięty u dołu) dodaj mu pod wartość zapamiętanego powiązania "+1". Podbiłeś właśnie jego klucz w Hash i połączyłeś z wierszem po Update! To działa błyskawicznie w tysięcznych sekundy dla całej bazy w HashMap Excel Array bez wywalania Excel Shift Down!
    Next k                                          ' Przeskocz pod następny w kluczu i napraw Hash plik z Dict Excel URNs (Pliki)
    For Each k In dictExcelFolders.Keys             ' Powtórz operację "Naprawa podbicia Numerycznego po Insert Excel Zsuwu Shift" dla Słownika powiązań Hash Tabel 'Zielonych nagłówków od drzewa ProjectWise', co się zepchnęły.
        If dictExcelFolders(k) >= insertRow Then dictExcelFolders(k) = dictExcelFolders(k) + 1 ' Dodaj jedynkę wyrównując powiązanie u dołu, pomijając te od góry Insert co leżą nieruszone.
    Next k                                          ' Skocz na koniec kluczy - zaktualizuj Dict (Słownik Folders!). Ocalono cały plik Excel przed chaosem
    
    For c = 0 To UBound(cols)                       ' Wracamy na utarte szlaki białego, gotowego po Insert nowego pustego pola z uaktualnionymi Dictionary Hash Tablicami. Wypełnij kolumny dla C (atrybutu rzędu plikowego nowej dostawy do bazy).
        If c <= UBound(csvColMapping) Then          ' Upewnij że znów C mieści się w dopuszczalnym zakresie indeksów po Mapach by nie sypnąć tablicy powiązań od 'Pustego - Ostatniego pustego wiersza w Splice Potokowego CSV'
            Dim tCol As Long                        ' Deklaracja ułatwiacza powiązań
            tCol = csvColMapping(c)                 ' Mapuj tak samo by Kolumna w pliku CSV na rzędzie 6 - zmapowała się na Excelu np pod G, unikając błędnych wpisów atrybutu (Zgodnie z ułożeniem usera na G).
            If tCol > 0 Then                        ' Potwierdź zmapowany 'prawdziwy atrybut docelowy Excel = >0' dla ominiętych 0 z fałszywych nagłówków nie dla podglądu (np Linki Folder Urn = Zera! Nie wstawiaj ich w widoczne numery na biały Excel 1,2..).
                ws1.Cells(insertRow, tCol).Value = cols(c) ' Zwykłe chamskie wstawienie nowego Atrybutu do nowej wstawionej pod Insert linii pod zielony folder, ładując nowość bez żółci jako "Świeży biały wiersz z CSV bez wersji".
                If tCol = urnCol And cols(c) <> "" And InStr(1, cols(c), "pw://", vbTextCompare) > 0 Then ' Ale dla linków systemowych doczep po raz enty z powrotem funkcję Hyperlink (Na stałe wpisane by przy generowaniu URN klikalne było otwierane poprzez Baza danych Platformy Bentley, sprawdzające i odcinające pusty brak wpisu przy instalacji).
                    ws1.Hyperlinks.Add Anchor:=ws1.Cells(insertRow, tCol), Address:=cols(c), TextToDisplay:=cols(c) ' Odepnij wizualnie i uformuj przycisk, używając od razu pełnego URN w wyświeltaniu bo "Taki chamski był URN na excelu, plik musi widać systemowo kod URN do odczytania powiązań z Dict, Hyperlink jest mu tylko osłoną".
                End If                              ' Zamknij zabezpieczenia URN dla NewLine / NewFile
            End If                                  ' Zakończ obsługę poprawnej TCol odczytanej z Map CSV dla komórek
        End If                                      ' Zakończ omijanie błędów przepełnień "Złomków/Cieńkich Ramek Splite | po C++"
    Next c                                          ' Zapętli dla reszty atrybutów nowo-powstającego wpisu z CSV...
    
    ws1.Range(ws1.Cells(insertRow, 1), ws1.Cells(insertRow, totalCols)).Interior.Color = xlNone ' Zdejmij Formatowanie by usunąć np stary kolor Żółty z insertowanego na zielonym czy biały pociągnięty po Shifcie z góry po Insertowaniu, narzucając pustkę koloru bez skaz (Biały) dla Czystości bazy na koniec wiersza do krańcówki TotalCols mapy - 'By biała linia wyczyściła błędy wiersza doczepionego i była Czysto biała'
    ws1.Range(ws1.Cells(insertRow, 1), ws1.Cells(insertRow, totalCols)).Font.Bold = False ' Zdjęcie powikłań czcionki pod Shiftowane wstawienie obok Nagłówków itp by nowy wpis był jak w "Normalnym szarym Pliku a nie Zielonej Grupie pogrubionej". Wyrównanie do Normalności
    ws1.Range(ws1.Cells(insertRow, 1), ws1.Cells(insertRow, totalCols)).Font.Strikethrough = False ' Zdejmij uszkodzone pod wstawkami z czerwoną linią przekreślenie ułomne. Ocali od fałszywych oznaczeń
    
    AddLogEntry "ADDED", cols(0), "-", "-", "-", currentUpdater, RGB(226, 239, 218) ' Loguj po zakończeniu! Wypuść "Zzieleniałe (Dodano Nowość - RGB zieleni wpis)" logowanie u dołu dla nowo wrzuconego Atrybutu Pliku ("-" dla odczytu pod puste i w nową pusta bo to pierwsze tchnienie).
    changesCount = changesCount + 1                 ' Dodaj powiadomienie że makro coś tam na arkuszu "zmajstrowało" (Pliki zmienione i Delta jest > 0!) by odpalić MsgBox ze statusem na dnie procedur głownych
    AddNewFile = insertRow                          ' ZWROT dla funkcji w c++ / VBA. Odeślij na wynik procedury ten "Gładki, nowowygenerowany, zmapowany po zsuwach InsertRow Numerek. By Główna procedura z Update Hash Dict wiedziała jak 'On ma teraz nowy rząd (np. ten plik ma = 345 z wygenerowanego)' i wstrzyknęła w główny Hash z duplikatami do Delete omijając go! "AddNewFile = np 345"! Powrót
End Function                                        ' Koniec ciężkiego odłamu z wciskaniem na chama w środek pełnego układu Excel z zsuwem!

Private Sub MarkDeletedItems(ws1 As Worksheet, urnCol As Long, pathCol As Long, lastCol As Long, ByVal numScopes As Long, ByVal arrCleanScopes As Variant, dictTxtLinks As Object, dictOkScopes As Object, ByRef changesCount As Long) ' Kod końcowy skryptu z usuwaniem elementów! Odpalany na tył po skończonym wgrywaniu nowości CSV vs Delta do bazy by nie uciekły (Wykonane odczyty do tyłu szukające martwych wpisów z Excel które ominęła Pętla Główna bo one "Zniknęły i nigdy do pętli wyżej z CSV nie wpadły jako Wpis!")
    Dim maxR As Long, maxRCell As Range             ' Range do poszukiwań "Koniec Exela 1 kartki".
    Set maxRCell = ws1.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious) ' Jak zawsze szybki wyszukiwacz Excelowy dna i granicy wpisów z wierszy do ostatniej "Gwiadzki" z czegokolwiek (*), od tyłu po komórkach arkuszy pliku.
    If Not maxRCell Is Nothing Then maxR = maxRCell.Row Else maxR = 11 ' Gdy skraj został zidentyfikowany to przypisz go, z fallback do 11 na pustym startowym ułomnym
    If maxR < 11 Then Exit Sub                      ' Ochroń by nie szukał po wierszach nr np. 3 gdy zepsuto arkusz 1 ręcznym usunięciem by rzucić i zawiesić w OutOfBounds
    
    Dim rExcel As Long                              ' Deklaracja od góry pętli dla szukania w Excel na odwrót po wierszach.
    For rExcel = 12 To maxR                         ' Główna pętla w pętli. Idzie po CAŁYM EXCEL OD SAMEJ GÓRY NAGŁÓWKÓW - Ale po EXCEL Wpisach zamiast CSV! I bada je wszystkie po kolei w tym co pozostało nienaruszone - "Z weryfikacją na Dict i Log z PowerShell = PRAWDA"
        Dim targetPath As String, exURN As String, fileNameEx As String ' Rezerwacja bufora do wysyłania sprawdzeń (Sprawdzająca z Excel Path "Ścieżki folderów pod nim" czy exURN = "Unikalny stary wpis pod Excelem - np starożytny PW://Link... "). Oraz pobranie nazwy na zapas.
        fileNameEx = Trim(ws1.Cells(rExcel, 1).Value) ' Zapis "Nazwa Starego Excel File do String"
        targetPath = Trim(ws1.Cells(rExcel, pathCol).Value) ' Zapis "Ścieżka do zapytania sprawdzającego na drzewo z txt powiązań z powłoki"
        If fileNameEx = "" Then GoTo SkipRow        ' Jeśli stary plik to "Pusty Śmieć ze złą nazwą = np wpis puste, lub biały pas Excel ręcznie narzucony lub inny brud". Przepnij ucieczką w pętli "GOTO SkipRow (Na Koniec Do NexT z Ominięciem Całego Bloku Poniżej)!"
        
        If ws1.Cells(rExcel, urnCol).Hyperlinks.count > 0 Then ' Sprawdź czy stara komórka nie wylosowała sobie starego dobrego Hyperlink jako odczytu... (Trzeba pobierać z bazy na stałe Hyperlink po adresie linku zamiast z samego Stringa EXCELA jeśli ten tam się znalazł "Pod maską w VBA = Z linka"). Ocalić ułomne URL.
            exURN = Trim(ws1.Cells(rExcel, urnCol).Hyperlinks(1).Address) ' Wyrwij od powiązanego w EXCEL Wierszu w Kolumnie URN (Niebieski link pod tekstem ukryty) i zapisz w surowym text URL z obcięciem spacji dla starej zmiennej exURN.
        Else                                        ' W innym przypadku
            exURN = Trim(CStr(ws1.Cells(rExcel, urnCol).Value)) ' Odrzuć hyperlink (nie miało po prostu zepsute!) i wyciąg tekstowy kod surowy tekst "pw://.." na starca exURN bez obcinania URL. Z samej CStr na pewniaka Excel Formatowania Komórki.
        End If                                      ' Zakończenie odczytu URN EXCEL
        
        exURN = Replace(exURN, "%7B", "{", 1, -1, vbTextCompare) ' BARDZO WAŻNE: Odzyskanie po eksporcie HyperLink z systemu "Twardo - Przeglądarkowych % Znaków Zastępczych". Hyperlink VBA robi %7B podczas ukrywania pod spód arkusza - Czego Pętla nie zrozumie! Odtwórz Klamrowe Otwarcie "%7B -> {" (By porównało na czysto %7B jako Czyste Hash)
        exURN = Replace(exURN, "%7D", "}", 1, -1, vbTextCompare) ' Odtwórz zamknięcie powiązane w Project Wise Hash - Nawias "} zamykający zamiast webowych ułomnych znaków ASCII po starym hyperlink %7D"
        exURN = Replace(exURN, "%20", " ", 1, -1, vbTextCompare) ' Webowy kod spacji z URL zepsuł by Dict-Txt hash z TXT Log od PowerShell - Odrzuć '%20' na spacji normalnej ASCII by odnalazł zgodność
        exURN = Replace(exURN, vbCr, "")            ' Obetnij ukryte zniekształcenia potoku od Mac jeśli Excel HyperLink wniósł dziwne potoki
        exURN = Replace(exURN, vbLf, "")            ' Odłup końcówki nowej lini Linuxowe (Jeśli URN się zwinęło z wklejenia)
        exURN = Trim(exURN)                         ' Spłaszcz jeszcze raz i podetnij by był gotowy, odżywiony, naprawiony URL (exURN = "Czysty")
        
        If targetPath <> "" And InStr(1, exURN, "pw://", vbTextCompare) > 0 Then ' Najwyższa faza weryfikacji pliku pod usunięcie lub błąd. "Czy Stary Exel z Tego rzędu pod 'rExel' - To w ogóle jest 'STARY WYKAZ PLIKU' czy zwykły Zielony Nagłówek Grupy!". Plik posiada 'PW:// ' i Path? TZN JEST Weryfikowanym Do Delete elementem CSV!
            Dim searchTarget As String              ' Zabezpiecz wyciągnięty w starym excel target Path w formacie "Ucase" z jednorodnym backslash "\" dla wyszukiwania zgodności katalogów na drzewie (Żeby nie pisał Delete dla całego innego wariantu)
            searchTarget = UCase(Replace(targetPath, "/", "\")) ' Przepchana przez Backslash C:\
            
            Dim isUnderScope As Boolean             ' Sprawdzenie awaryjne (Czy ten stary EXCEL 'Katalog' jest z nami dzisiaj 'Skanowany od nowa z bazy Sheet2?' - Żeby nie robić "DELETE W ZEPSUTYM KOSZU" gdy rzucimy makro bez skanowania tego foldera dzisiaj (wykasowaliśmy go w sheet 2) co by podświetliło fałszywie miliony "Nie zeskanowanych plików jako usunięte z ProjectWise w bazie Czerwienią" - Weryfikacja Scopes chroni bazę Exela dla pliku z powłoki!)
            isUnderScope = False                    ' Domyślnie NIE JEST W ODCZYTACH z TXT Folders - Odmów wywalenia (Do Not Delete)
            
            Dim i As Long                           ' Otwarcie licznika na pętle
            For i = 1 To numScopes                  ' Dla każdego foldera wrzuconego w TXT Scopes list z PowerShell do Exela na początku macro...
                Dim cleanScope As String            ' Oczyść scope nazwe (Root'a) od Powłoki pod szukanie na dysku twardym Excel w Target ścieżkach pliku w rExel.
                cleanScope = UCase(Replace(arrCleanScopes(i), "/", "\")) ' Wydziel w 1 format C:\ i "DRZEWO\DRZEWO2\" (Wielkie Litery)
                If dictOkScopes.exists(cleanScope) And InStr(1, searchTarget, cleanScope, vbTextCompare) > 0 Then ' Jeśli "Target stary EXCEL np. C:\Baza\Stary" znajduje się w wywołanym dzisiaj skanie na Root "C:\Baza\" z CSV Powłoki...
                    isUnderScope = True             ' Znam ten obszar, był dzisiaj skanowany od stóp do głów w PowerShell CSV TXT. - FLAGA = "Możesz bezpiecznie wywołać Delete/Czerwony Kosz jeśli pliku 'Uciekło z Txt Log Tree'".
                    Exit For                        ' Kończenie testowania folderu (ZNALEZIONO POWIĄZANIE DLA ZNAJDUJĄCEGO SIĘ SCOPE!). Zamykamy Scope For po znalezieniu prawdy!
                End If                              ' Wyjście dla InStr testu na Stringach
            Next i                                  ' Poszukaj po następnym Txt Scope czy nie jest powiązany stary Excel plik z inną gałęzią od powłoki w Sheet2!
            
            If isUnderScope Then                    ' Mamy PRAWDA NA ZGODNOŚĆ Z ZAKRESEM DZISIAJSZYCH BADAŃ (Ten folder wchodził pod skanowanie. Możemy bezpiecznie osądzić plik o "DELETED" z serwera!)
                If Not dictTxtLinks.exists(exURN) Then ' PRAWIDŁOWY TEST - Czy ten 'Naprawiony ze znaków EX-Urn w Excel' NIE ISTNIEJE w TXT "Log.Txt z Drzewka Zbudowanego pod Powłoką, gdzie zapisał wszystkie zrzucone aktualnie linki ze 100% bazy Bentley!". Skoro NIE MA GO (Klucz Nie istnieje - Hash odrzuca False)... To uważa on że ten plik NIE Istnieje u niego po skanie, tzn = WYRZUCILI GO Z PROJECTWISE (USUNIĘTY LUB ZMIENIŁ FOLDER!). EXEL POSIADA GO FAŁSZYWIE!
                    If ws1.Cells(rExcel, 1).Interior.Color <> CLR_DELETED Then ' A dodatkowo plik ten "JUŻ NIE ZOSTAŁ ZMALTRETOWANY na kolor Czerwony Czysty = Usunięty w Excel". Tzn że 'Dopiero co go wyrzucili - jest świeżynką'.
                        ws1.Range(ws1.Cells(rExcel, 1), ws1.Cells(rExcel, lastCol)).Interior.Color = CLR_DELETED ' ZMALUJ TEN Wiersz na RGB Czerwony Koszmar do KRAŃCA PLIKU Excela! ("OZNACZENIE EXCEL_DELETE")
                        ws1.Range(ws1.Cells(rExcel, 1), ws1.Cells(rExcel, lastCol)).Font.Strikethrough = True ' Dodatkowo Strikethrough Czerwonym Przekreśleniem od "Czerwona Linia" (Wyróżni uszkodzenie pliku przed okiem) po całej linii Rzędu do Last Col!
                        AddLogEntry "DELETED", fileNameEx, "-", "Existed", "Deleted from PW", "-", CLR_DELETED ' Zarzuć mu czerwony Log na samym dole by podpadł pod "Delete History (Usunięto Z Projektu ProjectWise)" do wglądu Excel Log
                        changesCount = changesCount + 1 ' Dolicz jako statystyka do ogłoszeń Makra zmian jako Czysta Różnica bazy - MsgBox wyskoczy z radosną nowiną (Mimo tragedii w usunięciu) by nie omijać logiki >0
                    End If                          ' Zakończ obsługę odmalowania dla nie odznaczonych truposzy w starym pliku EXCEL.
                Else                                ' "JEST ODNOGA DRUGA - ODNALAZŁ GO HASH! DictTxt Odnalazł go pod sobą!". Plik Txt miał go wygenerowanego i wykrzyczał 'Jestem URN-XYZ'. Plik z Exela był w Koszu/lub go nie było a teraz go "Trafiłem". Więc on POWRÓCIŁ. LUB NIGDY GO NIE SKASOWALI A SKRYPT GO ZŁAPAŁ
                    If ws1.Cells(rExcel, 1).Interior.Color = CLR_DELETED Then ' Jeśli odzyskano URN który w starym excel był np Czerwony bo w zeszłym tygodniu usunęli plik do kosza z serwera? Ożył bo ktoś przywrócił backup na Cloud?
                        ws1.Range(ws1.Cells(rExcel, 1), ws1.Cells(rExcel, lastCol)).Interior.Color = xlNone ' Zdejmuj Kolor Czerwień ze Starego! (Oczyść Wiersz Excel R z Czerwonego i oddaj białe) - ZMARTWYCHWSTANIE
                        ws1.Range(ws1.Cells(rExcel, 1), ws1.Cells(rExcel, lastCol)).Font.Strikethrough = False ' Zdejmuj czerwoną linię przekreślającą by wrócił normalny plik Excel
                    End If                          ' Zakończenie "Zmartwychwstania" i przywracania czerwonych plam z bazy
                End If                              ' Zakończenie węzła Logiki usunięć w starym exelu (If Txt Exists Hash Search -> False = Mark / True = Save and DeMark).
            End If                                  ' Zamknięcie testu powiązania skanowanych folderów "Scope / Drzewo w CSV Sheet_2" (isUnderScope). By uniknąć wściekłej czerwieni dla miliona ominiętych folderów z Sheet_2 wyłączonych z bazy dzisiaj!
        End If                                      ' Zakończenie głownego bloku sprawdzającego "Czy ten wiersz starego exela to FOLDER zielony czy PLIK w ogóle!" po odczytach "PW://"
SkipRow:                                            ' Zakotwiczenie punktu powrotnego ucieczki Goto (W przypadku gdy plik miał np puste wiersze/zepsute szare brudy nie od bazy a od użytkonwika i nie był nawet brany pod uwagę "Dziadostwo Ręczne - Zepsuta Nazwa"). GOTO Rzuca z góry kod tutaj z pominięciem setki linijek na twardo
    Next rExcel                                     ' Zapętlij dla Excela (Przejdź z powrotem by testować Mark Delete starych plików do usunięcia z kosza (do następnego rExcel wyżej do końca maxR na dole pliku)).
End Sub                                             ' Kaskadowe opuszczenie algorytmu weryfikacyjnego i Delta z powrotem.

Public Sub ApplySearch(ByVal txt As String)         ' Pokazuje tylko wiersze zawierajace fraze (w dowolnej kolumnie); folder widoczny, gdy ma pasujacy plik
    Dim ws As Worksheet, lastR As Long, lastC As Long, r As Long, c As Long, n As Long
    Set ws = ThisWorkbook.Sheets(SH_DATA)
    lastR = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    If lastR <= DATA_HEADER_ROW Then Exit Sub
    txt = Trim(txt)
    Application.ScreenUpdating = False
    ws.rows(DATA_HEADER_ROW + 1 & ":" & lastR).Hidden = False
    If Len(txt) >= 2 Then
        lastC = ws.Cells(DATA_HEADER_ROW, ws.Columns.count).End(xlToLeft).Column
        Dim v As Variant
        v = ws.Range(ws.Cells(DATA_HEADER_ROW + 1, 1), ws.Cells(lastR, lastC)).Value
        n = UBound(v, 1)
        Dim hit() As Boolean, isFolder() As Boolean, anyHit As Boolean
        ReDim hit(1 To n): ReDim isFolder(1 To n)
        For r = 1 To n
            isFolder(r) = (Trim(CStr(v(r, 2))) = "" And Trim(CStr(v(r, 1))) <> "")
            For c = 1 To lastC
                If Not IsError(v(r, c)) Then
                    If InStr(1, CStr(v(r, c)), txt, vbTextCompare) > 0 Then hit(r) = True: Exit For
                End If
            Next c
        Next r
        anyHit = False
        For r = n To 1 Step -1                      ' Od dolu: plik pasuje -> widoczny; folder widoczny, jesli pasuje on sam lub ktorykolwiek jego plik
            If isFolder(r) Then
                ws.rows(DATA_HEADER_ROW + r).Hidden = Not (anyHit Or hit(r))
                anyHit = False
            Else
                ws.rows(DATA_HEADER_ROW + r).Hidden = Not hit(r)
                If hit(r) Then anyHit = True
            End If
        Next r
    End If
    Application.ScreenUpdating = True
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
    If MsgBox("Operacja zintegrowana powiodła się. Czy chcesz usunąć utworzone pliki tymczasowe (TXT, PS1, CSV)?", vbYesNo + vbQuestion, "Czyszczenie środowiska") = vbYes Then ' Spytaj MsgBoxem przed skasowaniem by uratować logi dla np trybu developerskiego lub inspekcji błędów
        On Error Resume Next                        ' Odpalenie systemu ochronnego przed błędem w kasowaniu (np. PowerShell wciąz wisi i powłoka zblokowała plik przed Exelem pod LockFile)
        fso.DeleteFile txtPath                      ' Strzel po poleceniu zniszczenia pliku tekstowego z Drzewem Scopes z Dysku C:\
        fso.DeleteFile psPath                       ' Usuń również plik Generatora skryptu Powershell *.ps1
        fso.DeleteFile csvPath                      ' Wykasuj ogromny plik CSV ze strumienia by zwolnić potworne zapychacze pamięci po całej robocie
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
        If Trim(CStr(ws3.Cells(r, 1).Value)) <> "" Then ' Puste wiersze są pomijane - tak samo jak w GetAttributesFromSheet3, więc kolejność kolumn zgadza się z CSV
            ws1.Cells(11, cCol).Value = Trim(CStr(ws3.Cells(r, 1).Value))
            StyleHeader ws1.Cells(11, cCol)
            cCol = cCol + 1
        End If
    Next r
End Sub
