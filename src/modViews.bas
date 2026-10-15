Attribute VB_Name = "modViews"
Option Explicit

' =======================================================================
' WIDOKI: Baza_PW wyswietlona wg szablonu z arkusza Widoki.
' Definicja widoku (kolumna B): lista kolumn Baza_PW w kolejnosci wyswietlania, oddzielona , lub ;
'   - nazwa naglowka:  File Name, Version, WorkflowState
'   - litery kolumn:   A, B, F, L, TK   (tez bez przecinkow: A B F L TK)
'   - *  = wszystkie kolumny
' Baza_PW nie jest zmieniana - widok to osobna kopia w arkuszu Widok.
' =======================================================================
Public Sub BuildView()
    Dim wsV As Worksheet, wsD As Worksheet, wsDef As Worksheet
    Dim viewName As String, spec As String, r As Long, defRow As Long, lastDef As Long
    Dim lastR As Long, lastC As Long, i As Long
    Dim stp As String
    On Error GoTo Fail
    stp = "start"
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    Set wsD = ThisWorkbook.Sheets(SH_DATA)
    Set wsDef = ThisWorkbook.Sheets(SH_VIEWS)
    
    viewName = Trim(CStr(wsV.Range(VIEW_SELECTOR).Value))
    If viewName = "" Then Exit Sub
    lastDef = wsDef.Cells(wsDef.Rows.Count, 1).End(xlUp).Row
    For r = VIEWS_FIRST_ROW To lastDef
        If StrComp(Trim(CStr(wsDef.Cells(r, 1).Value)), viewName, vbTextCompare) = 0 Then defRow = r: Exit For
    Next r
    If defRow = 0 Then
        Notify "Nie ma widoku '" & viewName & "' w arkuszu " & SH_VIEWS & "."
        Exit Sub
    End If
    spec = Trim(CStr(wsDef.Cells(defRow, 2).Value))
    
    lastC = wsD.Cells(DATA_HEADER_ROW, wsD.Columns.Count).End(xlToLeft).Column
    lastR = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    
    Dim colList() As Long, n As Long, missing As String
    ReDim colList(1 To lastC + 50)
    If spec = "" Or spec = "*" Then
        For i = 1 To lastC: n = n + 1: colList(n) = i: Next i
    Else
        Dim parts() As String, p As Variant, sub1 As Variant, c As Long
        parts = Split(Replace(spec, ";", ","), ",")
        For Each p In parts
            If Trim(CStr(p)) <> "" Then
                c = ResolveColumn(wsD, Trim(CStr(p)), lastC)
                If c > 0 Then
                    n = n + 1: If n > UBound(colList) Then ReDim Preserve colList(1 To n + 50)
                    colList(n) = c
                ElseIf InStr(Trim(CStr(p)), " ") > 0 Then       ' "A B F" bez przecinkow
                    For Each sub1 In Split(Trim(CStr(p)), " ")
                        If CStr(sub1) <> "" Then
                            c = ResolveColumn(wsD, CStr(sub1), lastC)
                            If c > 0 Then
                                n = n + 1: If n > UBound(colList) Then ReDim Preserve colList(1 To n + 50)
                                colList(n) = c
                            Else
                                missing = missing & CStr(sub1) & ", "
                            End If
                        End If
                    Next sub1
                Else
                    missing = missing & Trim(CStr(p)) & ", "
                End If
            End If
        Next p
    End If
    
    stp = "czyszczenie"
    Application.ScreenUpdating = False
    Dim old As Range
    Set old = wsV.Range(wsV.Rows(VIEW_HEADER_ROW), wsV.Rows(wsV.Rows.Count))
    old.ClearOutline
    old.EntireRow.Hidden = False                    ' po filtrowaniu Szukaj - odkryj wiersze przed przebudowa
    old.Hyperlinks.Delete
    old.Clear
    old.ClearComments
    
    stp = "kopiowanie"
    If n > 0 And lastR >= DATA_HEADER_ROW Then
        For i = 1 To n
            wsD.Range(wsD.Cells(DATA_HEADER_ROW, colList(i)), wsD.Cells(lastR, colList(i))).Copy wsV.Cells(VIEW_HEADER_ROW, i + 1)
            wsV.Columns(i + 1).ColumnWidth = wsD.Columns(colList(i)).ColumnWidth
        Next i
        Application.CutCopyMode = False
        BuildTree wsV, wsD, lastR
    End If
    stp = "opis"
    wsV.Range("D3").Value = "Kolumn: " & n & "  |  Wierszy: " & IIf(lastR > DATA_HEADER_ROW, lastR - DATA_HEADER_ROW, 0) & "  |  Odswiezono: " & Format(Now, "yyyy-mm-dd hh:nn")
    wsV.Range("D4").Value = ""
    If missing <> "" Then wsV.Range("D4").Value = "Nie znaleziono kolumn: " & Left(missing, Len(missing) - 2)
    wsV.Range("A7").Value = CStr(wsD.Range("A8").Value)    ' status ostatniej aktualizacji z Baza_PW
    stp = "filtr"
    FilterRows wsV, VIEW_HEADER_ROW, CStr(wsV.Range(VIEW_SEARCH).Value)   ' zachowaj aktywne wyszukiwanie
    If Len(Trim(CStr(wsV.Range(VIEW_SEARCH).Value))) < 2 Then CollapseFiles wsV
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Err.Raise Err.Number, , Err.Description & " [BuildView: " & stp & ", kolumna " & i & "]"
End Sub

' Kolumna A widoku = drzewo folderow (wciecie wg glebokosci sciezki); pliki folderu sa pod nim, od kolumny B, i da sie je zwinac (+/- z lewej)
Private Sub BuildTree(wsV As Worksheet, wsD As Worksheet, ByVal lastR As Long)
    Dim n As Long, r As Long, d() As Long, isF() As Boolean, txt As String, minD As Long, k As Long
    n = lastR - DATA_HEADER_ROW
    If n < 1 Then Exit Sub
    ReDim d(1 To n): ReDim isF(1 To n)
    minD = 99
    For r = 1 To n
        If wsD.Cells(DATA_HEADER_ROW + r, 1).Interior.Color = CLR_FOLDER Then
            isF(r) = True
            txt = Replace(CStr(wsD.Cells(DATA_HEADER_ROW + r, 1).Value), "/", "")
            d(r) = Len(txt) - Len(Replace(txt, "", ""))
            If d(r) < minD Then minD = d(r)
        End If
    Next r
    With wsV.Cells(VIEW_HEADER_ROW, 1)
        .Value = "Drzewo folderow"
        .Interior.Color = CLR_HEADER: .Font.Color = CLR_WHITE: .Font.Bold = True
    End With
    wsV.Columns(1).ColumnWidth = 44
    wsV.Outline.SummaryRow = xlSummaryAbove
    Dim lvl As Long, fileLvl As Long
    fileLvl = 1
    For r = 1 To n
        k = VIEW_HEADER_ROW + r
        If isF(r) Then
            lvl = d(r) - minD + 1: If lvl > 7 Then lvl = 7
            txt = Replace(CStr(wsD.Cells(DATA_HEADER_ROW + r, 1).Value), "/", "")
            If Right(txt, 1) = "" Then txt = Left(txt, Len(txt) - 1)
            With wsV.Cells(k, 1)
                .Value = Mid(txt, InStrRev(txt, "") + 1)
                .IndentLevel = lvl - 1
                .Interior.Color = CLR_FOLDER: .Font.Bold = True
            End With
            wsV.Rows(k).OutlineLevel = lvl
            fileLvl = lvl + 1: If fileLvl > 8 Then fileLvl = 8
        Else
            wsV.Rows(k).OutlineLevel = fileLvl
        End If
    Next r
End Sub

' Zwija pliki pod folderami (foldery zostaja widoczne)
Public Sub CollapseFiles(wsV As Worksheet)
    Dim lastR As Long, r As Long
    lastR = wsV.UsedRange.Row + wsV.UsedRange.Rows.Count - 1
    For r = VIEW_HEADER_ROW + 1 To lastR
        If wsV.Cells(r, 1).Interior.Color <> CLR_FOLDER Then wsV.Rows(r).Hidden = True
    Next r
End Sub

' Sciezka folderu dla wiersza widoku (z kolumny 1 Baza_PW, z koncowym \)
Private Function FolderKey(ByVal vRow As Long) As String
    Dim t As String
    t = Replace(CStr(ThisWorkbook.Sheets(SH_DATA).Cells(vRow - VIEW_HEADER_ROW + DATA_HEADER_ROW, 1).Value), "/", "\")
    If Right(t, 1) <> "\" Then t = t & "\"
    FolderKey = t
End Function

' Klik w folder: pokazuje tylko ten folder, jego podfoldery i ich pliki
Public Sub ShowFolderRows(ByVal vRow As Long, Optional ByVal baseKey As String = "")
    Dim wsV As Worksheet, r As Long, lastR As Long, base As String, vis As Boolean, runStart As Long
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    lastR = wsV.UsedRange.Row + wsV.UsedRange.Rows.Count - 1
    base = IIf(baseKey = "", FolderKey(vRow), baseKey)
    Application.ScreenUpdating = False
    wsV.Rows(VIEW_HEADER_ROW + 1 & ":" & lastR).Hidden = True
    For r = vRow To lastR + 1
        If r <= lastR Then
            If wsV.Cells(r, 1).Interior.Color = CLR_FOLDER Then vis = (StrComp(Left(FolderKey(r), Len(base)), base, vbTextCompare) = 0)
        Else
            vis = False
        End If
        If vis And runStart = 0 Then runStart = r
        If Not vis And runStart > 0 Then wsV.Rows(runStart & ":" & (r - 1)).Hidden = False: runStart = 0
    Next r
    Application.ScreenUpdating = True
End Sub

' Folder z nawigatora: znajdz pierwszy wiersz folderu o tej sciezce (lub podfolderu) i pokaz jego zawartosc; zwraca wiersz albo 0
Public Function ShowFolderByKey(ByVal key As String) As Long
    Dim wsV As Worksheet, r As Long, lastR As Long
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    lastR = wsV.UsedRange.Row + wsV.UsedRange.Rows.Count - 1
    For r = VIEW_HEADER_ROW + 1 To lastR
        If wsV.Cells(r, 1).Interior.Color = CLR_FOLDER Then
            If StrComp(Left(FolderKey(r), Len(key)), key, vbTextCompare) = 0 Then
                ShowFolderRows r, key
                ShowFolderByKey = r
                Exit Function
            End If
        End If
    Next r
End Function

' Klik w naglowek "Drzewo folderow": pokaz wszystkie foldery, pliki zwiniete
Public Sub ShowAllFolders()
    Dim wsV As Worksheet, lastR As Long
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    lastR = wsV.UsedRange.Row + wsV.UsedRange.Rows.Count - 1
    If lastR <= VIEW_HEADER_ROW Then Exit Sub
    Application.ScreenUpdating = False
    wsV.Rows(VIEW_HEADER_ROW + 1 & ":" & lastR).Hidden = False
    CollapseFiles wsV
    Application.ScreenUpdating = True
End Sub

' Szukaj w Baza_PW (pole tekstowe na tym arkuszu)
Public Sub ApplySearch(ByVal txt As String)
    FilterRows ThisWorkbook.Sheets(SH_DATA), DATA_HEADER_ROW, txt
End Sub

' Pokazuje tylko wiersze zawierajace fraze (w dowolnej kolumnie); wiersz folderu (niebieski) widoczny, gdy pasuje on sam lub ktorys jego plik
Public Sub FilterRows(ws As Worksheet, ByVal headerRow As Long, ByVal txt As String)
    Dim lastR As Long, lastC As Long, r As Long, c As Long, n As Long
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastR <= headerRow Then Exit Sub
    txt = Trim(txt)
    Application.ScreenUpdating = False
    ws.Rows(headerRow + 1 & ":" & lastR).Hidden = False
    If Len(txt) >= 2 Then
        lastC = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        Dim v As Variant, hit() As Boolean, isFolder() As Boolean, anyHit As Boolean
        v = ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastR, lastC)).Value
        n = UBound(v, 1)
        ReDim hit(1 To n): ReDim isFolder(1 To n)
        For r = 1 To n
            isFolder(r) = (ws.Cells(headerRow + r, 1).Interior.Color = CLR_FOLDER)
            For c = 1 To lastC
                If Not IsError(v(r, c)) Then
                    If InStr(1, CStr(v(r, c)), txt, vbTextCompare) > 0 Then hit(r) = True: Exit For
                End If
            Next c
        Next r
        anyHit = False
        For r = n To 1 Step -1                      ' od dolu: plik widoczny, gdy pasuje; folder - gdy pasuje on lub ktorys jego plik
            If isFolder(r) Then
                ws.Rows(headerRow + r).Hidden = Not (anyHit Or hit(r))
                anyHit = False
            Else
                ws.Rows(headerRow + r).Hidden = Not hit(r)
                If hit(r) Then anyHit = True
            End If
        Next r
    End If
    Application.ScreenUpdating = True
End Sub

' Zwraca numer kolumny Baza_PW dla nazwy naglowka lub liter ("TK"); 0 = nie znaleziono
Private Function ResolveColumn(wsD As Worksheet, ByVal tok As String, ByVal lastC As Long) As Long
    Dim c As Long, ch As Long, isLetters As Boolean
    For c = 1 To lastC
        If StrComp(Trim(CStr(wsD.Cells(DATA_HEADER_ROW, c).Value)), tok, vbTextCompare) = 0 Then ResolveColumn = c: Exit Function
    Next c
    isLetters = (Len(tok) >= 1 And Len(tok) <= 3)
    For ch = 1 To Len(tok)
        If Not (UCase(Mid(tok, ch, 1)) >= "A" And UCase(Mid(tok, ch, 1)) <= "Z") Then isLetters = False
    Next ch
    If isLetters Then
        On Error Resume Next
        c = wsD.Range(tok & "1").Column
        On Error GoTo 0
        If c >= 1 And c <= lastC Then ResolveColumn = c
    End If
End Function
