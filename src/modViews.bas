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
        MsgBox "Nie ma widoku '" & viewName & "' w arkuszu " & SH_VIEWS & ".", vbExclamation
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
    
    Application.ScreenUpdating = False
    Dim old As Range
    Set old = wsV.Range(wsV.Rows(VIEW_HEADER_ROW), wsV.Rows(wsV.Rows.Count))
    old.Hyperlinks.Delete
    old.Clear
    old.ClearComments
    
    If n > 0 And lastR >= DATA_HEADER_ROW Then
        For i = 1 To n
            wsD.Range(wsD.Cells(DATA_HEADER_ROW, colList(i)), wsD.Cells(lastR, colList(i))).Copy wsV.Cells(VIEW_HEADER_ROW, i)
            wsV.Columns(i).ColumnWidth = wsD.Columns(colList(i)).ColumnWidth
        Next i
        Application.CutCopyMode = False
    End If
    wsV.Range("D3").Value = "Kolumn: " & n & "  |  Wierszy: " & IIf(lastR > DATA_HEADER_ROW, lastR - DATA_HEADER_ROW, 0) & "  |  Odswiezono: " & Format(Now, "yyyy-mm-dd hh:nn")
    wsV.Range("D4").Value = IIf(missing = "", "", "Nie znaleziono kolumn: " & Left(missing, Len(missing) - 2))
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
