Attribute VB_Name = "modViews"
Option Explicit

' =======================================================================
' DASHBOARD = OKNO NA PW_Data (bez kopiowania bazy)
' Definicja widoku (arkusz Views, kolumna B): lista kolumn PW_Data w kolejnosci wyswietlania, oddzielona , lub ;
'   - nazwa naglowka:  File Name, Version, WorkflowState
'   - litery kolumn:   A, B, F, L, TK   (tez bez przecinkow: A B F L TK)
'   - *  = wszystkie kolumny
' Dashboard pokazuje jedna strone (PAGE_SIZE wierszy) wynikow zapytania: filtr folderu (nawigator) + szukanie. Cala baza zostaje w PW_Data.
' =======================================================================
Private mBusy As Boolean

' Kolumna PW_Data dla tokenu: nazwa naglowka albo litery kolumny; 0 = brak
Private Function ResolveToken(ByVal tok As String) As Long
    Dim c As Long, ch As Long, isLetters As Boolean
    For c = 1 To gC
        If StrComp(gHdr(c), tok, vbTextCompare) = 0 Then ResolveToken = c: Exit Function
    Next c
    isLetters = (Len(tok) >= 1 And Len(tok) <= 3)
    For ch = 1 To Len(tok)
        If Not (UCase(Mid(tok, ch, 1)) >= "A" And UCase(Mid(tok, ch, 1)) <= "Z") Then isLetters = False
    Next ch
    If isLetters Then
        On Error Resume Next
        c = ThisWorkbook.Sheets(SH_DATA).Range(tok & "1").Column
        On Error GoTo 0
        If c >= 1 And c <= gC Then ResolveToken = c
    End If
End Function

' Zamienia definicje widoku na liste kolumn; missing = nierozpoznane
Private Function ViewColumns(ByVal spec As String, ByRef missing As String) As Long()
    Dim out() As Long, n As Long, i As Long, c As Long, p As Variant, s As Variant
    ReDim out(1 To gC + 60)
    spec = Trim(spec)
    If spec = "" Or spec = "*" Then
        For i = 1 To gC
            If gHdr(i) <> "" Then n = n + 1: out(n) = i
        Next i
    Else
        For Each p In Split(Replace(spec, ";", ","), ",")
            If Trim(CStr(p)) <> "" Then
                c = ResolveToken(Trim(CStr(p)))
                If c > 0 Then
                    n = n + 1: out(n) = c
                ElseIf InStr(Trim(CStr(p)), " ") > 0 Then       ' "A B F" bez przecinkow
                    For Each s In Split(Trim(CStr(p)), " ")
                        If CStr(s) <> "" Then
                            c = ResolveToken(CStr(s))
                            If c > 0 Then n = n + 1: out(n) = c Else missing = missing & CStr(s) & ", "
                        End If
                    Next s
                Else
                    missing = missing & Trim(CStr(p)) & ", "
                End If
            End If
        Next p
    End If
    If n = 0 Then n = 1: out(1) = 1
    ReDim Preserve out(1 To n)
    ViewColumns = out
End Function

Public Sub BuildView()                              ' zgodnosc wsteczna - odswieza Dashboard
    RenderDashboard False
End Sub

Public Sub RenderDashboard(Optional ByVal resetPage As Boolean = False)
    If mBusy Then Exit Sub
    mBusy = True
    Dim wsV As Worksheet, wsDef As Worksheet
    Dim stp As String
    On Error GoTo Fail
    stp = "start"
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    Set wsDef = ThisWorkbook.Sheets(SH_VIEWS)
    Dim prevEv As Boolean, prevScr As Boolean
    prevEv = Application.EnableEvents: prevScr = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    
    EnsureData
    Dim viewName As String, spec As String, r As Long, defRow As Long, lastDef As Long
    viewName = Trim(CStr(wsV.Range(VIEW_SELECTOR).Value))
    lastDef = wsDef.Cells(wsDef.Rows.Count, 1).End(xlUp).Row
    If viewName = "" Then viewName = Trim(CStr(wsDef.Cells(VIEWS_FIRST_ROW, 1).Value)): wsV.Range(VIEW_SELECTOR).Value = viewName
    For r = VIEWS_FIRST_ROW To lastDef
        If StrComp(Trim(CStr(wsDef.Cells(r, 1).Value)), viewName, vbTextCompare) = 0 Then defRow = r: Exit For
    Next r
    If defRow > 0 Then spec = Trim(CStr(wsDef.Cells(defRow, 2).Value))
    
    stp = "query"
    Dim cols() As Long, nCols As Long, missing As String
    cols = ViewColumns(spec, missing)
    nCols = UBound(cols)
    If gC = 0 Then nCols = 0
    Dim total As Long, pages As Long, page As Long, k As Long, first As Long
    total = RunQuery(CStr(wsV.Range(VIEW_SEARCH).Value), CStr(wsV.Range(FOLDER_CELL).Value))
    pages = (total + PAGE_SIZE - 1) \ PAGE_SIZE: If pages < 1 Then pages = 1
    page = 1
    If Not resetPage Then page = Val(CStr(wsV.Range(PAGE_CELL).Value))
    If page < 1 Then page = 1
    If page > pages Then page = pages
    wsV.Range(PAGE_CELL).Value = page
    first = (page - 1) * PAGE_SIZE + 1
    k = total - first + 1: If k > PAGE_SIZE Then k = PAGE_SIZE
    If k < 0 Then k = 0
    
    stp = "theme"
    Dim subT As String, warnT As String
    subT = nCols & " columns  |  " & total & " results  |  refreshed " & Format(Now, "yyyy-mm-dd hh:nn")
    If defRow = 0 Then warnT = "View '" & viewName & "' not found in sheet " & SH_VIEWS
    If missing <> "" Then warnT = "Columns not found: " & Left(missing, Len(missing) - 2)
    SetPaging page, pages, total, CStr(wsV.Range(FOLDER_CELL).Value)
    RefreshChrome wsV, subT, warnT, CStr(ThisWorkbook.Sheets(SH_DATA).Range("A8").Value)
    
    stp = "clear"
    Dim t As TTheme
    t = CurrentTheme(wsV)
    Dim lastUsed As Long
    lastUsed = wsV.UsedRange.Row + wsV.UsedRange.Rows.Count - 1
    If lastUsed < VIEW_HEADER_ROW + PAGE_SIZE + 5 Then lastUsed = VIEW_HEADER_ROW + PAGE_SIZE + 5
    wsV.Range(wsV.Rows(VIEW_HEADER_ROW), wsV.Rows(lastUsed)).Clear
    wsV.Range(wsV.Rows(VIEW_HEADER_ROW), wsV.Rows(lastUsed)).Interior.Color = t.dataBg
    
    stp = "write"
    Dim iLink As Long, iSt As Long, j As Long, i As Long, ix As Long, v As String, arr() As Variant
    iLink = ColOf("PW Link"): iSt = ColOf(COL_STATUS)
    If nCols > 0 Then
        ReDim arr(1 To k + 1, 1 To nCols)
        For j = 1 To nCols: arr(1, j) = gHdr(cols(j)): Next j
        For i = 1 To k
            ix = gIdx(first + i - 1)
            For j = 1 To nCols
                v = CStr(gData(ix, cols(j)))
                If cols(j) = iLink And Left(v, 5) = "pw://" And Len(v) < 250 Then v = "=HYPERLINK(""" & v & """,""" & v & """)"
                arr(i + 1, j) = v
            Next j
        Next i
    End If
    Dim blk As Range
    If nCols > 0 Then
        Set blk = wsV.Range(wsV.Cells(VIEW_HEADER_ROW, 2), wsV.Cells(VIEW_HEADER_ROW + k, nCols + 1))
        blk.NumberFormat = "@"
        If k > 0 Then
            For j = 1 To nCols
                If cols(j) = iLink Then wsV.Range(wsV.Cells(VIEW_HEADER_ROW + 1, j + 1), wsV.Cells(VIEW_HEADER_ROW + k, j + 1)).NumberFormat = "General"
            Next j
        End If
        blk.Value = arr
        ' szerokosci kolumn z naglowka i probki danych
        Dim w As Long, m As Long, mx As Long
        For j = 1 To nCols
            mx = Len(arr(1, j)) + 3
            m = k: If m > 60 Then m = 60
            For i = 1 To m
                If Len(arr(i + 1, j)) > mx Then mx = Len(arr(i + 1, j))
            Next i
            If cols(j) = iLink Then mx = 40
            If mx < 10 Then mx = 10
            If mx > 60 Then mx = 60
            wsV.Columns(j + 1).ColumnWidth = mx
        Next j
    End If
    wsV.Columns(1).ColumnWidth = 44
    EnsureFrozenWidth wsV
    If Abs(wsV.Columns(3).Left - mFzUsed) > 1 Then DrawHeader wsV      ' szerokosc kolumn A:B sie zmienila - przerysuj panel, zeby nie przecinal granicy zamrozenia
    
    stp = "format"
    Dim lastC As Long: lastC = nCols + 1: If lastC < 12 Then lastC = 12
    With wsV.Range(wsV.Cells(VIEW_HEADER_ROW, 1), wsV.Cells(VIEW_HEADER_ROW, lastC))
        .Interior.Color = t.hdrBg: .Font.Color = t.hdrFg: .Font.Bold = True: .Font.Name = "Segoe UI": .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    wsV.Rows(VIEW_HEADER_ROW).RowHeight = 28
    wsV.Cells(VIEW_HEADER_ROW, 1).Value = "FOLDERS  (click = show all)"
    If k > 0 And nCols > 0 Then
        With wsV.Range(wsV.Cells(VIEW_HEADER_ROW + 1, 2), wsV.Cells(VIEW_HEADER_ROW + k, nCols + 1))
            .Font.Color = t.dataFg: .Font.Name = "Segoe UI": .Font.Size = 10
            .VerticalAlignment = xlCenter
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(205, 210, 220)
            .Borders(xlInsideVertical).LineStyle = xlContinuous      ' pionowe linie miedzy wszystkimi kolumnami
            .Borders(xlInsideVertical).Color = RGB(205, 210, 220)
        End With
        If iSt > 0 Then                              ' kolory wierszy wg Status (pasy tego samego statusu naraz)
            Dim s0 As Long, sNow As String, sPrev As String, clr As Long
            For i = 1 To k + 1
                If i <= k Then sNow = CStr(gData(gIdx(first + i - 1), iSt)) Else sNow = "#END"
                If i = 1 Then
                    sPrev = sNow: s0 = 1
                ElseIf sNow <> sPrev Then
                    Select Case sPrev
                        Case "NEW": clr = t.newBg
                        Case "CHANGED": clr = t.changedBg
                        Case "DELETED": clr = t.deletedBg
                        Case Else: clr = -1
                    End Select
                    If clr >= 0 Then wsV.Range(wsV.Cells(VIEW_HEADER_ROW + s0, 2), wsV.Cells(VIEW_HEADER_ROW + i - 1, nCols + 1)).Interior.Color = clr
                    s0 = i: sPrev = sNow
                End If
            Next i
        End If
    End If
    
    Application.ScreenUpdating = prevScr: Application.EnableEvents = prevEv
    mBusy = False
    Exit Sub
Fail:
    Dim em As String
    em = Err.Description & " [Dashboard: " & stp & "]"
    Application.ScreenUpdating = True: Application.EnableEvents = True
    mBusy = False
    Notify "Error: " & em
End Sub

' --- stan interfejsu (ukryte komorki) i akcje ---
Public Sub SetFolderFilter(ByVal key As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_VIEW)
    ws.Range(FOLDER_CELL).Value = key               ' zmiana komorki uruchamia odswiezenie (Workbook_SheetChange)
End Sub

Public Sub ClearFolderFilter()
    SetFolderFilter ""
End Sub

Public Sub PagePrev()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_VIEW)
    If Val(CStr(ws.Range(PAGE_CELL).Value)) > 1 Then ws.Range(PAGE_CELL).Value = Val(CStr(ws.Range(PAGE_CELL).Value)) - 1
End Sub

Public Sub PageNext()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_VIEW)
    ws.Range(PAGE_CELL).Value = Val(CStr(ws.Range(PAGE_CELL).Value)) + 1     ' RenderDashboard przytnie do ostatniej strony
End Sub

' Zmiana ktorejs z ukrytych komorek stanu: odswiez Dashboard (strona 1 przy zmianie widoku/szukania/folderu)
Public Sub OnDashboardChange(ByVal Target As Range)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_VIEW)
    If Not Intersect(Target, ws.Range(VIEW_SEARCH & "," & FOLDER_CELL & "," & VIEW_SELECTOR)) Is Nothing Then
        RenderDashboard True
    ElseIf Not Intersect(Target, ws.Range(PAGE_CELL & "," & THEME_CELL)) Is Nothing Then
        RenderDashboard False
    End If
End Sub

Public Sub ApplySearch(ByVal txt As String)         ' PW_Data to czysta baza - szukanie jest w Dashboard (zgodnosc wsteczna: nic nie robi)
End Sub
