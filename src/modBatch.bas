Attribute VB_Name = "modBatch"
Option Explicit

' =======================================================================
' BATCH SEARCH: wklej liste nazw dokumentow (arkusz Batch, kolumna A od wiersza 5),
' RUN sprawdza kazda nazwe w bazie (File Name, bez rozrozniania wielkosci liter, * i ? dozwolone),
' SHOW ALL MATCHES pokazuje wszystkie trafienia w Dashboard (zaznacz wiersze i pobierz je przyciskiem DOWNLOAD FROM PW).
' =======================================================================
Private Const FIRST_ROW As Long = 5

Public Sub OpenBatchSheet()
    Dim ws As Worksheet
    Set ws = EnsureBatchSheet()
    ws.Activate
    ws.Range("A" & FIRST_ROW).Select
End Sub

Private Function EnsureBatchSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_BATCH)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(SH_VIEW))
        ws.Name = SH_BATCH
        BuildBatchLayout ws
    End If
    Set EnsureBatchSheet = ws
End Function

Private Sub BuildBatchLayout(ws As Worksheet)
    Dim shp As Shape, wn As String
    wn = "'" & ThisWorkbook.Name & "'!"
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "BATCH SEARCH"
    ws.Range("A1").Font.Size = 18: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Paste document names (one per line) in column A from row " & FIRST_ROW & ". A name matches when it is contained in File Name (case-insensitive); * and ? wildcards are allowed."
    ws.Range("A2").Font.Color = RGB(90, 100, 120)
    ws.Range("A4:E4").Value = Array("Document name (paste here)", "Matches", "Status", "First match: path", "First match: PW Link")
    ws.Range("A4:E4").Font.Bold = True: ws.Range("A4:E4").Font.Color = RGB(255, 255, 255): ws.Range("A4:E4").Interior.Color = RGB(31, 56, 100)
    ws.Range("A" & FIRST_ROW & ":A5000").NumberFormat = "@"
    ws.Range("D" & FIRST_ROW & ":D5000").NumberFormat = "@"
    ws.Columns(1).ColumnWidth = 50: ws.Columns(2).ColumnWidth = 10: ws.Columns(3).ColumnWidth = 14: ws.Columns(4).ColumnWidth = 60: ws.Columns(5).ColumnWidth = 60
    Dim nm As Variant, cp As Variant, mc As Variant, k As Long, x As Double
    nm = Array("btnBatchRun", "btnBatchShow", "btnBatchClear")
    cp = Array("RUN SEARCH", "SHOW ALL MATCHES IN DASHBOARD", "CLEAR LIST")
    mc = Array("RunBatchSearch", "BatchToDashboard", "ClearBatch")
    x = 8
    For k = 0 To 2
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, ws.Range("A3").Top + 1, IIf(k = 1, 230, 130), 22)
        shp.Name = CStr(nm(k))
        shp.Fill.ForeColor.RGB = IIf(k = 0, RGB(31, 122, 61), IIf(k = 1, RGB(47, 84, 150), RGB(120, 120, 120)))
        shp.Line.Visible = msoFalse
        shp.TextFrame2.TextRange.Text = CStr(cp(k))
        shp.TextFrame2.TextRange.Font.Size = 9: shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.OnAction = wn & CStr(mc(k))
        x = x + shp.Width + 8
    Next k
    ws.Rows(3).RowHeight = 26
End Sub

' wzorce: pusty = pominiety wiersz; z * lub ? -> wzorzec Like (escape [ i #), inaczej szukanie tekstu
Private Function ReadNames(ws As Worksheet, ByRef pats() As String, ByRef wild() As Boolean) As Long
    Dim last As Long, r As Long, n As Long, s As String
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < FIRST_ROW Then Exit Function
    n = last - FIRST_ROW + 1
    ReDim pats(1 To n): ReDim wild(1 To n)
    For r = FIRST_ROW To last
        s = LCase(Trim(CStr(ws.Cells(r, 1).Value)))
        If s <> "" Then
            If InStr(s, "*") > 0 Or InStr(s, "?") > 0 Then
                s = Replace(Replace(s, "[", "[[]"), "#", "[#]")
                pats(r - FIRST_ROW + 1) = "*" & s & "*"
                wild(r - FIRST_ROW + 1) = True
            Else
                pats(r - FIRST_ROW + 1) = s
            End If
        End If
    Next r
    ReadNames = n
End Function

Private Function IsHit(ByVal v As String, ByVal p As String, ByVal w As Boolean) As Boolean
    If w Then IsHit = (v Like p) Else IsHit = (InStr(v, p) > 0)
End Function

Public Sub RunBatchSearch()
    Dim ws As Worksheet, pats() As String, wild() As Boolean, n As Long
    Dim iName As Long, iPath As Long, iLink As Long, k As Long, i As Long
    Set ws = EnsureBatchSheet()
    EnsureData
    iName = ColOf("File Name"): iPath = ColOf("PW Source Path"): iLink = ColOf("PW Link")
    If gN = 0 Or iName = 0 Then Notify "No data - run UPDATE DATA first.": Exit Sub
    n = ReadNames(ws, pats, wild)
    If n = 0 Then Notify "Paste document names in column A of sheet " & SH_BATCH & " (from row " & FIRST_ROW & ").": Exit Sub
    Dim low() As String
    ReDim low(1 To gN)
    For k = 1 To gN: low(k) = LCase(CStr(gData(k, iName))): Next k
    Dim out() As Variant, hits As Long, first As Long, v As String, nFound As Long, nMulti As Long, nMiss As Long
    ReDim out(1 To n, 1 To 4)
    Application.ScreenUpdating = False
    For i = 1 To n
        If pats(i) <> "" Then
            hits = 0: first = 0
            For k = 1 To gN
                If IsHit(low(k), pats(i), wild(i)) Then
                    hits = hits + 1
                    If first = 0 Then first = k
                End If
            Next k
            out(i, 1) = hits
            If hits = 0 Then
                out(i, 2) = "NOT FOUND": nMiss = nMiss + 1
            ElseIf hits = 1 Then
                out(i, 2) = "FOUND": nFound = nFound + 1
            Else
                out(i, 2) = "MULTIPLE": nMulti = nMulti + 1
            End If
            If first > 0 Then
                If iPath > 0 Then out(i, 3) = CStr(gData(first, iPath))
                If iLink > 0 Then
                    v = CStr(gData(first, iLink))
                    If Left(v, 5) = "pw://" And Len(v) < 250 Then out(i, 4) = "=HYPERLINK(""" & v & """,""" & v & """)" Else out(i, 4) = v
                End If
            End If
        End If
    Next i
    ws.Range("B" & FIRST_ROW).Resize(n, 4).Value = out
    For i = 1 To n
        Select Case CStr(out(i, 2))
            Case "FOUND": ws.Cells(FIRST_ROW + i - 1, 3).Interior.Color = RGB(226, 239, 218)
            Case "MULTIPLE": ws.Cells(FIRST_ROW + i - 1, 3).Interior.Color = RGB(255, 242, 204)
            Case "NOT FOUND": ws.Cells(FIRST_ROW + i - 1, 3).Interior.Color = RGB(255, 199, 206)
            Case Else: ws.Cells(FIRST_ROW + i - 1, 3).Interior.ColorIndex = xlNone
        End Select
    Next i
    Application.ScreenUpdating = True
    Notify "Batch search: " & nFound & " found, " & nMulti & " with several matches, " & nMiss & " not found."
End Sub

' wszystkie trafienia (suma po nazwach) jako wyniki w Dashboard (szukanie "batch:" czyta liste z arkusza Batch)
Public Sub BatchToDashboard()
    Dim wsV As Worksheet
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    EnsureBatchSheet
    wsV.Range(FOLDER_CELL).Value = ""
    wsV.Activate
    If LCase(CStr(wsV.Range(VIEW_SEARCH).Value)) = "batch:" Then
        RenderDashboard True
    Else
        wsV.Range(VIEW_SEARCH).Value = "batch:"
    End If
End Sub

Public Sub ClearBatch()
    Dim ws As Worksheet, last As Long
    Set ws = EnsureBatchSheet()
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last >= FIRST_ROW Then ws.Range("A" & FIRST_ROW & ":E" & last).Clear
    ws.Range("A" & FIRST_ROW & ":A5000").NumberFormat = "@"
    ws.Range("D" & FIRST_ROW & ":D5000").NumberFormat = "@"
End Sub

' wywolywane z RunQuery dla frazy "batch:" - wypelnia gIdx / gHits
Public Function BatchQuery() As Long
    Dim ws As Worksheet, pats() As String, wild() As Boolean, n As Long
    Dim iName As Long, k As Long, i As Long, v As String
    gHits = 0
    iName = ColOf("File Name")
    Set ws = EnsureBatchSheet()
    n = ReadNames(ws, pats, wild)
    If n = 0 Or iName = 0 Then Exit Function
    For k = 1 To gN
        v = LCase(CStr(gData(k, iName)))
        For i = 1 To n
            If pats(i) <> "" Then
                If IsHit(v, pats(i), wild(i)) Then
                    gHits = gHits + 1: gIdx(gHits) = k
                    Exit For
                End If
            End If
        Next i
    Next k
    BatchQuery = gHits
End Function
