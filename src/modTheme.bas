Attribute VB_Name = "modTheme"
Option Explicit

' =======================================================================
' WYGLAD APLIKACJI: dashboard w arkuszu Widok (naglowek, wyszukiwarka, legenda, motywy Light / Dark / Blue)
' Motyw wybiera sie w Widok!H4. Kolory danych sa nakladane po kazdej przebudowie widoku (BuildView).
' =======================================================================
Private Type TTheme
    titleBg As Long
    titleFg As Long
    panelBg As Long
    panelFg As Long
    inputBg As Long
    inputFg As Long
    accent As Long
    hdrBg As Long
    hdrFg As Long
    dataBg As Long
    dataFg As Long
    folderBg As Long
    changedBg As Long
    deletedBg As Long
    cardBg As Long
    cardMuted As Long
End Type

Private Function GetTheme(ByVal nm As String) As TTheme
    Dim t As TTheme
    Select Case LCase(Trim(nm))
        Case "ciemny", "dark"
            t.titleBg = RGB(18, 18, 18): t.titleFg = RGB(240, 240, 240)
            t.panelBg = RGB(37, 37, 38): t.panelFg = RGB(210, 210, 210)
            t.inputBg = RGB(60, 60, 64): t.inputFg = RGB(255, 255, 255): t.accent = RGB(0, 150, 255)
            t.hdrBg = RGB(0, 90, 158): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(30, 30, 30): t.dataFg = RGB(224, 224, 224)
            t.folderBg = CLR_FOLDER_DARK: t.changedBg = RGB(120, 100, 20): t.deletedBg = RGB(120, 40, 48)
            t.cardBg = RGB(52, 52, 56): t.cardMuted = RGB(160, 160, 165)
        Case "niebieski", "blue"
            t.titleBg = RGB(0, 70, 127): t.titleFg = RGB(255, 255, 255)
            t.panelBg = RGB(228, 238, 250): t.panelFg = RGB(0, 56, 104)
            t.inputBg = RGB(255, 255, 255): t.inputFg = RGB(20, 20, 20): t.accent = RGB(0, 92, 170)
            t.hdrBg = RGB(0, 92, 170): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(250, 252, 255): t.dataFg = RGB(30, 30, 30)
            t.folderBg = CLR_FOLDER_BLUE: t.changedBg = CLR_CHANGED: t.deletedBg = CLR_DELETED
            t.cardBg = RGB(255, 255, 255): t.cardMuted = RGB(90, 110, 140)
        Case Else
            t.titleBg = RGB(31, 56, 100): t.titleFg = RGB(255, 255, 255)
            t.panelBg = RGB(242, 246, 252): t.panelFg = RGB(31, 56, 100)
            t.inputBg = RGB(255, 255, 255): t.inputFg = RGB(20, 20, 20): t.accent = RGB(47, 85, 151)
            t.hdrBg = RGB(31, 56, 100): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(255, 255, 255): t.dataFg = RGB(32, 32, 32)
            t.folderBg = CLR_FOLDER: t.changedBg = CLR_CHANGED: t.deletedBg = CLR_DELETED
            t.cardBg = RGB(255, 255, 255): t.cardMuted = RGB(110, 120, 140)
    End Select
    GetTheme = t
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If StrComp(w.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next w
End Function

' Nowe (angielskie) nazwy arkuszy: zmienia nazwy ze starych polskich i naprawia odnosniki wewnetrzne
Public Sub EnsureSheetNames()
    Dim pairs As Variant, i As Long, ws As Worksheet, hl As Hyperlink, sh As Worksheet
    pairs = Array("Instrukcja", SH_START, "Foldery", SH_FOLDERS, "Atrybuty", SH_ATTRS, "Baza_PW", SH_DATA, _
                  "Diagnostyka", SH_DIAG, "Widoki", SH_VIEWS, "Widok", SH_VIEW)
    On Error Resume Next
    For i = 0 To UBound(pairs) Step 2
        Set ws = Nothing
        Set ws = ThisWorkbook.Sheets(CStr(pairs(i)))
        If Not ws Is Nothing Then
            If Not SheetExists(CStr(pairs(i + 1))) Then
                ws.Name = CStr(pairs(i + 1))
                For Each sh In ThisWorkbook.Worksheets          ' odnosniki wewnetrzne (SubAddress) nie aktualizuja sie same
                    For Each hl In sh.Hyperlinks
                        If InStr(hl.SubAddress, CStr(pairs(i))) > 0 Then hl.SubAddress = Replace(hl.SubAddress, CStr(pairs(i)), CStr(pairs(i + 1)))
                    Next hl
                Next sh
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

' Siatka wylaczona, kolumna A (drzewo) i naglowek zablokowane (freeze panes) - tylko gdy Dashboard jest aktywny
Public Sub PrepareWindow()
    On Error Resume Next
    If ActiveSheet.Name <> SH_VIEW Then Exit Sub
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    If ActiveWindow.SplitColumn <> 1 Or ActiveWindow.SplitRow <> VIEW_HEADER_ROW Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
        ActiveSheet.Range("B" & (VIEW_HEADER_ROW + 1)).Select
        ActiveWindow.FreezePanes = True
    End If
End Sub

' Ustawia przyciski w wierszu 6 (jeden na kolumne A, B, C) i nadaje angielskie napisy
Private Sub PlaceButtons(wsV As Worksheet)
    Dim names As Variant, caps As Variant, cols As Variant, i As Long, shp As Shape
    names = Array("btnUpdateV", "btnNavV", "btnGetFiles")
    caps = Array("UPDATE DATA", "FOLDER NAVIGATOR", "DOWNLOAD FROM PW")
    On Error Resume Next
    For i = 0 To 2
        Set shp = Nothing
        Set shp = wsV.Shapes(CStr(names(i)))
        If Not shp Is Nothing Then
            shp.Placement = xlFreeFloating
            shp.Left = wsV.Cells(6, i + 1).Left + 6
            shp.Width = wsV.Cells(6, i + 1).Width - 12
            shp.Top = wsV.Cells(6, 1).Top + 4
            shp.Height = wsV.Cells(6, 1).Height - 8
            shp.TextFrame2.TextRange.Text = CStr(caps(i))
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub Card(wsV As Worksheet, ByVal cols As String, ByVal lbl As String, ByVal valTxt As String, t As TTheme, ByVal accentClr As Long)
    Dim lblR As Range, valR As Range, c1 As String, c2 As String
    c1 = Split(cols & ":" & cols, ":")(0): c2 = Split(cols & ":" & cols, ":")(1)
    Set lblR = wsV.Range(c1 & "2:" & c2 & "2")
    Set valR = wsV.Range(c1 & "3:" & c2 & "3")
    If c1 <> c2 Then lblR.Merge: valR.Merge
    lblR.Value = lbl
    valR.NumberFormat = "@"
    valR.Value = valTxt
    With wsV.Range(c1 & "2:" & c2 & "3")
        .Interior.Color = t.cardBg
        .HorizontalAlignment = xlLeft: .IndentLevel = 1
        .Borders(xlEdgeLeft).Color = t.panelBg: .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeRight).Color = t.panelBg: .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeTop).Color = accentClr: .Borders(xlEdgeTop).Weight = xlThick
    End With
    lblR.Font.Size = 8: lblR.Font.Bold = True: lblR.Font.Color = t.cardMuted
    valR.Font.Size = 20: valR.Font.Bold = True: valR.Font.Color = accentClr
End Sub

' Naklada motyw i uklad dashboardu na arkusz Dashboard; lastR = ostatni wiersz danych, lastC = ostatnia kolumna tabeli
Public Sub StyleDashboard(wsV As Worksheet, ByVal lastR As Long, ByVal lastC As Long, ByVal subtitle As String, ByVal warn As String, ByVal statusTxt As String)
    Dim t As TTheme, nm As String, i As Long, r As Long
    Dim nFold As Long, nDel As Long, nFile As Long, nChg As Long
    Dim selV As String, srchV As String, tgtV As String, thV As String
    On Error Resume Next
    selV = CStr(wsV.Range(VIEW_SELECTOR).Value): srchV = CStr(wsV.Range(VIEW_SEARCH).Value)
    tgtV = CStr(wsV.Range(TARGET_CELL).Value): thV = CStr(wsV.Range(THEME_CELL).Value)
    If Trim(thV) = "" Then thV = "Light"
    nm = thV
    t = GetTheme(nm)
    If lastR < VIEW_HEADER_ROW Then lastR = VIEW_HEADER_ROW
    If lastC < 12 Then lastC = 12
    Application.ScreenUpdating = False
    
    ' --- zmienione komorki zapamietaj przed nadpisaniem koloru ---
    Dim chg As Collection, fc As Range, firstAddr As String, scan As Range
    Set chg = New Collection
    Application.FindFormat.Clear
    Application.FindFormat.Interior.Color = CLR_CHANGED
    If lastR > VIEW_HEADER_ROW Then
        Set scan = wsV.Range(wsV.Cells(VIEW_HEADER_ROW + 1, 2), wsV.Cells(lastR, lastC))
        Set fc = scan.Find(What:="", SearchFormat:=True)
        If Not fc Is Nothing Then
            firstAddr = fc.Address
            Do
                chg.Add fc.Address
                Set fc = scan.FindNext(fc)
            Loop While Not fc Is Nothing And chg.Count < 5000 And fc.Address <> firstAddr
        End If
    End If
    Application.FindFormat.Clear
    
    ' --- tlo za tabela i dane ---
    wsV.Range(wsV.Columns(lastC + 1), wsV.Columns(40)).Interior.Color = t.dataBg
    wsV.Range(wsV.Rows(lastR + 1), wsV.Rows(lastR + 200)).Interior.Color = t.dataBg
    If lastR > VIEW_HEADER_ROW Then
        With wsV.Range(wsV.Cells(VIEW_HEADER_ROW + 1, 1), wsV.Cells(lastR, lastC))
            .Interior.Color = t.dataBg
            .Font.Color = t.dataFg
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .VerticalAlignment = xlCenter
            .Borders(xlInsideHorizontal).Color = RGB(205, 210, 220)
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        For r = 1 To gRowCount
            Select Case gRowType(r)
                Case 1
                    nFold = nFold + 1
                    With wsV.Range(wsV.Cells(VIEW_HEADER_ROW + r, 1), wsV.Cells(VIEW_HEADER_ROW + r, lastC))
                        .Interior.Color = t.folderBg: .Font.Bold = True
                    End With
                Case 2
                    nDel = nDel + 1
                    wsV.Range(wsV.Cells(VIEW_HEADER_ROW + r, 1), wsV.Cells(VIEW_HEADER_ROW + r, lastC)).Interior.Color = t.deletedBg
                Case Else
                    nFile = nFile + 1
            End Select
        Next r
        For i = 1 To chg.Count
            wsV.Range(chg(i)).Interior.Color = t.changedBg
        Next i
        nChg = chg.Count
    End If
    
    ' --- naglowek tabeli ---
    With wsV.Range(wsV.Cells(VIEW_HEADER_ROW, 1), wsV.Cells(VIEW_HEADER_ROW, lastC))
        .Interior.Color = t.hdrBg: .Font.Color = t.hdrFg: .Font.Bold = True: .Font.Name = "Segoe UI": .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    wsV.Rows(VIEW_HEADER_ROW).RowHeight = 28
    wsV.Cells(VIEW_HEADER_ROW, 1).Value = "FOLDER TREE  (click here = show all)"
    
    ' --- panel gorny: wiersze 1-7 budowane od zera ---
    With wsV.Range(wsV.Cells(1, 1), wsV.Cells(7, 40))
        .UnMerge
        .Validation.Delete
        .Clear
        .Interior.Color = t.panelBg: .Font.Color = t.panelFg: .Font.Name = "Segoe UI": .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    wsV.Rows(1).RowHeight = 40: wsV.Rows(2).RowHeight = 18: wsV.Rows(3).RowHeight = 34: wsV.Rows(4).RowHeight = 8
    wsV.Rows(5).RowHeight = 32: wsV.Rows(6).RowHeight = 40: wsV.Rows(7).RowHeight = 26
    With wsV.Range(wsV.Cells(1, 1), wsV.Cells(1, 40))
        .Interior.Color = t.titleBg: .Font.Color = t.titleFg
    End With
    With wsV.Range("A1"): .Value = "PW TRACKER": .Font.Size = 22: .Font.Bold = True: .IndentLevel = 1: End With
    wsV.Range("A2").Value = subtitle: wsV.Range("A2").Font.Size = 9: wsV.Range("A2").Font.Color = t.cardMuted: wsV.Range("A2").IndentLevel = 1
    wsV.Range("A3").Value = warn: wsV.Range("A3").Font.Size = 9: wsV.Range("A3").Font.Color = RGB(192, 0, 0): wsV.Range("A3").WrapText = True: wsV.Range("A3").IndentLevel = 1
    
    ' --- karty KPI ---
    Card wsV, "B", "FOLDERS", CStr(nFold), t, RGB(47, 117, 181)
    Card wsV, "C", "FILES", CStr(nFile), t, RGB(0, 150, 136)
    Card wsV, "D", "CHANGED CELLS", CStr(nChg), t, RGB(237, 125, 49)
    Card wsV, "E:F", "DELETED FROM PW", CStr(nDel), t, RGB(214, 69, 65)
    Card wsV, "G:I", "LAST UPDATE", Replace(statusTxt, "Last Database Update: ", ""), t, t.accent
    wsV.Range("G3").Font.Size = 11
    
    ' --- sterowanie: widok, wyszukiwarka, motyw, folder docelowy ---
    Dim lbl As Variant
    For Each lbl In Array("A5", "C5", "E5", "G5")
        With wsV.Range(CStr(lbl)): .Font.Bold = True: .Font.Size = 9: .HorizontalAlignment = xlRight: .Font.Color = t.cardMuted: End With
    Next lbl
    wsV.Range("A5").Value = "VIEW": wsV.Range("C5").Value = "SEARCH": wsV.Range("E5").Value = "THEME": wsV.Range("G5").Value = "DOWNLOAD TO"
    wsV.Range(VIEW_SELECTOR).Value = selV: wsV.Range(VIEW_SEARCH).Value = srchV
    wsV.Range(TARGET_CELL).Value = tgtV: wsV.Range(THEME_CELL).Value = thV
    wsV.Range("H5:J5").Merge
    Dim fld As Variant
    For Each fld In Array(VIEW_SELECTOR, VIEW_SEARCH, THEME_CELL, "H5:J5")
        With wsV.Range(CStr(fld))
            .Interior.Color = t.inputBg: .Font.Color = t.inputFg: .Font.Bold = True
            .Borders.LineStyle = xlContinuous: .Borders.Color = t.accent: .Borders.Weight = xlThin
            .HorizontalAlignment = xlLeft: .IndentLevel = 1
        End With
    Next fld
    With wsV.Range(VIEW_SEARCH)                              ' wyszukiwarka: wyrazna, gruba obwodka i wieksza czcionka
        .Borders.Weight = xlThick: .Borders.Color = t.accent: .Font.Size = 12
    End With
    With wsV.Range(VIEW_SELECTOR).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & SH_VIEWS & "!$A$" & VIEWS_FIRST_ROW & ":$A$25"
    End With
    With wsV.Range(THEME_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Light,Dark,Blue"
    End With
    
    ' --- legenda kolorow ---
    wsV.Range("A7").Value = "COLOR LEGEND"
    With wsV.Range("A7"): .Font.Bold = True: .Font.Size = 9: .HorizontalAlignment = xlRight: .Font.Color = t.cardMuted: End With
    Dim chipTxt As Variant, chipBg As Variant, chipAt As Variant, k As Long
    chipTxt = Array("Folder", "Changed value", "Deleted from PW", "Regular file")
    chipBg = Array(t.folderBg, t.changedBg, t.deletedBg, t.dataBg)
    chipAt = Array("B7", "C7", "D7", "E7:F7")
    For k = 0 To 3
        With wsV.Range(CStr(chipAt(k)))
            If k = 3 Then .Merge
            .Value = chipTxt(k)
            .Interior.Color = chipBg(k): .Font.Color = t.dataFg
            .HorizontalAlignment = xlCenter: .Font.Bold = True: .Font.Size = 9
            .Borders.LineStyle = xlContinuous: .Borders.Color = t.accent: .Borders.Weight = xlThin
        End With
    Next k
    With wsV.Range("G7"): .Value = "Click a folder in column A to filter. Click the header (A8) to show all.": .Font.Italic = True: .Font.Size = 9: .Font.Color = t.cardMuted: .IndentLevel = 1: End With
    
    ' --- linia pod panelem ---
    With wsV.Range(wsV.Cells(7, 1), wsV.Cells(7, 40)).Borders(xlEdgeBottom)
        .Color = t.accent: .Weight = xlMedium
    End With
    PlaceButtons wsV
    PrepareWindow
    Application.ScreenUpdating = True
End Sub
