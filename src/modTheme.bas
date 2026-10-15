Attribute VB_Name = "modTheme"
Option Explicit

' =======================================================================
' WYGLAD APLIKACJI: dashboard w arkuszu Widok (naglowek, wyszukiwarka, legenda, motywy Light / Dark / Blue)
' Motyw wybiera sie w Widok!H4. Kolory danych sa nakladane po kazdej przebudowie widoku (BuildView).
' =======================================================================
Private mFold As Long, mFile As Long, mChg As Long, mDel As Long
Private mSub As String, mWarn As String, mStatus As String

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

' Siatka wylaczona, kolumny A:B (nawigator/drzewo) i naglowek zablokowane (freeze panes) - tylko gdy Dashboard jest aktywny
Public Sub PrepareWindow()
    On Error Resume Next
    If ActiveSheet.Name <> SH_VIEW Then Exit Sub
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    If ActiveWindow.SplitColumn <> 2 Or ActiveWindow.SplitRow <> VIEW_HEADER_ROW Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
        ActiveSheet.Range("C" & (VIEW_HEADER_ROW + 1)).Select
        ActiveWindow.FreezePanes = True
    End If
End Sub

' Naklada motyw i uklad dashboardu na arkusz Dashboard; lastR = ostatni wiersz danych, lastC = ostatnia kolumna tabeli
Public Sub StyleDashboard(wsV As Worksheet, ByVal lastR As Long, ByVal lastC As Long, ByVal subtitle As String, ByVal warn As String, ByVal statusTxt As String)
    Dim t As TTheme, nm As String, i As Long, r As Long
    Dim nFold As Long, nDel As Long, nFile As Long, nChg As Long
    Dim selV As String, srchV As String, tgtV As String, thV As String
    On Error Resume Next
    selV = CStr(wsV.Range(VIEW_SELECTOR).Value): srchV = CStr(wsV.Range(VIEW_SEARCH).Value)
    tgtV = CStr(wsV.Range(TARGET_CELL).Value): thV = CStr(wsV.Range(THEME_CELL).Value)
    If Trim(thV) = "" Then thV = "Light": wsV.Range(THEME_CELL).Value = thV
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
    wsV.Cells.Interior.Color = t.dataBg                      ' cale tlo arkusza (takze poza tabela) w kolorze motywu
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
    wsV.Range(wsV.Cells(VIEW_HEADER_ROW, 1), wsV.Cells(VIEW_HEADER_ROW, lastC)).UnMerge
    If lastR <= VIEW_HEADER_ROW Then wsV.Range(wsV.Cells(VIEW_HEADER_ROW, 2), wsV.Cells(VIEW_HEADER_ROW, lastC)).ClearContents
    With wsV.Range(wsV.Cells(VIEW_HEADER_ROW, 1), wsV.Cells(VIEW_HEADER_ROW, lastC))
        .Interior.Color = t.hdrBg: .Font.Color = t.hdrFg: .Font.Bold = True: .Font.Name = "Segoe UI": .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    wsV.Rows(VIEW_HEADER_ROW).RowHeight = 28
    wsV.Cells(VIEW_HEADER_ROW, 1).Value = "FOLDER TREE  (click here = show all)"
    
    ' --- panel gorny: wiersze 1-7 to tylko tlo; elementy (karty, pola, przyciski) to ksztalty niezalezne od szerokosci kolumn ---
    With wsV.Range(wsV.Cells(1, 1), wsV.Cells(7, 40))
        .UnMerge
        .Validation.Delete
        .Clear
        .Interior.Color = t.panelBg: .Font.Name = "Segoe UI": .Font.Size = 10
    End With
    wsV.Range(wsV.Cells(1, 1), wsV.Cells(1, 40)).Interior.Color = t.titleBg
    wsV.Rows(1).RowHeight = 40: wsV.Rows(2).RowHeight = 62: wsV.Rows(3).RowHeight = 38: wsV.Rows(4).RowHeight = 42
    wsV.Rows(5).RowHeight = 28: wsV.Rows(6).RowHeight = 4: wsV.Rows(7).RowHeight = 4
    With wsV.Range(wsV.Cells(7, 1), wsV.Cells(7, 40)).Borders(xlEdgeBottom)
        .Color = t.accent: .Weight = xlMedium
    End With
    ' stan interfejsu przechowywany w ukrytych komorkach (kolumny AA:AB)
    wsV.Range(VIEW_SELECTOR).Value = selV: wsV.Range(VIEW_SEARCH).Value = srchV
    wsV.Range(TARGET_CELL).Value = tgtV: wsV.Range(THEME_CELL).Value = thV
    wsV.Range("AA:AB").Font.Color = t.panelBg
    wsV.Columns("AA:AB").Hidden = True
    
    mFold = nFold: mFile = nFile: mChg = nChg: mDel = nDel
    mSub = subtitle: mWarn = warn: mStatus = Replace(statusTxt, "Last Database Update: ", "")
    DrawHeader wsV
    PrepareWindow
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = "frmNavigator" Then uf.ApplyNavTheme
    Next uf
    Application.ScreenUpdating = True
End Sub

Private Function UiShape(wsV As Worksheet, ByVal nm As String, ByVal kind As Long, ByVal l As Double, ByVal tp As Double, ByVal w As Double, ByVal h As Double, _
                         ByVal fillClr As Long, ByVal lineClr As Long, ByVal txt As String, ByVal fgClr As Long, ByVal fsz As Long, ByVal macro As String, ByVal alignCenter As Boolean) As Shape
    Dim shp As Shape
    On Error Resume Next
    Set shp = wsV.Shapes(nm)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = wsV.Shapes.AddShape(kind, l, tp, w, h)
        shp.Name = nm
    End If
    shp.Placement = xlFreeFloating
    shp.Left = l: shp.Top = tp: shp.Width = w: shp.Height = h
    shp.Fill.ForeColor.RGB = fillClr
    If lineClr < 0 Then
        shp.Line.Visible = msoFalse
    Else
        shp.Line.Visible = msoTrue: shp.Line.ForeColor.RGB = lineClr: shp.Line.Weight = 1
    End If
    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 10: .MarginRight = 8: .MarginTop = 2: .MarginBottom = 2
        .WordWrap = msoFalse
        .TextRange.Text = txt
        .TextRange.Font.Fill.ForeColor.RGB = fgClr
        .TextRange.Font.Size = fsz
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.ParagraphFormat.Alignment = IIf(alignCenter, msoAlignCenter, msoAlignLeft)
    End With
    shp.OnAction = macro
    Set UiShape = shp
End Function

' Rysuje caly gorny panel z ksztaltow (pozycje bezwzgledne - zmiana szerokosci kolumn niczego nie psuje)
Public Sub DrawHeader(wsV As Worksheet)
    Dim t As TTheme, x As Double, y As Double, shp As Shape, k As Long
    Dim vw As String, sr As String, th As String, tg As String
    On Error Resume Next
    th = CStr(wsV.Range(THEME_CELL).Value): If Trim(th) = "" Then th = "Light"
    t = GetTheme(th)
    vw = CStr(wsV.Range(VIEW_SELECTOR).Value): sr = CStr(wsV.Range(VIEW_SEARCH).Value): tg = CStr(wsV.Range(TARGET_CELL).Value)
    Dim wn As String: wn = "'" & ThisWorkbook.Name & "'!"
    
    ' tytul + info
    Set shp = UiShape(wsV, "ui_title", msoShapeRectangle, 0, 0, 4000, wsV.Rows(2).Top, t.titleBg, -1, "PW TRACKER", t.titleFg, 22, "", False)
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    Set shp = UiShape(wsV, "ui_info", msoShapeRectangle, 250, 6, 900, 28, t.titleBg, -1, mSub & IIf(mWarn <> "", "   |   " & mWarn, ""), t.cardMuted, 9, "", False)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(mWarn <> "", RGB(255, 140, 130), RGB(170, 185, 210))
    
    ' karty KPI
    Dim caps As Variant, vals As Variant, clrs As Variant, wds As Variant
    caps = Array("FOLDERS", "FILES", "CHANGED CELLS", "DELETED FROM PW", "LAST UPDATE")
    vals = Array(CStr(mFold), CStr(mFile), CStr(mChg), CStr(mDel), mStatus)
    clrs = Array(RGB(47, 117, 181), RGB(0, 150, 136), RGB(237, 125, 49), RGB(214, 69, 65), t.accent)
    wds = Array(130, 130, 150, 160, 300)
    x = 10: y = wsV.Rows(2).Top + 6
    For k = 0 To 4
        Set shp = UiShape(wsV, "ui_card" & k, msoShapeRectangle, x, y, CDbl(wds(k)), 50, t.cardBg, RGB(205, 210, 220), CStr(caps(k)) & vbCr & CStr(vals(k)), t.cardMuted, 8, "", False)
        shp.Line.ForeColor.RGB = t.panelBg
        With shp.TextFrame2.TextRange
            .Paragraphs(1).Font.Size = 8: .Paragraphs(1).Font.Bold = msoTrue
            .Paragraphs(2).Font.Size = IIf(k = 4, 12, 20): .Paragraphs(2).Font.Bold = msoTrue
            .Paragraphs(2).Font.Fill.ForeColor.RGB = clrs(k)
        End With
        shp.Line.Visible = msoFalse
        x = x + wds(k) + 8
    Next k
    
    ' pola sterowania: widok, szukaj, motyw, folder docelowy (klik = menu / okno)
    y = wsV.Rows(3).Top + 4: x = 10
    Set shp = UiShape(wsV, "ui_view", msoShapeRoundedRectangle, x, y, 180, 30, t.inputBg, t.accent, "VIEW:  " & vw & "   " & ChrW(9660), t.inputFg, 10, wn & "ShowViewMenu", False): x = x + 188
    Set shp = UiShape(wsV, "ui_search", msoShapeRoundedRectangle, x, y, 290, 30, t.inputBg, t.accent, "SEARCH:  " & IIf(sr = "", "click to type...", sr), t.inputFg, 11, wn & "EditSearch", False)
    shp.Line.Weight = 2.5: x = x + 298
    Set shp = UiShape(wsV, "ui_theme", msoShapeRoundedRectangle, x, y, 140, 30, t.inputBg, t.accent, "THEME:  " & th & "   " & ChrW(9660), t.inputFg, 10, wn & "ShowThemeMenu", False): x = x + 148
    Set shp = UiShape(wsV, "ui_target", msoShapeRoundedRectangle, x, y, 330, 30, t.inputBg, t.accent, "DOWNLOAD TO:  " & IIf(tg = "", "click to choose...", tg), t.inputFg, 10, wn & "ChooseTargetFolder", False)
    
    ' przyciski akcji
    y = wsV.Rows(4).Top + 5: x = 10
    Dim nms As Variant, cps As Variant, wd As Variant
    nms = Array("btnUpdateV", "btnNavV", "btnGetFiles")
    cps = Array("UPDATE DATA", "FOLDER NAVIGATOR", "DOWNLOAD FROM PW")
    wd = Array(180, 180, 190)
    For k = 0 To 2
        Set shp = Nothing
        Set shp = wsV.Shapes(CStr(nms(k)))
        If Not shp Is Nothing Then
            shp.Placement = xlFreeFloating
            shp.Left = x: shp.Top = y: shp.Width = wd(k): shp.Height = 32
            shp.TextFrame2.TextRange.Text = CStr(cps(k))
            shp.ZOrder msoBringToFront
        End If
        x = x + wd(k) + 8
    Next k
    
    ' legenda
    y = wsV.Rows(5).Top + 4: x = 10
    Set shp = UiShape(wsV, "ui_legend", msoShapeRectangle, x, y, 110, 20, t.panelBg, -1, "COLOR LEGEND", t.cardMuted, 8, "", False)
    shp.TextFrame2.TextRange.Font.Bold = msoTrue: x = x + 112
    Dim ct As Variant, cb As Variant
    ct = Array("Folder", "Changed value", "Deleted from PW", "Regular file")
    cb = Array(t.folderBg, t.changedBg, t.deletedBg, t.dataBg)
    For k = 0 To 3
        Set shp = UiShape(wsV, "ui_chip" & k, msoShapeRectangle, x, y, 130, 20, CLng(cb(k)), t.accent, CStr(ct(k)), t.dataFg, 9, "", True)
        x = x + 136
    Next k
    Set shp = UiShape(wsV, "ui_hint", msoShapeRectangle, x + 10, y, 520, 20, t.panelBg, -1, "Click a folder in column A to filter. Click the header (A8) to show all.", t.cardMuted, 8, "", False)
    shp.TextFrame2.TextRange.Font.Italic = msoTrue
    On Error GoTo 0
End Sub

' --- menu rozwijane: widok / motyw ---
Public Sub ShowViewMenu()
    Dim cb As CommandBar, wsDef As Worksheet, r As Long, ctl As CommandBarButton, lastDef As Long, cur As String
    Set wsDef = ThisWorkbook.Sheets(SH_VIEWS)
    cur = CStr(ThisWorkbook.Sheets(SH_VIEW).Range(VIEW_SELECTOR).Value)
    On Error Resume Next
    Application.CommandBars("PWViewMenu").Delete
    On Error GoTo 0
    Set cb = Application.CommandBars.Add("PWViewMenu", msoBarPopup, False, True)
    lastDef = wsDef.Cells(wsDef.Rows.Count, 1).End(xlUp).Row
    For r = VIEWS_FIRST_ROW To lastDef
        If Trim(CStr(wsDef.Cells(r, 1).Value)) <> "" Then
            Set ctl = cb.Controls.Add(msoControlButton)
            ctl.Caption = CStr(wsDef.Cells(r, 1).Value)
            ctl.OnAction = "'" & ThisWorkbook.Name & "'!PickView"
            ctl.Parameter = CStr(wsDef.Cells(r, 1).Value)
            If StrComp(ctl.Caption, cur, vbTextCompare) = 0 Then ctl.State = msoButtonDown
        End If
    Next r
    cb.ShowPopup
End Sub

Public Sub PickView()
    ThisWorkbook.Sheets(SH_VIEW).Range(VIEW_SELECTOR).Value = CStr(Application.CommandBars.ActionControl.Parameter)   ' zmiana komorki uruchamia BuildView
End Sub

Public Sub ShowThemeMenu()
    Dim cb As CommandBar, ctl As CommandBarButton, nm As Variant
    On Error Resume Next
    Application.CommandBars("PWThemeMenu").Delete
    On Error GoTo 0
    Set cb = Application.CommandBars.Add("PWThemeMenu", msoBarPopup, False, True)
    For Each nm In Array("Light", "Dark", "Blue")
        Set ctl = cb.Controls.Add(msoControlButton)
        ctl.Caption = CStr(nm)
        ctl.OnAction = "'" & ThisWorkbook.Name & "'!PickTheme"
        ctl.Parameter = CStr(nm)
    Next nm
    cb.ShowPopup
End Sub

Public Sub PickTheme()
    ThisWorkbook.Sheets(SH_VIEW).Range(THEME_CELL).Value = CStr(Application.CommandBars.ActionControl.Parameter)
End Sub

' --- szukaj / folder docelowy ---
Public Sub EditSearch()
    Dim ws As Worksheet, v As String
    Set ws = ThisWorkbook.Sheets(SH_VIEW)
    v = InputBox("Search phrase (min. 2 characters; leave empty to clear):", "Search", CStr(ws.Range(VIEW_SEARCH).Value))
    If StrPtr(v) = 0 Then Exit Sub                  ' Cancel
    ws.Range(VIEW_SEARCH).Value = v                 ' zmiana komorki uruchamia filtr
    DrawHeader ws
End Sub

Public Sub ChooseTargetFolder()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_VIEW)
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choose the target folder for downloaded files"
        If .Show = -1 Then ws.Range(TARGET_CELL).Value = .SelectedItems(1)
    End With
    DrawHeader ws
End Sub

' Przebudowuje nawigator od nowa (drzewo z aktualnej bazy) - po aktualizacji danych
Public Sub RefreshNavigator()
    Dim uf As Object
    On Error Resume Next
    For Each uf In VBA.UserForms
        If uf.Name = "frmNavigator" Then Unload uf: Exit For
    Next uf
    If ActiveSheet.Name = SH_VIEW Then
        frmNavigator.Show 0
        frmNavigator.DockToView
    End If
End Sub
