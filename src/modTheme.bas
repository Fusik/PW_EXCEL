Attribute VB_Name = "modTheme"
Option Explicit
Public mFzUsed As Double                             ' szerokosc zamrozonych kolumn uzyta przy ostatnim rysowaniu panelu

' =======================================================================
' WYGLAD APLIKACJI: dashboard w arkuszu Widok (naglowek, wyszukiwarka, legenda, motywy Light / Dark / Blue)
' Motyw wybiera sie w Widok!H4. Kolory danych sa nakladane po kazdej przebudowie widoku (BuildView).
' =======================================================================
Private mFold As Long, mFile As Long, mChg As Long, mDel As Long
Private mSub As String, mWarn As String, mStatus As String
Private mPage As Long, mPages As Long, mTotal As Long, mFolder As String
Private mDrawnTheme As String

Public Type TTheme
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
    newBg As Long
    cardBg As Long
    cardMuted As Long
End Type

Public Function GetTheme(ByVal nm As String) As TTheme
    Dim t As TTheme
    Select Case LCase(Trim(nm))
        Case "ciemny", "dark"
            t.titleBg = RGB(18, 18, 18): t.titleFg = RGB(240, 240, 240)
            t.panelBg = RGB(37, 37, 38): t.panelFg = RGB(210, 210, 210)
            t.inputBg = RGB(60, 60, 64): t.inputFg = RGB(255, 255, 255): t.accent = RGB(0, 150, 255)
            t.hdrBg = RGB(0, 90, 158): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(30, 30, 30): t.dataFg = RGB(224, 224, 224)
            t.folderBg = CLR_FOLDER_DARK: t.changedBg = RGB(120, 100, 20): t.deletedBg = RGB(120, 40, 48)
            t.cardBg = RGB(52, 52, 56): t.cardMuted = RGB(160, 160, 165): t.newBg = RGB(40, 85, 55)
        Case "niebieski", "blue"
            t.titleBg = RGB(0, 70, 127): t.titleFg = RGB(255, 255, 255)
            t.panelBg = RGB(228, 238, 250): t.panelFg = RGB(0, 56, 104)
            t.inputBg = RGB(255, 255, 255): t.inputFg = RGB(20, 20, 20): t.accent = RGB(0, 92, 170)
            t.hdrBg = RGB(0, 92, 170): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(250, 252, 255): t.dataFg = RGB(30, 30, 30)
            t.folderBg = CLR_FOLDER_BLUE: t.changedBg = CLR_CHANGED: t.deletedBg = CLR_DELETED
            t.cardBg = RGB(255, 255, 255): t.cardMuted = RGB(90, 110, 140): t.newBg = CLR_NEW
        Case Else
            t.titleBg = RGB(31, 56, 100): t.titleFg = RGB(255, 255, 255)
            t.panelBg = RGB(242, 246, 252): t.panelFg = RGB(31, 56, 100)
            t.inputBg = RGB(255, 255, 255): t.inputFg = RGB(20, 20, 20): t.accent = RGB(47, 85, 151)
            t.hdrBg = RGB(31, 56, 100): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(255, 255, 255): t.dataFg = RGB(32, 32, 32)
            t.folderBg = CLR_FOLDER: t.changedBg = CLR_CHANGED: t.deletedBg = CLR_DELETED
            t.cardBg = RGB(255, 255, 255): t.cardMuted = RGB(110, 120, 140): t.newBg = CLR_NEW
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

' Siatka wylaczona, naglowki wierszy/kolumn wlaczone, kolumny A:B (nawigator/drzewo) i naglowek zablokowane (freeze panes) - tylko gdy Dashboard jest aktywny
Public Sub PrepareWindow()
    On Error Resume Next
    If ActiveSheet.Name <> SH_VIEW Then Exit Sub
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    If ActiveWindow.SplitColumn <> 2 Or ActiveWindow.SplitRow <> VIEW_HEADER_ROW Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
        ActiveSheet.Range("C" & (VIEW_HEADER_ROW + 1)).Select
        ActiveWindow.FreezePanes = True
    End If
End Sub

Public Function CurrentTheme(wsV As Worksheet) As TTheme
    Dim th As String
    th = Trim(CStr(wsV.Range(THEME_CELL).Value))
    If th = "" Then th = "Light"
    CurrentTheme = GetTheme(th)
End Function

Public Sub SetPaging(ByVal page As Long, ByVal pages As Long, ByVal total As Long, ByVal folder As String)
    mPage = page: mPages = pages: mTotal = total: mFolder = folder
End Sub

' Odswieza gorny panel: pelne przerysowanie tylko przy zmianie motywu / pierwszym uruchomieniu, inaczej tylko teksty kart i przyciskow
Public Sub RefreshChrome(wsV As Worksheet, ByVal subtitle As String, ByVal warn As String, ByVal statusTxt As String)
    Dim th As String, hasUi As Boolean, shp As Shape
    th = Trim(CStr(wsV.Range(THEME_CELL).Value))
    If th = "" Then th = "Light": wsV.Range(THEME_CELL).Value = th
    On Error Resume Next
    Set shp = wsV.Shapes("ui_title")
    On Error GoTo 0
    hasUi = Not shp Is Nothing
    If Not hasUi Or StrComp(th, mDrawnTheme, vbTextCompare) <> 0 Then
        StyleDashboard wsV, subtitle, warn, statusTxt
    Else
        DataStats mFold, mFile, mChg, mDel
        mSub = subtitle: mWarn = warn: mStatus = Replace(statusTxt, "Last Database Update: ", "")
        DrawHeader wsV
        Dim uf As Object
        For Each uf In VBA.UserForms
            If uf.Name = "frmNavigator" Then uf.ApplyNavTheme
        Next uf
    End If
End Sub

' Pelne przerysowanie panelu i tla arkusza dla biezacego motywu
Public Sub StyleDashboard(wsV As Worksheet, ByVal subtitle As String, ByVal warn As String, ByVal statusTxt As String)
    Dim t As TTheme
    Dim selV As String, srchV As String, tgtV As String, thV As String, fldV As String, pgV As String
    On Error Resume Next
    selV = CStr(wsV.Range(VIEW_SELECTOR).Value): srchV = CStr(wsV.Range(VIEW_SEARCH).Value)
    tgtV = CStr(wsV.Range(TARGET_CELL).Value): thV = CStr(wsV.Range(THEME_CELL).Value)
    fldV = CStr(wsV.Range(FOLDER_CELL).Value): pgV = CStr(wsV.Range(PAGE_CELL).Value)
    If Trim(thV) = "" Then thV = "Light"
    t = GetTheme(thV)
    Application.ScreenUpdating = False
    wsV.Cells.Interior.Color = t.dataBg                      ' cale tlo arkusza w kolorze motywu
    
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
    ' stan interfejsu w ukrytych komorkach (kolumny AA:AB)
    wsV.Range(VIEW_SELECTOR).Value = selV: wsV.Range(VIEW_SEARCH).Value = srchV
    wsV.Range(TARGET_CELL).Value = tgtV: wsV.Range(THEME_CELL).Value = thV
    wsV.Range(FOLDER_CELL).Value = fldV: wsV.Range(PAGE_CELL).Value = pgV
    wsV.Range("AA:AB").Font.Color = t.panelBg
    wsV.Columns("AA:AB").Hidden = True
    
    DataStats mFold, mFile, mChg, mDel
    mSub = subtitle: mWarn = warn: mStatus = Replace(statusTxt, "Last Database Update: ", "")
    mDrawnTheme = thV
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

' Zamrozone kolumny A:B musza miec stala, wystarczajaca szerokosc (rowna niezaleznie od danych) - inaczej panel sie rozjezdza
Public Sub EnsureFrozenWidth(wsV As Worksheet)
    Const MIN_FROZEN As Double = 548            ' pt (ok. 730 px)
    Dim need As Double
    If wsV.Columns(3).Left < MIN_FROZEN - 1 Then
        need = MIN_FROZEN - wsV.Columns(2).Left
        If need > 20 And wsV.Columns(2).Width > 0 Then wsV.Columns(2).ColumnWidth = wsV.Columns(2).ColumnWidth * need / wsV.Columns(2).Width
    End If
End Sub

' Jesli element zaczynalby sie przed granica zamrozonych kolumn, a konczyl za nia - przesun go za granice
Private Function SnapX(ByVal x As Double, ByVal w As Double, ByVal fz As Double) As Double
    If x < fz And x + w > fz - 2 Then SnapX = fz + 8 Else SnapX = x
End Function

' Rysuje caly gorny panel z ksztaltow (pozycje bezwzgledne - zmiana szerokosci kolumn niczego nie psuje)
Public Sub DrawHeader(wsV As Worksheet)
    Dim t As TTheme, x As Double, y As Double, shp As Shape, k As Long
    Dim vw As String, sr As String, th As String, tg As String
    EnsureFrozenWidth wsV
    Dim fz As Double: fz = wsV.Columns(3).Left: mFzUsed = fz   ' granica zamrozonych kolumn A:B - zaden element nie moze jej przecinac
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
    
    ' uklad: zamrozone kolumny A:B (szerokosc fz) zawieraja KPI x4, widok/szukaj/motyw, 3 przyciski akcji i legende; reszta (ostatnia aktualizacja, folder docelowy, skan atrybutow, log, strony) jest za granica
    Dim iw As Double, xr As Double, gap As Double, cw As Double
    iw = fz - 20: gap = 8: xr = fz + 12
    Dim caps As Variant, vals As Variant, clrs As Variant
    caps = Array("FOLDERS", "FILES", "NEW / CHANGED", "DELETED FROM PW", "LAST UPDATE")
    vals = Array(CStr(mFold), CStr(mFile), CStr(mChg), CStr(mDel), mStatus)
    clrs = Array(RGB(47, 117, 181), RGB(0, 150, 136), RGB(237, 125, 49), RGB(214, 69, 65), t.accent)
    cw = (iw - 3 * gap) / 4
    y = wsV.Rows(2).Top + 6
    For k = 0 To 4
        If k < 4 Then x = 10 + k * (cw + gap) Else x = xr
        Set shp = UiShape(wsV, "ui_card" & k, msoShapeRectangle, x, y, IIf(k < 4, cw, 300), 50, t.cardBg, RGB(205, 210, 220), CStr(caps(k)) & vbCr & CStr(vals(k)), t.cardMuted, 8, "", False)
        With shp.TextFrame2.TextRange
            .Paragraphs(1).Font.Size = 8: .Paragraphs(1).Font.Bold = msoTrue
            .Paragraphs(2).Font.Size = IIf(k = 4, 12, 18): .Paragraphs(2).Font.Bold = msoTrue
            .Paragraphs(2).Font.Fill.ForeColor.RGB = clrs(k)
        End With
        shp.Line.Visible = msoFalse
    Next k
    
    ' pola sterowania: widok, szukaj, motyw (w A:B); folder docelowy za granica
    Dim w1 As Double, w2 As Double, w3 As Double
    w1 = Int(iw * 0.28): w3 = Int(iw * 0.24): w2 = iw - w1 - w3 - 2 * gap
    y = wsV.Rows(3).Top + 4: x = 10
    Set shp = UiShape(wsV, "ui_view", msoShapeRoundedRectangle, x, y, w1, 30, t.inputBg, t.accent, "VIEW:  " & vw & "  " & ChrW(9660), t.inputFg, 10, wn & "ShowViewMenu", False): x = x + w1 + gap
    Set shp = UiShape(wsV, "ui_search", msoShapeRoundedRectangle, x, y, w2, 30, t.inputBg, t.accent, "SEARCH:  " & IIf(sr = "", "click to type...", sr), t.inputFg, 11, wn & "EditSearch", False)
    shp.Line.Weight = 2.5: x = x + w2 + gap
    Set shp = UiShape(wsV, "ui_theme", msoShapeRoundedRectangle, x, y, w3, 30, t.inputBg, t.accent, "THEME:  " & th & "  " & ChrW(9660), t.inputFg, 10, wn & "ShowThemeMenu", False)
    Set shp = UiShape(wsV, "ui_target", msoShapeRoundedRectangle, xr, y, 330, 30, t.inputBg, t.accent, "DOWNLOAD TO:  " & IIf(tg = "", "click to choose...", tg), t.inputFg, 10, wn & "ChooseTargetFolder", False)
    
    ' przyciski akcji: 3 w A:B, skan i log za granica
    y = wsV.Rows(4).Top + 5: x = 10
    Dim nms As Variant, cps As Variant, bw As Double
    nms = Array("btnUpdateV", "btnNavV", "btnGetFiles")
    cps = Array("UPDATE DATA", "FOLDER NAVIGATOR", "DOWNLOAD FROM PW")
    bw = (iw - 2 * gap) / 3
    For k = 0 To 2
        Set shp = Nothing
        Set shp = wsV.Shapes(CStr(nms(k)))
        If Not shp Is Nothing Then
            shp.Placement = xlFreeFloating
            shp.Left = x: shp.Top = y: shp.Width = bw: shp.Height = 32
            shp.TextFrame2.TextRange.Text = CStr(cps(k))
            shp.ZOrder msoBringToFront
        End If
        x = x + bw + gap
    Next k
    Set shp = UiShape(wsV, "ui_scan", msoShapeRoundedRectangle, xr, y, 190, 32, t.inputBg, t.accent, "SCAN ATTRIBUTES (background)", t.inputFg, 9, wn & "StartAttrScan", False)
    Set shp = UiShape(wsV, "ui_pslog", msoShapeRoundedRectangle, xr + 198, y, 130, 32, t.inputBg, t.accent, "OPEN PS LOG", t.inputFg, 9, wn & "OpenLastPsLog", False)
    Set shp = UiShape(wsV, "ui_batch", msoShapeRoundedRectangle, xr + 336, y, 130, 32, t.inputBg, t.accent, "BATCH SEARCH", t.inputFg, 9, wn & "OpenBatchSheet", False)
    
    ' legenda (A:B) + stronicowanie (za granica)
    y = wsV.Rows(5).Top + 4: x = 10
    Dim lw As Double: lw = 84
    Set shp = UiShape(wsV, "ui_legend", msoShapeRectangle, x, y, lw, 20, t.panelBg, -1, "COLOR LEGEND", t.cardMuted, 8, "", False)
    shp.TextFrame2.TextRange.Font.Bold = msoTrue: x = x + lw + 2
    Dim ct As Variant, cb As Variant, chw As Double
    ct = Array("New", "Changed", "Deleted from PW", "Unchanged")
    cb = Array(t.newBg, t.changedBg, t.deletedBg, t.dataBg)
    chw = (iw - lw - 2 - 4 * 4) / 4
    For k = 0 To 3
        Set shp = UiShape(wsV, "ui_chip" & k, msoShapeRectangle, x, y, chw, 20, CLng(cb(k)), t.accent, CStr(ct(k)), t.dataFg, 9, "", True)
        x = x + chw + 4
    Next k
    x = xr
    Set shp = UiShape(wsV, "ui_prev", msoShapeRoundedRectangle, x, y, 70, 20, t.inputBg, t.accent, ChrW(9664) & " PREV", t.inputFg, 9, wn & "PagePrev", True): x = x + 74
    Dim pg As String
    pg = "Page " & mPage & " / " & mPages & "   |   " & mTotal & " results" & IIf(mFolder <> "", "   |   folder: " & Replace(Left(mFolder, Len(mFolder) - 1), "\", " > "), "")
    Set shp = UiShape(wsV, "ui_page", msoShapeRectangle, x, y, 430, 20, t.panelBg, -1, pg, t.panelFg, 9, wn & "ClearFolderFilter", False)
    shp.TextFrame2.TextRange.Font.Bold = msoTrue: x = x + 434
    Set shp = UiShape(wsV, "ui_next", msoShapeRoundedRectangle, x, y, 70, 20, t.inputBg, t.accent, "NEXT " & ChrW(9654), t.inputFg, 9, wn & "PageNext", True)
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
    v = InputBox("Search phrase (empty = clear)." & vbCrLf & vbCrLf & "  HRX            contains HRX in any column" & vbCrLf & "  HRX-0??-*      * = any text, ? = any single character" & vbCrLf & "  filename:*HRX*   search only in a column (filename, path, link, updater, updated, status or any column name)" & vbCrLf & "  filename:""my file*""   quotes keep spaces together" & vbCrLf & "  -pdf  -path:Templates   exclude" & vbCrLf & "  several terms = all must match", "Search", CStr(ws.Range(VIEW_SEARCH).Value))
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
