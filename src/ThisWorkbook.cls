Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
    On Error Resume Next
    ' Znacznik wersji kodu - jesli nie widzisz "build ..." w Instrukcja!B3, import kodu sie nie powiodl
    ThisWorkbook.Sheets(SH_START).Range("B3").Value = "Wersja skoroszytu: V1.6  ·  build " & BUILD_ID & "  ·  Wystarczy wkleić link w arkuszu Foldery – resztę robi makro. Własne ustawienia: Atrybuty (TAK/NIE) i Widoki."
    ' Uruchamia Nawigator w tle (Modeless = 0) od razu po starcie Excela
    EnsureViewButton
    frmNavigator.Show 0
    frmNavigator.DockToView
    If ActiveSheet.Name <> SH_VIEW Then HideNav     ' nawigator tylko na arkuszu Widok
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    CancelDownloadPoll                              ' zatrzymaj odpytywanie pobierania w tle
End Sub

' Automaty: zmiana widoku (Widok!B3) oraz wklejenie linku folderu (Foldery, kolumna B)
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Fail
    If Sh.Name = SH_VIEW Then
        If Not Intersect(Target, Sh.Range(VIEW_SELECTOR & "," & THEME_CELL)) Is Nothing Then
            Application.EnableEvents = False
            BuildView
            Application.EnableEvents = True
        End If
        If Not Intersect(Target, Sh.Range(VIEW_SEARCH)) Is Nothing Then
            Application.EnableEvents = False
            FilterRows Sh, VIEW_HEADER_ROW, CStr(Sh.Range(VIEW_SEARCH).Value)
            Application.EnableEvents = True
        End If
    ElseIf Sh.Name = SH_FOLDERS Then
        If Not Intersect(Target, Sh.Range(Sh.Cells(FOLDERS_FIRST_ROW, FOLDERS_LINK_COL), Sh.Cells(Sh.Rows.Count, FOLDERS_LINK_COL))) Is Nothing Then
            OnFolderLinkPasted Intersect(Target, Sh.Range(Sh.Cells(FOLDERS_FIRST_ROW, FOLDERS_LINK_COL), Sh.Cells(Sh.Rows.Count, FOLDERS_LINK_COL)))
        End If
    End If
    Exit Sub
Fail:
    Application.EnableEvents = True
    Notify "Blad: " & Err.Description
End Sub

' Klik w kolumne A widoku: folder = filtruje tabele do jego zawartosci, naglowek "Drzewo folderow" = pokaz wszystko
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Static busy As Boolean
    If busy Then Exit Sub
    If Sh.Name <> SH_VIEW Then Exit Sub
    If Target.Cells.Count <> 1 Or Target.Column <> 1 Then Exit Sub
    On Error GoTo Done
    busy = True
    Application.EnableEvents = False
    If Target.Row = VIEW_HEADER_ROW Then
        ShowAllFolders
    ElseIf Target.Row > VIEW_HEADER_ROW Then
        If IsFolderColor(Target.Interior.Color) Then ShowFolderRows Target.Row
    End If
Done:
    Application.EnableEvents = True
    busy = False
End Sub

' Nawigator jest widoczny tylko na arkuszu Widok i przypiety do kolumny A
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name = SH_VIEW Then
        frmNavigator.Show 0
        frmNavigator.DockToView
    Else
        HideNav
    End If
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
    On Error Resume Next
    If ActiveSheet.Name = SH_VIEW Then frmNavigator.DockToView
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    HideNav                                         ' inny skoroszyt na wierzchu
End Sub

Private Sub Workbook_Activate()
    On Error Resume Next
    If ActiveSheet.Name = SH_VIEW Then
        frmNavigator.Show 0
        frmNavigator.DockToView
    End If
End Sub

' Zamyka nawigator tylko jesli jest zaladowany (kolejne Show buduje drzewo od nowa z aktualnych danych)
Private Sub HideNav()
    Dim f As Object
    For Each f In VBA.UserForms
        If f.Name = "frmNavigator" Then Unload f: Exit For
    Next f
End Sub
