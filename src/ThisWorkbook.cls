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
    frmNavigator.Show 0
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
        If Not Intersect(Target, Sh.Range(VIEW_SELECTOR)) Is Nothing Then
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
