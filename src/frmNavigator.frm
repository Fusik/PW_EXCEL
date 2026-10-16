Attribute VB_Name = "frmNavigator"
Attribute VB_Base = "0{7A9FA555-645C-4080-A45B-8C938F0902D6}{EEE34524-D1FA-4404-BEFD-78011502BE2F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' =====================================================================
' BILIOTEKI SYSTEMU WINDOWS (API) DO ZMIANY WYGLĄDU OKNA
' =====================================================================
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const SC_CLOSE As Long = &HF060
Private Const SC_MOVE As Long = &HF010
Private Const SC_SIZE As Long = &HF000
Private Const MF_BYCOMMAND As Long = &H0

' =====================================================================
' ZDARZENIE ACTIVATE: Wymusza ikone na pasku i wywołuje odświeżenie GUI
' =====================================================================
Private Sub UserForm_Activate()
    #If VBA7 Then
        Dim hWnd As LongPtr, lStyle As LongPtr, hMenu As LongPtr, lExStyle As LongPtr
    #Else
        Dim hWnd As Long, lStyle As Long, hMenu As Long, lExStyle As Long
    #End If
    
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    
    If hWnd <> 0 Then
        lStyle = GetWindowLongPtr(hWnd, GWL_STYLE)
        lStyle = lStyle And Not (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)   ' okno przypiete: bez zmiany rozmiaru/minimalizacji
        SetWindowLongPtr hWnd, GWL_STYLE, lStyle
        
        lExStyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE)
        lExStyle = lExStyle Or WS_EX_APPWINDOW
        SetWindowLongPtr hWnd, GWL_EXSTYLE, lExStyle
        
        hMenu = GetSystemMenu(hWnd, 0)
        If hMenu <> 0 Then
            RemoveMenu hMenu, SC_CLOSE, MF_BYCOMMAND
            RemoveMenu hMenu, SC_MOVE, MF_BYCOMMAND                  ' bez przesuwania
            RemoveMenu hMenu, SC_SIZE, MF_BYCOMMAND
        End If
        
        DrawMenuBar hWnd
    End If
    
    ' Wymuszenie dopasowania drzewka natychmiast po pokazaniu okna
    Call UserForm_Resize
End Sub

' =====================================================================
' PRZYPIECIE DO WIDOKU: nawigator lezy na kolumnie A (od wiersza naglowka 8 w dol, na wysokosc okna)
' =====================================================================
Public Sub DockToView()
    On Error Resume Next
    If ActiveSheet.Name <> SH_VIEW Then Exit Sub
    Dim ws As Worksheet, pBot As Pane, dc As LongPtr, k As Double, sc As Double, vr As Range
    Set ws = ThisWorkbook.Sheets(SH_VIEW)
    dc = GetDC(0): k = 72 / GetDeviceCaps(dc, 88): ReleaseDC 0, dc      ' 88 = LOGPIXELSX -> piksele na punkty formularza
    ActiveWindow.Panes(ActiveWindow.Panes.Count).Activate               ' dolny prawy panel (tabela) - jedyny, ktorego geometrie liczymy
    Set pBot = ActiveWindow.Panes(ActiveWindow.Panes.Count)
    Set vr = pBot.VisibleRange
    Dim px0 As Double, py0 As Double, py1 As Double
    px0 = pBot.PointsToScreenPixelsX(vr.Left)                           ' lewa krawedz panelu tabeli = prawa krawedz zamrozonych kolumn
    py0 = pBot.PointsToScreenPixelsY(vr.Top)                            ' gorna krawedz panelu tabeli = dolna krawedz zamrozonego naglowka
    py1 = pBot.PointsToScreenPixelsY(vr.Top + vr.Height)
    sc = (pBot.PointsToScreenPixelsY(vr.Top + 100) - py0) / 100         ' piksele na punkt arkusza (DPI * zoom)
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double
    x1 = px0 - (ws.Columns(3).Left - ws.Columns(1).Left) * sc            ' lewa krawedz zamrozonych kolumn A:B
    x2 = x1 + (ws.Columns(2).Left - ws.Columns(1).Left) * sc             ' nawigator ma szerokosc kolumny A
    y1 = py0 - (ws.Rows(VIEW_HEADER_ROW + 1).Top - ws.Rows(VIEW_HEADER_ROW).Top) * sc   ' od wiersza naglowka tabeli
    y1 = y1 + 14                                                       ' korekta ramki okna
    y2 = py1 - 5                                                       ' nawigator siega do paska zakladek
    If x2 - x1 < 150 Then x2 = x1 + 300
    If y2 - y1 < 200 Then y2 = y1 + 400                                 ' zabezpieczenie przed spłaszczeniem
    If y1 <= 0 Then y1 = 100
    Me.StartUpPosition = 0
    Me.Left = x1 * k: Me.Top = y1 * k
    Me.Width = (x2 - x1) * k: Me.Height = (y2 - y1) * k
    ApplyNavTheme
End Sub

' Kolory drzewa i tla zgodne z motywem arkusza Dashboard (Dark = ciemne)
Public Sub ApplyNavTheme()
    On Error Resume Next
    Dim th As String, dark As Boolean, bg As Long, fg As Long, hForm As LongPtr, hTree As LongPtr
    th = LCase(Trim(CStr(ThisWorkbook.Sheets(SH_VIEW).Range(THEME_CELL).Value)))
    dark = (th = "dark" Or th = "ciemny")
    bg = IIf(dark, RGB(30, 30, 30), RGB(255, 255, 255))
    fg = IIf(dark, RGB(224, 224, 224), RGB(0, 0, 0))
    Me.BackColor = IIf(dark, RGB(37, 37, 38), &H8000000F)
    hForm = FindWindow("ThunderDFrame", Me.Caption)
    hTree = FindWindowEx(hForm, 0, "TreeView20WndClass", vbNullString)
    If hTree <> 0 Then
        SendMessage hTree, &H111D, 0, bg              ' TVM_SETBKCOLOR
        SendMessage hTree, &H111E, 0, fg              ' TVM_SETTEXTCOLOR
    End If
End Sub

' =====================================================================
' INICJALIZACJA FORMULARZA: Budowanie drzewka z Arkusza 1
' =====================================================================
Private Sub UserForm_Initialize()
    Dim i As Long, j As Long, n As Long, pth As String, segs() As String, key As String, parentKey As String
    Dim cnt As Object, ord As Object, iPath As Long
    On Error Resume Next
    EnsureData
    iPath = ColOf("PW Source Path")
    On Error GoTo 0
    Me.TreeView1.Nodes.Clear
    Me.TreeView1.LineStyle = tvwRootLines
    Me.TreeView1.Indentation = 12
    If gN = 0 Or iPath = 0 Then Exit Sub
    
    Set cnt = CreateObject("Scripting.Dictionary"): cnt.CompareMode = 1     ' klucz folderu (z \) -> liczba plikow w calym poddrzewie
    Set ord = CreateObject("Scripting.Dictionary"): ord.CompareMode = 1     ' unikalne sciezki w kolejnosci wystapienia
    For i = 1 To gN
        pth = Replace(Trim(CStr(gData(i, iPath))), "/", "\")
        If pth <> "" Then
            If Not ord.exists(pth) Then ord(pth) = 1
            segs = Split(pth, "\")
            key = ""
            For j = 0 To UBound(segs)
                If Trim(segs(j)) <> "" Then
                    key = key & segs(j) & "\"
                    cnt(key) = cnt(key) + 1
                End If
            Next j
        End If
    Next i
    
    Dim p As Variant, nd As MSComctlLib.Node, made As Object, done As Long
    Set made = CreateObject("Scripting.Dictionary"): made.CompareMode = 1     ' utworzone klucze (wyszukiwanie klucza w TreeView jest liniowe)
    For Each p In ord.Keys
        segs = Split(CStr(p), "\")
        parentKey = "": key = ""
        For j = 0 To UBound(segs)
            If Trim(segs(j)) <> "" Then
                key = key & segs(j) & "\"
                If Not made.exists(key) Then
                    If parentKey = "" Then
                        Set nd = Me.TreeView1.Nodes.Add(Key:=key, Text:=segs(j) & "  (" & cnt(key) & ")")
                    Else
                        Set nd = Me.TreeView1.Nodes.Add(Relative:=parentKey, Relationship:=tvwChild, Key:=key, Text:=segs(j) & "  (" & cnt(key) & ")")
                    End If
                    nd.Expanded = (parentKey = "")
                    made(key) = 1
                End If
                parentKey = key
            End If
        Next j
        done = done + 1: If done Mod 500 = 0 Then Application.StatusBar = "Building folder tree: " & done & " / " & ord.Count: DoEvents
    Next p
    Application.StatusBar = False
    Me.TreeView1.Sorted = True
End Sub

Private Function NodeExists(ByVal key As String) As Boolean
    Dim nd As MSComctlLib.Node
    On Error Resume Next
    Set nd = Me.TreeView1.Nodes(key)
    NodeExists = Not nd Is Nothing
    On Error GoTo 0
End Function

' =====================================================================
' DYNAMICZNE ZMIENIANIE ROZMIARU OKNA
' =====================================================================
Private Sub UserForm_Resize()
    On Error Resume Next
    If Me.InsideWidth > 15 And Me.InsideHeight > 15 Then
        Me.TreeView1.Top = 5
        Me.TreeView1.Left = 5
        Me.TreeView1.Width = Me.InsideWidth - 10
        Me.TreeView1.Height = Me.InsideHeight - 10
    End If
    On Error GoTo 0
End Sub

' =====================================================================
' BLOKADA ZAMYKANIA FORMULARZA
' =====================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

' =====================================================================
' INTERAKCJA: Nawigowanie po arkuszu z drzewka
' =====================================================================
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    ThisWorkbook.Sheets(SH_VIEW).Activate
    SetFolderFilter Node.Key                        ' tabela w Dashboard pokazuje zawartosc tego folderu (zapytanie na tablicy w pamieci)
    On Error GoTo 0
    If Node.Children > 0 Then Node.Expanded = Not Node.Expanded
End Sub
