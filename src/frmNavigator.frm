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
    Dim ws As Worksheet, pTop As Pane, pBot As Pane, dc As LongPtr, k As Double
    Set ws = ThisWorkbook.Sheets(SH_VIEW)
    dc = GetDC(0): k = 72 / GetDeviceCaps(dc, 88): ReleaseDC 0, dc      ' 88 = LOGPIXELSX -> piksele na punkty formularza
    Set pTop = ActiveWindow.Panes(1)
    Set pBot = ActiveWindow.Panes(ActiveWindow.Panes.Count)
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double
    x1 = pTop.PointsToScreenPixelsX(ws.Columns(1).Left)
    x2 = pTop.PointsToScreenPixelsX(ws.Columns(2).Left)
    y1 = pTop.PointsToScreenPixelsY(ws.Rows(VIEW_HEADER_ROW).Top)
    y2 = pBot.PointsToScreenPixelsY(pBot.VisibleRange.Top + pBot.VisibleRange.Height)
    If x2 - x1 < 150 Then x2 = x1 + 150
    If y2 - y1 < 200 Then y2 = y1 + 400                 ' zabezpieczenie: gdy obliczenie pozycji sie nie powiodlo, okno nie zostanie puste/splaszczone
    If y1 <= 0 Then y1 = 100
    Me.StartUpPosition = 0
    Me.Left = x1 * k: Me.Top = y1 * k
    Me.Width = (x2 - x1) * k: Me.Height = (y2 - y1) * k
End Sub

' =====================================================================
' INICJALIZACJA FORMULARZA: Budowanie drzewka z Arkusza 1
' =====================================================================
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DATA)
    
    Dim maxR As Long, r As Long
    Dim pathCol As Long, nameCol As Long
    
    Dim fCell As Range
    Set fCell = ws.rows(11).Find(What:="PW Source Path", LookIn:=xlValues, LookAt:=xlWhole)
    If fCell Is Nothing Then Exit Sub
    pathCol = fCell.Column
    nameCol = 1
    
    Set fCell = ws.Columns(1).Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not fCell Is Nothing Then maxR = fCell.Row Else Exit Sub
    
    Me.TreeView1.Nodes.Clear
    Me.TreeView1.LineStyle = tvwRootLines
    Me.TreeView1.Indentation = 12
    
    Dim fullPath As String, fName As String
    Dim arrFolders() As String, i As Long
    Dim parentKey As String, currentKey As String
    Dim newNode As MSComctlLib.Node
    
    For r = 12 To maxR
        fName = Trim(ws.Cells(r, nameCol).Value)
        fullPath = Trim(ws.Cells(r, pathCol).Value)
        
        If fName <> "" Then
            If fullPath = "" Then fullPath = fName
            fullPath = Replace(fullPath, "/", "\")
            arrFolders = Split(fullPath, "\")
            parentKey = ""
            
            For i = LBound(arrFolders) To UBound(arrFolders)
                If Trim(arrFolders(i)) <> "" Then
                    currentKey = parentKey & arrFolders(i) & "\"
                    
                    On Error Resume Next
                    If parentKey = "" Then
                        Set newNode = Me.TreeView1.Nodes.Add(Key:=currentKey, Text:=arrFolders(i))
                    Else
                        Set newNode = Me.TreeView1.Nodes.Add(Relative:=parentKey, Relationship:=tvwChild, Key:=currentKey, Text:=arrFolders(i))
                    End If
                    On Error GoTo 0
                    
                    If Not newNode Is Nothing Then newNode.Expanded = False
                    parentKey = currentKey
                End If
            Next i
            
            If Trim(ws.Cells(r, pathCol).Value) <> "" Then
                Dim fileKey As String
                fileKey = parentKey & fName & "_" & r
                
                On Error Resume Next
                Set newNode = Me.TreeView1.Nodes.Add(Relative:=parentKey, Relationship:=tvwChild, Key:=fileKey, Text:=fName)
                On Error GoTo 0
                
                If Not newNode Is Nothing Then
                    newNode.Tag = r
                    If ws.Cells(r, 1).Interior.Color = CLR_DELETED Then
                        newNode.ForeColor = vbRed
                    Else
                        newNode.ForeColor = RGB(0, 102, 204)
                    End If
                End If
            Else
                ' Przypisanie wiersza zielonego nagłówka folderu
                On Error Resume Next
                Me.TreeView1.Nodes(currentKey).Tag = r
                On Error GoTo 0
            End If
        End If
    Next r
    
    ' Post-processing: Barwienie usuniętych (czerwonych) folderów
    Dim n As MSComctlLib.Node
    For Each n In Me.TreeView1.Nodes
        If n.Parent Is Nothing Then
            CheckAndColorFolders n
        End If
    Next n
End Sub

' =====================================================================
' FUNKCJA REKURENCYJNA: Oznacza foldery na czerwono jeśli wszystko usunięto
' =====================================================================
Private Function CheckAndColorFolders(ByVal n As MSComctlLib.Node) As Boolean
    Dim childNode As MSComctlLib.Node
    Dim allRed As Boolean
    
    If n.Children = 0 Then
        CheckAndColorFolders = (n.ForeColor = vbRed)
    Else
        allRed = True
        Set childNode = n.Child
        
        Do While Not childNode Is Nothing
            If Not CheckAndColorFolders(childNode) Then
                allRed = False
            End If
            Set childNode = childNode.Next
        Loop
        
        If allRed Then
            n.ForeColor = vbRed
            CheckAndColorFolders = True
        Else
            n.ForeColor = vbBlack
            CheckAndColorFolders = False
        End If
    End If
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
    Dim targetRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_DATA)
    
    Dim wsV As Worksheet, vRow As Long, sr As Long
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    On Error Resume Next
    If Right(Node.Key, 1) = "\" Then                                ' folder: filtruj widok do jego zawartosci
        wsV.Activate
        vRow = ShowFolderByKey(Node.Key)
    ElseIf Node.Tag <> "" And IsNumeric(Node.Tag) Then              ' plik: odslon wiersz w widoku
        targetRow = CLng(Node.Tag)
        vRow = targetRow - DATA_HEADER_ROW + VIEW_HEADER_ROW        ' wiersz Baza_PW -> wiersz w Widok (kopia 1:1)
        If vRow > VIEW_HEADER_ROW Then
            wsV.Activate
            wsV.Rows(vRow).Hidden = False
        End If
    End If
    If vRow > VIEW_HEADER_ROW Then
        wsV.Cells(vRow, 2).Select
        sr = vRow - 10: If sr < VIEW_HEADER_ROW + 1 Then sr = VIEW_HEADER_ROW + 1
        ActiveWindow.ScrollRow = sr
    End If
    On Error GoTo 0
    
    ' Zwijanie/Rozwijanie folderów po kliknięciu w samą nazwę
    If Node.Children > 0 Then
        Node.Expanded = Not Node.Expanded
    End If
End Sub
