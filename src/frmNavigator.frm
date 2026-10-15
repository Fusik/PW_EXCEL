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
    Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const SC_CLOSE As Long = &HF060
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
        lStyle = lStyle Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
        SetWindowLongPtr hWnd, GWL_STYLE, lStyle
        
        lExStyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE)
        lExStyle = lExStyle Or WS_EX_APPWINDOW
        SetWindowLongPtr hWnd, GWL_EXSTYLE, lExStyle
        
        hMenu = GetSystemMenu(hWnd, 0)
        If hMenu <> 0 Then
            RemoveMenu hMenu, SC_CLOSE, MF_BYCOMMAND
        End If
        
        DrawMenuBar hWnd
    End If
    
    ' Wymuszenie dopasowania drzewka natychmiast po pokazaniu okna
    Call UserForm_Resize
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
    
    ' Skok do wybranego wiersza
    If Node.Tag <> "" And IsNumeric(Node.Tag) Then
        targetRow = CLng(Node.Tag)
        Dim wsV As Worksheet, vRow As Long
        Set wsV = ThisWorkbook.Sheets(SH_VIEW)
        vRow = targetRow - DATA_HEADER_ROW + VIEW_HEADER_ROW        ' wiersz Baza_PW -> wiersz w Widok (kopia 1:1)
        If StrComp(CStr(wsV.Cells(vRow, 1).Value), CStr(ws.Cells(targetRow, 1).Value), vbTextCompare) = 0 And Not wsV.Rows(vRow).Hidden Then
            wsV.Activate                                            ' glowny arkusz pracy - Widok
            wsV.Rows(vRow).Select
            ActiveWindow.ScrollRow = vRow - 10
            If ActiveWindow.ScrollRow < VIEW_HEADER_ROW + 1 Then ActiveWindow.ScrollRow = VIEW_HEADER_ROW + 1
        Else
            ws.Activate                                             ' widok nie pokazuje tego wiersza (kolumna A inna lub ukryty filtrem) - skok w Baza_PW
            ws.Rows(targetRow).Select
            ActiveWindow.ScrollRow = targetRow - 10
            If ActiveWindow.ScrollRow < 1 Then ActiveWindow.ScrollRow = 1
        End If
    End If
    
    ' Zwijanie/Rozwijanie folderów po kliknięciu w samą nazwę
    If Node.Children > 0 Then
        Node.Expanded = Not Node.Expanded
    End If
End Sub
