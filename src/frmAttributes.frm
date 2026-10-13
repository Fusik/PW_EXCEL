Attribute VB_Name = "frmAttributes"
Attribute VB_Base = "0{DDD8D4F0-1992-44C0-A0AA-CE539BF07DC9}{8DEA90B5-A148-4293-A831-E2058993ED01}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public SelectedAttributes As Collection
Public IsCancelled As Boolean

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Select attributes to export"
    IsCancelled = True
    Set SelectedAttributes = New Collection
    
    Dim ws3 As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim attrName As String
    
    On Error Resume Next
    Set ws3 = ThisWorkbook.Sheets(SH_ATTRS)
    On Error GoTo 0
    
    If ws3 Is Nothing Then
        MsgBox "Sheet '" & SH_ATTRS & "' not found.", vbCritical
        Exit Sub
    End If
    
    lastRow = ws3.Cells(ws3.rows.count, "A").End(xlUp).Row
    
    For i = ATTRS_FIRST_ROW To lastRow
        attrName = Trim(ws3.Cells(i, 1).Value)
        If attrName <> "" Then
            ListBox1.AddItem attrName
        End If
    Next i
    
    If ListBox1.ListCount = 0 Then
        MsgBox "Column A in sheet '" & SH_ATTRS & "' is empty. Add attribute names there!", vbExclamation
    End If
End Sub

Private Sub CommandButton1_Click()
    IsCancelled = False
    Dim i As Long
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            SelectedAttributes.Add ListBox1.List(i)
        End If
    Next i
    Me.Hide
End Sub
