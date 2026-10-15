Option Explicit

' Formularz pobierania plikow z PW: wklej linki (po jednym w wierszu) i folder docelowy.
' Kontrolki tworzy Import-VBA.ps1 (txtLinks, txtFolder, btnSel, btnStart, btnClose).

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.txtFolder.Text = CStr(ThisWorkbook.Sheets(SH_VIEW).Range(TARGET_CELL).Value)
End Sub

Private Sub btnSel_Click()
    Me.txtLinks.Text = Replace(LinksFromSelection(), vbLf, vbCrLf)
End Sub

Private Sub btnStart_Click()
    StartPobieranie Me.txtLinks.Text, Me.txtFolder.Text
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
