Attribute VB_Name = "modTheme"
Option Explicit

' =======================================================================
' WYGLAD APLIKACJI: dashboard w arkuszu Widok (naglowek, wyszukiwarka, legenda, motywy Jasny / Ciemny / Niebieski)
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
End Type

Private Function GetTheme(ByVal nm As String) As TTheme
    Dim t As TTheme
    Select Case LCase(Trim(nm))
        Case "ciemny"
            t.titleBg = RGB(18, 18, 18): t.titleFg = RGB(240, 240, 240)
            t.panelBg = RGB(37, 37, 38): t.panelFg = RGB(210, 210, 210)
            t.inputBg = RGB(60, 60, 64): t.inputFg = RGB(255, 255, 255): t.accent = RGB(0, 150, 255)
            t.hdrBg = RGB(0, 90, 158): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(30, 30, 30): t.dataFg = RGB(224, 224, 224)
            t.folderBg = CLR_FOLDER_DARK: t.changedBg = RGB(120, 100, 20): t.deletedBg = RGB(120, 40, 48)
        Case "niebieski"
            t.titleBg = RGB(0, 70, 127): t.titleFg = RGB(255, 255, 255)
            t.panelBg = RGB(228, 238, 250): t.panelFg = RGB(0, 56, 104)
            t.inputBg = RGB(255, 255, 255): t.inputFg = RGB(20, 20, 20): t.accent = RGB(0, 92, 170)
            t.hdrBg = RGB(0, 92, 170): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(250, 252, 255): t.dataFg = RGB(30, 30, 30)
            t.folderBg = CLR_FOLDER_BLUE: t.changedBg = CLR_CHANGED: t.deletedBg = CLR_DELETED
        Case Else
            t.titleBg = RGB(31, 56, 100): t.titleFg = RGB(255, 255, 255)
            t.panelBg = RGB(242, 246, 252): t.panelFg = RGB(31, 56, 100)
            t.inputBg = RGB(255, 255, 255): t.inputFg = RGB(20, 20, 20): t.accent = RGB(47, 85, 151)
            t.hdrBg = RGB(31, 56, 100): t.hdrFg = RGB(255, 255, 255)
            t.dataBg = RGB(255, 255, 255): t.dataFg = RGB(32, 32, 32)
            t.folderBg = CLR_FOLDER: t.changedBg = CLR_CHANGED: t.deletedBg = CLR_DELETED
    End Select
    GetTheme = t
End Function

' Nakłada motyw na caly arkusz Widok; lastR = ostatni wiersz danych, lastC = ostatnia kolumna tabeli
Public Sub StyleDashboard(wsV As Worksheet, ByVal lastR As Long, ByVal lastC As Long)
    Dim t As TTheme, nm As String, i As Long, r As Long
    Dim nFold As Long, nDel As Long, nFile As Long, nChg As Long
    On Error Resume Next
    If Trim(CStr(wsV.Range(THEME_CELL).Value)) = "" Then wsV.Range(THEME_CELL).Value = "Jasny"
    nm = CStr(wsV.Range(THEME_CELL).Value)
    t = GetTheme(nm)
    If lastR < VIEW_HEADER_ROW Then lastR = VIEW_HEADER_ROW
    If lastC < 12 Then lastC = 12
    Application.ScreenUpdating = False
    
    ' --- zmienione komórki zapamiętaj przed nadpisaniem koloru ---
    Dim chg As Collection, fc As Range, firstAddr As String
    Set chg = New Collection
    Application.FindFormat.Clear
    Application.FindFormat.Interior.Color = CLR_CHANGED
    If lastR > VIEW_HEADER_ROW Then
        Set fc = wsV.Range(wsV.Cells(VIEW_HEADER_ROW + 1, 2), wsV.Cells(lastR, lastC)).Find(What:="", SearchFormat:=True)
        If Not fc Is Nothing Then
            firstAddr = fc.Address
            Do
                chg.Add fc.Address
                Set fc = wsV.Range(wsV.Cells(VIEW_HEADER_ROW + 1, 2), wsV.Cells(lastR, lastC)).FindNext(fc)
            Loop While Not fc Is Nothing And fc.Address <> firstAddr And chg.Count < 5000
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
            .Borders(xlInsideHorizontal).Color = RGB(200, 200, 200)
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
    wsV.Rows(VIEW_HEADER_ROW).RowHeight = 26
    
    ' --- panel gorny (wiersze 1-7) ---
    With wsV.Range(wsV.Cells(1, 1), wsV.Cells(7, 40))
        .Interior.Color = t.panelBg: .Font.Color = t.panelFg: .Font.Name = "Segoe UI": .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    With wsV.Range(wsV.Cells(1, 1), wsV.Cells(2, 40))
        .Interior.Color = t.titleBg: .Font.Color = t.titleFg
    End With
    With wsV.Range("A1"): .Font.Size = 20: .Font.Bold = True: .Value = "PW Tracker": End With
    wsV.Rows(1).RowHeight = 38
    wsV.Rows(2).RowHeight = 22
    wsV.Range("A2").Value = "Foldery: " & nFold & "   |   Pliki: " & nFile & "   |   Zmienione komorki: " & nChg & "   |   Usuniete z PW: " & nDel
    wsV.Range("A2").Font.Size = 10
    wsV.Rows(3).RowHeight = 26: wsV.Rows(4).RowHeight = 30
    
    ' etykiety i pola
    Dim lbl As Variant
    For Each lbl In Array("A3", "A4", "G3", "G4")
        With wsV.Range(CStr(lbl)): .Font.Bold = True: .HorizontalAlignment = xlRight: .Font.Color = t.panelFg: End With
    Next lbl
    wsV.Range("A3").Value = "WIDOK:": wsV.Range("A4").Value = "SZUKAJ:"
    wsV.Range("G3").Value = "FOLDER DOCELOWY:": wsV.Range("G4").Value = "MOTYW:"
    Dim fld As Variant
    For Each fld In Array("B3", "B4", "H3", "H4")
        With wsV.Range(CStr(fld))
            .Interior.Color = t.inputBg: .Font.Color = t.inputFg: .Font.Bold = True
            .Borders.LineStyle = xlContinuous: .Borders.Color = t.accent: .Borders.Weight = xlThin
            .HorizontalAlignment = xlLeft: .IndentLevel = 1
        End With
    Next fld
    With wsV.Range("B4")                                     ' wyszukiwarka: wyrazna, gruba obwodka
        .Borders.Weight = xlMedium: .Font.Size = 12
    End With
    With wsV.Range("C4"): .Value = "<- wpisz fraze (min. 2 znaki): lista filtruje sie od razu po wszystkich kolumnach": .Font.Italic = True: .Font.Bold = False: End With
    With wsV.Range(THEME_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Jasny,Ciemny,Niebieski"
    End With
    wsV.Range("D3").Font.Italic = True
    wsV.Range("D4").Font.Color = RGB(192, 0, 0)
    
    ' --- legenda kolorow ---
    wsV.Rows(6).RowHeight = 26
    wsV.Range("A6").Value = "LEGENDA KOLOROW"
    With wsV.Range("A6"): .Font.Bold = True: .HorizontalAlignment = xlRight: End With
    Dim chipCol As Variant, chipTxt As Variant, chipBg As Variant, k As Long
    chipTxt = Array("Folder", "Zmieniona wartosc", "Usuniety z PW", "Zwykly plik")
    chipBg = Array(t.folderBg, t.changedBg, t.deletedBg, t.dataBg)
    For k = 0 To 3
        With wsV.Cells(6, 2 + k)
            .Value = chipTxt(k)
            .Interior.Color = chipBg(k): .Font.Color = t.dataFg
            .HorizontalAlignment = xlCenter: .Font.Bold = True
            .Borders.LineStyle = xlContinuous: .Borders.Color = t.accent: .Borders.Weight = xlThin
        End With
    Next k
    
    ' --- pasek statusu ---
    wsV.Rows(7).RowHeight = 24
    With wsV.Range("A7"): .Font.Bold = True: .Font.Color = t.accent: End With
    wsV.Range(wsV.Cells(7, 1), wsV.Cells(7, 40)).Borders(xlEdgeBottom).Color = t.accent
    wsV.Range(wsV.Cells(7, 1), wsV.Cells(7, 40)).Borders(xlEdgeBottom).Weight = xlMedium
    Application.ScreenUpdating = True
End Sub
