Attribute VB_Name = "Module2"
Option Explicit

' =======================================================================
' NARZĘDZIE DIAGNOSTYCZNE ŚRODOWISKA
' Weryfikuje poprawne zapisanie pliku, listę potrzebnych arkuszy
' oraz czy kluczowy moduł bentley'a PWPS_DAB jest poprawnie zainstalowany.
' =======================================================================
Sub PW_diagnostyka()
    Dim msg As String
    msg = "=== DIAGNOSTYKA ŚRODOWISKA PROJECTWISE ===" & vbCrLf & vbCrLf

    ' 1. Sprawdzenie lokalizacji pliku
    If ThisWorkbook.Path = "" Then
        msg = msg & "? BŁĄD: Plik Excela nie jest zapisany na dysku!" & vbCrLf
    Else
        msg = msg & "? Ścieżka robocza OK: " & ThisWorkbook.Path & vbCrLf
    End If

    ' 2. Sprawdzenie obecności kluczowych arkuszy
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Sheet_1", "Sheet_2", "Sheet_3")
    Dim i As Integer
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        If ws Is Nothing Then
            msg = msg & "? BŁĄD: Brakuje arkusza '" & sheetNames(i) & "'!" & vbCrLf
        Else
            msg = msg & "? Arkusz '" & sheetNames(i) & "' jest gotowy." & vbCrLf
        End If
        Set ws = Nothing
    Next i

    ' 3. Sprawdzenie dostępności modułu PWPS_DAB w systemie PowerShell
    msg = msg & vbCrLf & "Sprawdzam integrację z ProjectWise (PWPS_DAB)..." & vbCrLf
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim psCommand As String, psResult As Integer
    
    ' Odpalamy cichy proces PS i wymuszamy zwrócenie flagi sukcesu
    psCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""if (Get-Module -ListAvailable -Name PWPS_DAB) { exit 0 } else { exit 1 }"""
    psResult = wsh.Run(psCommand, 0, True)

    If psResult = 0 Then
        msg = msg & "? Moduł PWPS_DAB jest zainstalowany i gotowy." & vbCrLf
    Else
        msg = msg & "? BŁĄD: Brak modułu PWPS_DAB w systemie!" & vbCrLf
    End If

    ' Wyświetlenie wyniku analizy
    MsgBox msg, vbInformation, "Status Środowiska"
End Sub

