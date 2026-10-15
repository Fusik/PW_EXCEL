Attribute VB_Name = "modDownload"
Option Explicit

' =======================================================================
' POBIERANIE PLIKOW Z PW PRZEZ AUTOMAT KLAWIATURY (konto report-only nie moze uzyc CopyOut-PWDocuments)
' Dla kazdego wiersza widoku z PW Link: otwiera link, w PW Explorer Alt+D, E (Export), kreator "Send to Folder",
' wkleja folder docelowy i konczy kreator. Dziala w osobnym oknie PowerShell - nie dotykaj klawiatury w tym czasie.
' =======================================================================
Public Sub PobierzPlikiPW()
    frmPobierz.Show vbModeless                      ' formularz: lista linkow + folder docelowy
End Sub

' Linki PW z zaznaczenia w Widoku: kilka wierszy = te wiersze + pliki zaznaczonych folderow; jedna komorka = wszystko (bez wierszy ukrytych wyszukiwaniem)
Public Function LinksFromSelection() As String
    Dim wsV As Worksheet, lastC As Long, c As Long, linkCol As Long, r As Long, lnk As String
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    lastC = wsV.Cells(VIEW_HEADER_ROW, wsV.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If StrComp(Trim(CStr(wsV.Cells(VIEW_HEADER_ROW, c).Value)), "PW Link", vbTextCompare) = 0 Then linkCol = c: Exit For
    Next c
    If linkCol = 0 Then Notify "Widok nie ma kolumny 'PW Link' - dodaj ja do widoku (arkusz Widoki).": Exit Function
    Dim multi As Boolean, folderSel As Boolean, inSel As Boolean, skipHidden As Boolean, lastR As Long
    lastR = wsV.UsedRange.Row + wsV.UsedRange.Rows.Count - 1
    multi = (ActiveSheet.Name = SH_VIEW And Selection.Rows.Count > 1)
    skipHidden = (Len(Trim(CStr(wsV.Range(VIEW_SEARCH).Value))) >= 2)
    For r = VIEW_HEADER_ROW + 1 To lastR
        If wsV.Cells(r, 1).Interior.Color = CLR_FOLDER Then
            folderSel = multi And Not Intersect(wsV.Rows(r), Selection) Is Nothing
        Else
            inSel = Not multi
            If multi Then inSel = folderSel Or Not Intersect(wsV.Rows(r), Selection) Is Nothing
            If inSel And Not (skipHidden And wsV.Rows(r).Hidden) Then
                If wsV.Cells(r, linkCol).Hyperlinks.Count > 0 Then lnk = wsV.Cells(r, linkCol).Hyperlinks(1).Address Else lnk = Trim(CStr(wsV.Cells(r, linkCol).Value))
                If UCase(Left(lnk, 3)) = "PW:" And InStr(lnk, "/Documents/D{") > 0 Then LinksFromSelection = LinksFromSelection & lnk & vbLf
            End If
        End If
    Next r
End Function

' Start automatu pobierania: linksText = linki PW (po jednym w wierszu), target = folder docelowy
Public Sub StartPobieranie(ByVal linksText As String, ByVal target As String)
    Dim lst As Object, ln As Variant, t As String
    Set lst = CreateObject("System.Collections.ArrayList")
    For Each ln In Split(Replace(Replace(linksText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
        t = Trim(CStr(ln))
        If UCase(Left(t, 3)) = "PW:" Then lst.Add t
    Next ln
    target = Trim(target)
    If lst.Count = 0 Then Notify "Brak linkow PW (pw://...) na liscie.": Exit Sub
    If target = "" Then Notify "Podaj folder docelowy.": Exit Sub
    If ThisWorkbook.Path = "" Then Notify "Zapisz najpierw plik na dysku": Exit Sub
    ThisWorkbook.Sheets(SH_VIEW).Range("H3").Value = target
    
    If MsgBox("Pobrac " & lst.Count & " plik(ow) do:" & vbCrLf & target & vbCrLf & vbCrLf & _
              "Skrypt steruje klawiatura i oknem ProjectWise - w trakcie nie dotykaj klawiatury ani myszy. Przerwanie: zamknij okno PowerShell.", _
              vbOKCancel + vbQuestion, "Pobieranie plikow z PW") <> vbOK Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(target) Then CreateObject("WScript.Shell").Run "cmd /c mkdir """ & target & """", 0, True
    
    Dim listPath As String, psPath As String, f As Object, i As Long
    listPath = ThisWorkbook.Path & "\PW_Pobierz_linki.txt"
    psPath = ThisWorkbook.Path & "\PW_Pobierz.ps1"
    Set f = fso.CreateTextFile(listPath, True, True)
    f.WriteLine target
    For i = 0 To lst.Count - 1: f.WriteLine lst(i): Next i
    f.Close
    
    Set f = fso.CreateTextFile(psPath, True, True)
    f.WriteLine "Add-Type -TypeDefinition @'"
    f.WriteLine "using System; using System.Text; using System.Runtime.InteropServices;"
    f.WriteLine "public class W {"
    f.WriteLine "  public delegate bool EP(IntPtr h, IntPtr l);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern bool EnumWindows(EP p, IntPtr l);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern bool IsWindowVisible(IntPtr h);"
    f.WriteLine "  [DllImport(""user32.dll"")] public static extern bool SetForegroundWindow(IntPtr h);"
    f.WriteLine "  public static IntPtr Find(string part) {"
    f.WriteLine "    IntPtr found = IntPtr.Zero;"
    f.WriteLine "    EnumWindows(delegate(IntPtr h, IntPtr l) {"
    f.WriteLine "      if (!IsWindowVisible(h)) return true;"
    f.WriteLine "      StringBuilder sb = new StringBuilder(256); GetWindowText(h, sb, 256);"
    f.WriteLine "      if (sb.ToString().IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0) { found = h; return false; }"
    f.WriteLine "      return true; }, IntPtr.Zero);"
    f.WriteLine "    return found; } }"
    f.WriteLine "'@"
    f.WriteLine "$lines = @(Get-Content -Path '" & listPath & "')"
    f.WriteLine "$target = $lines[0]; $links = $lines[1..($lines.Count - 1)]"
    f.WriteLine "$sh = New-Object -ComObject WScript.Shell"
    f.WriteLine "function Wait-Win($t, $sec) { $end = (Get-Date).AddSeconds($sec); while ((Get-Date) -lt $end) { $h = [W]::Find($t); if ($h -ne [IntPtr]::Zero) { return $h }; Start-Sleep -Milliseconds 300 }; return [IntPtr]::Zero }"
    f.WriteLine "function Focus($h) { if ($h -ne [IntPtr]::Zero) { [void][W]::SetForegroundWindow($h) }; Start-Sleep -Milliseconds 400 }"
    f.WriteLine "$ok = 0; $fail = 0; $i = 0"
    f.WriteLine "foreach ($link in $links) {"
    f.WriteLine "  $i++; Write-Host ('[' + $i + '/' + $links.Count + '] ' + $link) -ForegroundColor Cyan"
    f.WriteLine "  try {"
    f.WriteLine "    Start-Process $link"
    f.WriteLine "    $h = Wait-Win 'Security Notice' 3; if ($h -ne [IntPtr]::Zero) { Focus $h; $sh.SendKeys('y') }"
    f.WriteLine "    Start-Sleep -Seconds 5"
    f.WriteLine "    $h = Wait-Win 'ProjectWise' 20; Focus $h"
    f.WriteLine "    $sh.SendKeys('%d'); Start-Sleep -Milliseconds 500; $sh.SendKeys('e')"
    f.WriteLine "    $h = Wait-Win 'Folder Export Wizard' 15; if ($h -eq [IntPtr]::Zero) { throw 'brak okna Folder Export Wizard' }"
    f.WriteLine "    Focus $h; $sh.SendKeys('%s'); Start-Sleep -Milliseconds 300; $sh.SendKeys('%n')"
    f.WriteLine "    $h = Wait-Win 'Document Export Wizard' 15; if ($h -eq [IntPtr]::Zero) { throw 'brak okna Document Export Wizard' }"
    f.WriteLine "    Focus $h; Set-Clipboard -Value $target; $sh.SendKeys('^a'); $sh.SendKeys('^v'); Start-Sleep -Milliseconds 300; $sh.SendKeys('%n')"
    f.WriteLine "    $end = (Get-Date).AddSeconds(90)"
    f.WriteLine "    while ((Get-Date) -lt $end -and [W]::Find('Export Wizard') -ne [IntPtr]::Zero) { Start-Sleep -Seconds 2; $h = [W]::Find('Document Export Wizard'); if ($h -ne [IntPtr]::Zero) { Focus $h; $sh.SendKeys('%f') } }"
    f.WriteLine "    $ok++"
    f.WriteLine "  } catch { $fail++; Write-Host ('  BLAD: ' + $_.Exception.Message) -ForegroundColor Red }"
    f.WriteLine "}"
    f.WriteLine "Write-Host ('Gotowe. OK: ' + $ok + ', bledy: ' + $fail + '. Folder: ' + $target) -ForegroundColor Green"
    f.WriteLine "Read-Host 'Enter aby zamknac'"
    f.Close
    
    CreateObject("WScript.Shell").Run "powershell.exe -ExecutionPolicy Bypass -Sta -File """ & psPath & """", 1, False
End Sub

' =======================================================================
' PRZENIES DANE: kopiuje Baza_PW, Atrybuty, Diagnostyka (i definicje Widoki, Foldery) ze starszego skoroszytu do tego
' =======================================================================
Public Sub PrzeniesDanePW()
    Dim fn As Variant, src As Workbook, nm As Variant, wsS As Worksheet, wsT As Worksheet, done As String
    fn = Application.GetOpenFilename("Excel (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , "Wskaz STARY skoroszyt z danymi")
    If VarType(fn) = vbBoolean Then Exit Sub
    If StrComp(CStr(fn), ThisWorkbook.FullName, vbTextCompare) = 0 Then Notify "To jest ten sam plik": Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Fail
    Set src = Workbooks.Open(CStr(fn), False, True)      ' tylko do odczytu, bez aktualizacji lacz
    For Each nm In Array(SH_DATA, SH_ATTRS, SH_DIAG, SH_VIEWS, SH_FOLDERS)
        Set wsS = Nothing: Set wsT = Nothing
        On Error Resume Next
        Set wsS = src.Sheets(CStr(nm)): Set wsT = ThisWorkbook.Sheets(CStr(nm))
        On Error GoTo Fail
        If Not wsS Is Nothing And Not wsT Is Nothing Then
            wsT.Cells.Clear
            wsT.Cells.Hyperlinks.Delete
            wsS.UsedRange.Copy wsT.Range(wsS.UsedRange.Cells(1, 1).Address)
            Dim c As Long
            For c = 1 To wsS.UsedRange.Columns.Count
                wsT.Columns(wsS.UsedRange.Columns(c).Column).ColumnWidth = wsS.Columns(wsS.UsedRange.Columns(c).Column).ColumnWidth
            Next c
            done = done & nm & ", "
        End If
    Next nm
    Application.CutCopyMode = False
    src.Close False
    On Error Resume Next
    BuildView
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Notify "Przeniesiono arkusze: " & done
    Exit Sub
Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    src.Close False
    Notify "Blad przenoszenia danych: " & Err.Description
End Sub

' Przycisk "POBIERZ PLIKI" w Widoku: tworzy go, jesli skoroszyt go nie ma (stary plik z danymi), i podpina formularz
Public Sub EnsureViewButton()
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Sheets(SH_VIEW)
    On Error Resume Next
    Set shp = ws.Shapes("btnGetFiles")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("C5").Left + 3, ws.Range("C5").Top + 3, 190, 28)
        shp.Name = "btnGetFiles"
        shp.Fill.ForeColor.RGB = RGB(127, 79, 0)
        shp.Line.Visible = msoFalse
        shp.TextFrame2.TextRange.Text = "POBIERZ PLIKI Z PW"
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
    End If
    shp.OnAction = "PobierzPlikiPW"
End Sub
