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

' Linki PW do pobrania: zaznaczone wiersze Dashboard (kilka wierszy), a przy jednej komorce - WSZYSTKIE wyniki biezacego zapytania (wszystkie strony)
Public Function LinksFromSelection() As String
    Dim wsV As Worksheet, iLink As Long, i As Long, ix As Long, v As String, sb As String
    Set wsV = ThisWorkbook.Sheets(SH_VIEW)
    EnsureData
    iLink = ColOf("PW Link")
    If iLink = 0 Or gHits = 0 Then Notify "No results with a PW Link - run UPDATE DATA first.": Exit Function
    Dim first As Long, page As Long, r1 As Long, r2 As Long, multi As Boolean
    page = Val(CStr(wsV.Range(PAGE_CELL).Value)): If page < 1 Then page = 1
    first = (page - 1) * PAGE_SIZE + 1
    If ActiveSheet.Name = SH_VIEW Then
        If Selection.Rows.Count > 1 Then
            r1 = Selection.Row: r2 = Selection.Row + Selection.Rows.Count - 1
            If r1 <= VIEW_HEADER_ROW Then r1 = VIEW_HEADER_ROW + 1
            multi = (r2 >= r1)
        End If
    End If
    If multi Then
        For i = r1 To r2
            ix = first + (i - VIEW_HEADER_ROW) - 1
            If ix >= 1 And ix <= gHits Then
                v = CStr(gData(gIdx(ix), iLink))
                If UCase(Left(v, 3)) = "PW:" Then sb = sb & v & vbLf
            End If
        Next i
    Else
        For i = 1 To gHits
            v = CStr(gData(gIdx(i), iLink))
            If UCase(Left(v, 3)) = "PW:" Then sb = sb & v & vbLf
        Next i
    End If
    LinksFromSelection = sb
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
    If lst.Count = 0 Then Notify "No PW links (pw://...) in the list.": Exit Sub
    If target = "" Then Notify "Enter the target folder.": Exit Sub
    If ThisWorkbook.Path = "" Then Notify "Save the workbook to disk first": Exit Sub
    ThisWorkbook.Sheets(SH_VIEW).Range(TARGET_CELL).Value = target
    
    If MsgBox("Download " & lst.Count & " file(s) to:" & vbCrLf & target & vbCrLf & vbCrLf & _
              "The script drives the keyboard and the ProjectWise window - do not touch the keyboard or mouse meanwhile. To abort, close the PowerShell window.", _
              vbOKCancel + vbQuestion, "Download from PW") <> vbOK Then Exit Sub
    
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
    f.WriteLine "$dly = 2"
    f.WriteLine "function Key($k) { $sh.SendKeys($k); Start-Sleep -Seconds $dly }"
    f.WriteLine "function Focus($h) { if ($h -ne [IntPtr]::Zero) { [void][W]::SetForegroundWindow($h) }; Start-Sleep -Milliseconds 800 }"
    f.WriteLine "Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes"
    f.WriteLine "$UA = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]::Descendants"
    f.WriteLine "function Get-El($h, $type, $part) { try { $root = $UA::FromHandle($h); $cond = New-Object System.Windows.Automation.PropertyCondition($UA::ControlTypeProperty, $type); foreach ($e in $root.FindAll($TS, $cond)) { if ((-not $part) -or $e.Current.Name -like ('*' + $part + '*')) { if ($e.Current.IsEnabled) { return $e } } } } catch {}; return $null }"
    f.WriteLine "function Ui-Click($h, $name) { $e = Get-El $h ([System.Windows.Automation.ControlType]::Button) $name; if ($e) { try { $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); Start-Sleep -Seconds $dly; return $true } catch {} }; return $false }"
    f.WriteLine "function Ui-Select($h, $name) { $e = Get-El $h ([System.Windows.Automation.ControlType]::RadioButton) $name; if ($e) { try { $e.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep -Seconds 1; return $true } catch {} }; return $false }"
    f.WriteLine "function Ui-SetText($h, $text) { $e = Get-El $h ([System.Windows.Automation.ControlType]::Edit) $null; if ($e) { try { $e.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($text); Start-Sleep -Seconds 1; return $true } catch {} }; return $false }"
    f.WriteLine "function Ui-Uncheck($h, $name) { $e = Get-El $h ([System.Windows.Automation.ControlType]::CheckBox) $name; if ($e) { try { $tp = $e.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern); if ($tp.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) { $tp.Toggle(); Start-Sleep -Seconds 1 } } catch {} } }"
    f.WriteLine "function Has-Text($h, $txt) { try { $root = $UA::FromHandle($h); foreach ($e in $root.FindAll($TS, [System.Windows.Automation.Condition]::TrueCondition)) { if ($e.Current.Name -like ('*' + $txt + '*')) { return $true } } } catch {}; return $false }"
    f.WriteLine "$ok = 0; $fail = 0; $i = 0"
    f.WriteLine "foreach ($link in $links) {"
    f.WriteLine "  $i++; Write-Host ('[' + $i + '/' + $links.Count + '] ' + $link) -ForegroundColor Cyan"
    f.WriteLine "  try {"
    f.WriteLine "    Start-Process $link"
    f.WriteLine "    $h = Wait-Win 'Security Notice' 3; if ($h -ne [IntPtr]::Zero) { Focus $h; Key 'y' }"
    f.WriteLine "    Start-Sleep -Seconds 6"
    f.WriteLine "    $h = Wait-Win 'ProjectWise' 30; Focus $h"
    f.WriteLine "    Start-Sleep -Seconds 2; Key '%(dx)'"
    f.WriteLine "    $h = Wait-Win 'Export Wizard' 15"
    f.WriteLine "    if ($h -eq [IntPtr]::Zero) { Write-Host '  Alt+(DX): the wizard did not open yet - Enter' -ForegroundColor DarkGray; Key '{ENTER}'; $h = Wait-Win 'Export Wizard' 10 }"
    f.WriteLine "    if ($h -eq [IntPtr]::Zero) { Write-Host '  retry: Alt+D, then X separately' -ForegroundColor DarkGray; Key '{ESC}'; Key '%d'; Key 'x'; $h = Wait-Win 'Export Wizard' 15 }"
    f.WriteLine "    if ($h -eq [IntPtr]::Zero) { throw 'Export wizard window not found (menu Document > Export...)' }"
    f.WriteLine "    Start-Sleep -Seconds 2"
    f.WriteLine "    $end = (Get-Date).AddSeconds(240); $finished = $false; $last = ''; $pathSet = $false"
    f.WriteLine "    while ((Get-Date) -lt $end) {"
    f.WriteLine "      $w = [W]::Find('Warning!')"
    f.WriteLine "      if ($w -ne [IntPtr]::Zero) { Write-Host '  file exists - Yes to All' -ForegroundColor DarkGray; Focus $w; if (-not (Ui-Click $w 'Yes to All')) { [void](Ui-Click $w 'Yes') }; continue }"
    f.WriteLine "      $h = [W]::Find('Export Wizard'); if ($h -eq [IntPtr]::Zero) { $finished = $true; break }"
    f.WriteLine "      Focus $h"
    f.WriteLine "      if (Has-Text $h 'Review document export results') { if ($last -ne 'review') { Write-Host '  [4/4] results page - unticking Show exported files, Finish' -ForegroundColor DarkGray; $last = 'review' }; Ui-Uncheck $h 'Show exported'; [void](Ui-Click $h 'Finish') }"
    f.WriteLine "      elseif (Has-Text $h 'export in progress') { if ($last -ne 'progress') { Write-Host '  [3/4] export in progress...' -ForegroundColor DarkGray; $last = 'progress' } }"
    f.WriteLine "      elseif (Has-Text $h 'Define the export settings') { if ($last -ne 'settings') { Write-Host '  [2/4] export settings - setting the folder' -ForegroundColor DarkGray; $last = 'settings' }; if (-not $pathSet) { if (-not (Ui-SetText $h $target)) { Set-Clipboard -Value $target; Key '^a'; Key '^v' }; $pathSet = $true }; [void](Ui-Click $h 'Next') }"
    f.WriteLine "      elseif (Has-Text $h 'Welcome to the') { if ($last -ne 'welcome') { Write-Host '  [1/4] welcome page - Send to Folder, Next' -ForegroundColor DarkGray; $last = 'welcome' }; [void](Ui-Select $h 'Send to Folder'); [void](Ui-Click $h 'Next') }"
    f.WriteLine "      else { if ($last -ne 'unknown') { Write-Host '  unrecognised wizard page - trying Next / Finish' -ForegroundColor DarkGray; $last = 'unknown' }; if (-not (Ui-Click $h 'Finish')) { [void](Ui-Click $h 'Next') } }"
    f.WriteLine "      Start-Sleep -Seconds 2"
    f.WriteLine "    }"
    f.WriteLine "    if (-not $finished) { throw 'the export wizard did not finish within 4 minutes' }"
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
    fn = Application.GetOpenFilename("Excel (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , "Pick the OLD workbook with data")
    If VarType(fn) = vbBoolean Then Exit Sub
    If StrComp(CStr(fn), ThisWorkbook.FullName, vbTextCompare) = 0 Then Notify "That is the same file": Exit Sub
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
    InvalidateData
    RenderDashboard True
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Notify "Sheets moved: " & done
    Exit Sub
Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    src.Close False
    Notify "Error moving data: " & Err.Description
End Sub

' Button "DOWNLOAD FROM PW" on Dashboard: tworzy go, jesli skoroszyt go nie ma (stary plik z danymi), i podpina formularz
Public Sub EnsureViewButton()
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    ThisWorkbook.Sheets(SH_DATA).Shapes("btnUpdate").Delete          ' PW_Data to czysta baza - bez przyciskow
    ThisWorkbook.Sheets(SH_DATA).Shapes("btnNavigator").Delete
    ThisWorkbook.Sheets(SH_DATA).OLEObjects("TextBox1").Delete          ' szukanie jest w Dashboard
    Set ws = ThisWorkbook.Sheets(SH_VIEW)
    Set shp = ws.Shapes("btnGetFiles")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("C5").Left + 3, ws.Range("C5").Top + 3, 190, 28)
        shp.Name = "btnGetFiles"
        shp.Fill.ForeColor.RGB = RGB(127, 79, 0)
        shp.Line.Visible = msoFalse
        shp.TextFrame2.TextRange.Text = "DOWNLOAD FROM PW"
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
    End If
    shp.OnAction = "PobierzPlikiPW"
End Sub
