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
    f.WriteLine "try { $null = New-Item -ItemType Directory -Force -Path '" & ThisWorkbook.Path & "\PW_Logs'; Start-Transcript -Path ('" & ThisWorkbook.Path & "\PW_Logs\' + (Get-Date).ToString('yyyy-MM-dd_HHmmss') + '_download.log') -Force | Out-Null } catch {}"
    f.WriteLine "Add-Type -TypeDefinition @'"
    f.WriteLine "using System; using System.Text; using System.Runtime.InteropServices;"
    f.WriteLine "public class W {"
    f.WriteLine "  public delegate bool EP(IntPtr h, IntPtr l);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern bool EnumWindows(EP p, IntPtr l);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern bool IsWindowVisible(IntPtr h);"
    f.WriteLine "  [DllImport(""user32.dll"")] public static extern bool SetForegroundWindow(IntPtr h);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern bool SetCursorPos(int x, int y);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);"
    f.WriteLine "  public static void Click(int x, int y) { SetCursorPos(x, y); System.Threading.Thread.Sleep(150); mouse_event(2, 0, 0, 0, UIntPtr.Zero); mouse_event(4, 0, 0, 0, UIntPtr.Zero); }"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern bool EnumChildWindows(IntPtr p, EP cb, IntPtr l);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, StringBuilder s);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern IntPtr SendMessage(IntPtr h, int m, IntPtr w, IntPtr l);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern int GetClassName(IntPtr h, StringBuilder s, int n);"
    f.WriteLine "  [DllImport(""user32.dll"")] static extern bool IsWindowEnabled(IntPtr h);"
    f.WriteLine "  static string Txt(IntPtr h) { StringBuilder sb = new StringBuilder(1024); SendMessage(h, 0xD, (IntPtr)1024, sb); return sb.ToString(); }"
    f.WriteLine "  static string Cls(IntPtr h) { StringBuilder sb = new StringBuilder(64); GetClassName(h, sb, 64); return sb.ToString(); }"
    f.WriteLine "  public static string ChildTexts(IntPtr top) { StringBuilder all = new StringBuilder(); EnumChildWindows(top, delegate(IntPtr h, IntPtr l) { all.Append(Txt(h)).Append('|'); return true; }, IntPtr.Zero); return all.ToString(); }"
    f.WriteLine "  public static IntPtr ChildButton(IntPtr top, string part) { IntPtr f = IntPtr.Zero; EnumChildWindows(top, delegate(IntPtr h, IntPtr l) { if (Cls(h) == ""Button"") { string t = Txt(h).Replace(""&"", """"); if (t.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0) { f = h; return false; } } return true; }, IntPtr.Zero); return f; }"
    f.WriteLine "  public static bool Enabled(IntPtr h) { return IsWindowEnabled(h); }"
    f.WriteLine "  public static bool Checked(IntPtr h) { return SendMessage(h, 0xF0, IntPtr.Zero, IntPtr.Zero) != IntPtr.Zero; }"
    f.WriteLine "  public static void BtnClick(IntPtr h) { SendMessage(h, 0xF5, IntPtr.Zero, IntPtr.Zero); }"
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
    f.WriteLine "function Wait-Win($t, $sec) { $end = (Get-Date).AddSeconds($sec); while ((Get-Date) -lt $end) { $h = [W]::Find($t); if ($h -ne [IntPtr]::Zero) { return $h }; Start-Sleep -Milliseconds 150 }; return [IntPtr]::Zero }"
    f.WriteLine "$dly = 350   # ms - basic pause between actions; all other waits are multiples of it (raise it if PW is slow to load)"
    f.WriteLine "function Key($k) { $sh.SendKeys($k); Start-Sleep -Milliseconds $dly }"
    f.WriteLine "function Focus($h) { if ($h -ne [IntPtr]::Zero) { [void][W]::SetForegroundWindow($h) }; Start-Sleep -Milliseconds ($dly * 0.4) }"
    f.WriteLine "Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes"
    f.WriteLine "$UA = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]::Descendants"
    f.WriteLine "function Get-El($h, $type, $part) { try { $root = $UA::FromHandle($h); $cond = New-Object System.Windows.Automation.PropertyCondition($UA::ControlTypeProperty, $type); foreach ($e in $root.FindAll($TS, $cond)) { if ((-not $part) -or $e.Current.Name -like ('*' + $part + '*')) { if ($e.Current.IsEnabled) { return $e } } } } catch {}; return $null }"
    f.WriteLine "function Ui-Click($h, $name) { $e = Get-El $h ([System.Windows.Automation.ControlType]::Button) $name; if ($e) { try { $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); Start-Sleep -Milliseconds $dly; return $true } catch {} }; return $false }"
    f.WriteLine "function Ui-Select($h, $name) { $e = Get-El $h ([System.Windows.Automation.ControlType]::RadioButton) $name; if ($e) { try { $e.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep -Milliseconds ($dly * 0.5); return $true } catch {} }; return $false }"
    f.WriteLine "function Ui-SetText($h, $text) { $e = Get-El $h ([System.Windows.Automation.ControlType]::Edit) $null; if ($e) { try { $e.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($text); Start-Sleep -Milliseconds ($dly * 0.5); return $true } catch {} }; return $false }"
    f.WriteLine "function Ui-Uncheck($h, $name) { $e = Get-El $h ([System.Windows.Automation.ControlType]::CheckBox) $name; if ($e) { try { $tp = $e.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern); if ($tp.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) { $tp.Toggle(); Start-Sleep -Milliseconds ($dly * 0.5) }; return $true } catch {} }; return $false }"
    f.WriteLine "function Has-Text($h, $txt) { try { $root = $UA::FromHandle($h); foreach ($e in $root.FindAll($TS, [System.Windows.Automation.Condition]::TrueCondition)) { if ($e.Current.Name -like ('*' + $txt + '*')) { return $true } } } catch {}; return $false }"
    f.WriteLine "function Ui-Mouse($h, $name) { $e = Get-El $h ([System.Windows.Automation.ControlType]::Button) $name; if ($e) { try { $r = $e.Current.BoundingRectangle; [W]::Click([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)); Start-Sleep -Milliseconds $dly; return $true } catch {} }; return $false }"
    f.WriteLine "$script:tries = 0"
    f.WriteLine "function Press($h, $name, $alt) { $script:tries++; switch ($script:tries % 4) { 1 { if (-not (Ui-Click $h $name)) { Key $alt } } 2 { Key $alt } 3 { [void](Ui-Mouse $h $name) } 0 { Key '{ENTER}' } } }"
    f.WriteLine "$ok = 0; $fail = 0; $i = 0"
    f.WriteLine "foreach ($link in $links) {"
    f.WriteLine "  $i++; Write-Host ('[' + $i + '/' + $links.Count + '] ' + $link) -ForegroundColor Cyan"
    f.WriteLine "  try {"
    f.WriteLine "    Start-Process $link"
    f.WriteLine "    $h = Wait-Win 'Security Notice' 3; if ($h -ne [IntPtr]::Zero) { Focus $h; Key 'y' }"
    f.WriteLine "    Start-Sleep -Milliseconds ($dly * 6)"
    f.WriteLine "    $hp = Wait-Win 'ProjectWise Explorer' 30; if ($hp -eq [IntPtr]::Zero) { $hp = Wait-Win 'ProjectWise' 5 }; if ($hp -eq [IntPtr]::Zero) { throw 'ProjectWise window not found' }; Write-Host '  ProjectWise window found - opening the export wizard (Alt+D, X)' -ForegroundColor DarkGray"
    f.WriteLine "    $h = [IntPtr]::Zero"
    f.WriteLine "    for ($try = 1; $try -le 4 -and $h -eq [IntPtr]::Zero; $try++) {"
    f.WriteLine "      if ($try -gt 1) { Write-Host ('  wizard not open - attempt ' + $try + ' (waiting for PW to select the file)') -ForegroundColor DarkGray; Start-Sleep -Milliseconds ($dly * 3) }"
    f.WriteLine "      Focus $hp; Start-Sleep -Milliseconds ($dly * 2)"
    f.WriteLine "      if ($try -eq 3) { Key '%d'; Key 'x' } else { Key '%(dx)' }"
    f.WriteLine "      $h = Wait-Win 'Export Wizard' 8"
    f.WriteLine "    }"
    f.WriteLine "    if ($h -eq [IntPtr]::Zero) { throw 'Export wizard window not found (menu Document > Export...)' }"
    f.WriteLine "    Start-Sleep -Milliseconds $dly"
    f.WriteLine "    Write-Host '  [1/4] welcome page - Send to Folder, Next' -ForegroundColor DarkGray"
    f.WriteLine "    Focus $h; Key '%s'; Key '%n'"
    f.WriteLine "    Start-Sleep -Milliseconds ($dly * 1.5); $h = [W]::Find('Export Wizard'); if ($h -eq [IntPtr]::Zero) { throw 'wizard closed after the welcome page' }"
    f.WriteLine "    Write-Host '  [2/4] export settings - pasting the folder, Next' -ForegroundColor DarkGray"
    f.WriteLine "    Focus $h; Set-Clipboard -Value $target; Start-Sleep -Milliseconds ($dly * 0.5); Key '^a'; Key '^v'; Key '%n'"
    f.WriteLine "    Write-Host '  [3/4] export in progress...' -ForegroundColor DarkGray"
    f.WriteLine "    $t0 = Get-Date; $end = $t0.AddSeconds(300); $finished = $false; $unchk = $false; $since = $t0"
    f.WriteLine "    while ((Get-Date) -lt $end) {"
    f.WriteLine "      Start-Sleep -Milliseconds $dly"
    f.WriteLine "      $w = [W]::Find('Warning!')"
    f.WriteLine "      if ($w -ne [IntPtr]::Zero) { Write-Host '  file exists - Yes to All' -ForegroundColor DarkGray; Focus $w; $yb = [W]::ChildButton($w, 'Yes to All'); if ($yb -ne [IntPtr]::Zero) { [W]::BtnClick($yb) } elseif (-not (Ui-Click $w 'Yes to All')) { Key '{TAB}'; Key '{ENTER}' }; Start-Sleep -Milliseconds $dly; continue }"
    f.WriteLine "      $h = [W]::Find('Export Wizard'); if ($h -eq [IntPtr]::Zero) { $finished = $true; break }"
    f.WriteLine "      $now = Get-Date"
    f.WriteLine "      $fb = [W]::ChildButton($h, 'Finish'); $isDone = (([W]::ChildTexts($h)) -like '*successfully completed*') -or ($fb -ne [IntPtr]::Zero -and [W]::Enabled($fb)) -or (Has-Text $h 'Review document export results')"
    f.WriteLine "      if (-not $isDone -and ($now - $t0).TotalSeconds -gt 30 -and -not (Has-Text $h 'export in progress')) { $isDone = $true }"
    f.WriteLine "      if ($isDone) {"
    f.WriteLine "        Focus $h"
    f.WriteLine "        if (-not $unchk) { Write-Host '  [4/4] results page - unticking Show exported files, Finish' -ForegroundColor DarkGray; $unchk = $true; $cb = [W]::ChildButton($h, 'Show exported'); if ($cb -ne [IntPtr]::Zero) { if ([W]::Checked($cb)) { [W]::BtnClick($cb) } } elseif (-not (Ui-Uncheck $h 'Show exported')) { Key '%s' }; Start-Sleep -Milliseconds ($dly * 0.5); if ($fb -ne [IntPtr]::Zero) { [W]::BtnClick($fb) } else { Key '{ENTER}' }; $since = Get-Date }"
    f.WriteLine "        elseif (($now - $since).TotalSeconds -gt 4) { $since = $now; if ($fb -ne [IntPtr]::Zero) { [W]::BtnClick($fb) }; Key '{ENTER}'; if ([W]::Find('Export Wizard') -ne [IntPtr]::Zero) { [void](Ui-Mouse $h 'Finish') } }"
    f.WriteLine "      }"
    f.WriteLine "    }"
    f.WriteLine "    if (-not $finished) { throw 'the export wizard did not finish within 4 minutes' }"
    f.WriteLine "    $ok++; Start-Sleep -Milliseconds ($dly * 3)"
    f.WriteLine "  } catch { $fail++; Write-Host ('  BLAD: ' + $_.Exception.Message) -ForegroundColor Red }"
    f.WriteLine "}"
    f.WriteLine "Write-Host ('Gotowe. OK: ' + $ok + ', bledy: ' + $fail + '. Folder: ' + $target) -ForegroundColor Green"
    f.WriteLine "Read-Host 'Enter aby zamknac'"
    f.Close
    
    CreateObject("WScript.Shell").Run "powershell.exe -NoExit -ExecutionPolicy Bypass -Sta -File """ & psPath & """", 1, False
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
