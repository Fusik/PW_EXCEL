Attribute VB_Name = "modData"
Option Explicit

' =======================================================================
' PAMIEC PODRECZNA DANYCH + WYSZUKIWARKA
' PW_Data jest baza (plaska tabela, bez formatowania). Dashboard nie kopiuje danych - pokazuje tylko jedna strone wynikow.
' Cala baza jest czytana do tablicy raz (EnsureData); wyszukiwanie i filtrowanie dziala na tablicy, nie na komorkach.
' Skladnia wyszukiwania:
'   fraza                  - zawiera (w dowolnej kolumnie)          HRX
'   *  ?                   - znaki wieloznaczne                      HRX-0??-*   (* = dowolny ciag, ? = dowolny znak)
'   kolumna:wartosc        - tylko w tej kolumnie                    filename:*HRX*    path:Detailing    state:approved
'   "fraza ze spacja"      - cudzyslow grupuje                        filename:"my file*"
'   -term                  - wyklucz                                 -pdf   -path:Templates
'   kilka terminow         - wszystkie musza pasowac (AND)
' Aliasy kolumn: filename/name/file, path/folder, link/url, updater/user, updated/date, status; oraz nazwa dowolnej kolumny bez spacji (np. workflowstate)
' =======================================================================
Public gData As Variant                             ' tablica 2D (1..gN, 1..gC) z PW_Data (od wiersza 12)
Public gHdr() As String                             ' naglowki kolumn (1..gC)
Public gN As Long, gC As Long
Public gLoaded As Boolean
Public gIdx() As Long                               ' wynik ostatniego zapytania: numery wierszy gData
Public gHits As Long
Private gHay() As String                            ' (leniwie) polaczony tekst wiersza malymi literami - do szybkiego szukania
Private gHayOk As Boolean
Private mCI As Object                               ' nazwa kolumny (male litery bez spacji) -> numer

' Czyta blok komorek zawsze jako tablice 2D (1..r, 1..c)
Public Function ReadBlock(ws As Worksheet, ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long) As Variant
    Dim v As Variant, a() As Variant
    v = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value2
    If IsArray(v) Then
        ReadBlock = v
    Else
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = v
        ReadBlock = a
    End If
End Function

Public Function NormName(ByVal s As String) As String
    NormName = LCase(Replace(Replace(Replace(Trim(s), " ", ""), "_", ""), "-", ""))
End Function

Public Sub InvalidateData()
    gLoaded = False: gHayOk = False: gN = 0
End Sub

Private Sub BuildColIndex()
    Dim i As Long
    Set mCI = CreateObject("Scripting.Dictionary")
    mCI.CompareMode = 1
    For i = 1 To gC
        If gHdr(i) <> "" Then mCI(NormName(gHdr(i))) = i
    Next i
End Sub

' Kolumna po (znormalizowanej) nazwie naglowka; 0 = brak
Public Function ColOf(ByVal nm As String) As Long
    EnsureData
    If mCI Is Nothing Then Exit Function
    If mCI.exists(NormName(nm)) Then ColOf = mCI(NormName(nm))
End Function

Public Sub EnsureData()
    If gLoaded Then Exit Sub
    Dim ws As Worksheet, lastC As Long, lastR As Long, hv As Variant, i As Long
    Set ws = ThisWorkbook.Sheets(SH_DATA)
    gN = 0: gC = 0: gHayOk = False
    lastC = ws.Cells(DATA_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If Trim(CStr(ws.Cells(DATA_HEADER_ROW, lastC).Value)) = "" Then lastC = 0
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastC > 0 Then
        hv = ReadBlock(ws, DATA_HEADER_ROW, 1, DATA_HEADER_ROW, lastC)
        ReDim gHdr(1 To lastC)
        For i = 1 To lastC: gHdr(i) = Trim(CStr(hv(1, i))): Next i
        gC = lastC
        If lastR > DATA_HEADER_ROW Then
            gData = ReadBlock(ws, DATA_HEADER_ROW + 1, 1, lastR, lastC)
            gN = UBound(gData, 1)
        End If
    Else
        ReDim gHdr(1 To 1)
    End If
    BuildColIndex
    gLoaded = True
End Sub

' Ustawia pamiec podreczna prosto z tablicy wyniku importu (bez ponownego czytania arkusza)
Public Sub SetDataCache(res As Variant, hdr() As String, ByVal n As Long, ByVal c As Long)
    Dim i As Long, j As Long
    ReDim gHdr(1 To c)
    For i = 1 To c: gHdr(i) = hdr(i): Next i
    gC = c: gN = n
    If n > 0 Then
        Dim a() As Variant
        ReDim a(1 To n, 1 To c)
        For i = 1 To n
            For j = 1 To c: a(i, j) = res(i, j): Next j
        Next i
        gData = a
    End If
    gHayOk = False
    BuildColIndex
    gLoaded = True
End Sub

Private Sub BuildHay()
    If gHayOk Then Exit Sub
    Dim i As Long, j As Long, s As String
    If gN = 0 Then gHayOk = True: Exit Sub
    ReDim gHay(1 To gN)
    For i = 1 To gN
        s = ""
        For j = 1 To gC
            If Not IsError(gData(i, j)) Then s = s & CStr(gData(i, j)) & Chr(1)
        Next j
        gHay(i) = LCase(s)
    Next i
    gHayOk = True
End Sub

Private Function EscLike(ByVal s As String) As String
    s = Replace(s, "[", "[[]")
    EscLike = Replace(s, "#", "[#]")
End Function

Private Function ResolveField(ByVal f As String) As Long    ' >0 kolumna, 0 = dowolna, -1 = nieznane
    Dim k As String, key As Variant, cand As Long, hits As Long
    k = NormName(f)
    Select Case k
        Case "all", "any": ResolveField = 0: Exit Function
        Case "filename", "name", "file": k = "filename"
        Case "path", "folder", "sourcepath": k = "pwsourcepath"
        Case "link", "url", "urn": k = "pwlink"
        Case "user": k = "updater"
        Case "date": k = "updated"
    End Select
    If mCI.exists(k) Then ResolveField = mCI(k): Exit Function
    For Each key In mCI.Keys                        ' przedrostek jednoznaczny
        If Left(CStr(key), Len(k)) = k Then cand = mCI(key): hits = hits + 1
    Next key
    If hits = 1 Then ResolveField = cand: Exit Function
    hits = 0                                        ' brak przedrostka - sprobuj jednoznacznego fragmentu nazwy (np. state -> WorkflowState)
    For Each key In mCI.Keys
        If InStr(CStr(key), k) > 0 Then cand = mCI(key): hits = hits + 1
    Next key
    If hits = 1 Then ResolveField = cand Else ResolveField = -1
End Function

Private Function Tokenize(ByVal s As String) As Collection
    Dim col As New Collection, i As Long, ch As String, cur As String, q As Boolean
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = """" Then
            q = Not q
        ElseIf ch = " " And Not q Then
            If cur <> "" Then col.Add cur: cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    If cur <> "" Then col.Add cur
    Set Tokenize = col
End Function

' Wykonuje zapytanie: folder (klucz z koncowym \, "" = wszystko) + fraza; wynik w gIdx(1..gHits); zwraca liczbe wynikow
Public Function RunQuery(ByVal srch As String, ByVal folder As String) As Long
    EnsureData
    gHits = 0
    If gN = 0 Then ReDim gIdx(1 To 1): RunQuery = 0: Exit Function
    ReDim gIdx(1 To gN)
    If LCase(Trim(srch)) = "batch:" Then RunQuery = BatchQuery(): Exit Function     ' szukanie zbiorcze - lista nazw z arkusza Batch
    Dim i As Long, j As Long
    Dim iPath As Long: iPath = ColOf("PW Source Path")
    
    ' --- folder: pliki wprost w folderze; jesli ich brak - cale poddrzewo ---
    Dim inF() As Byte, f As String, pk As String, direct As Long, sub1 As Long
    ReDim inF(1 To gN)
    f = LCase(Replace(Trim(folder), "/", "\"))
    If Trim(srch) <> "" Then f = ""                  ' szukanie dziala zawsze po calej bazie (filtr folderu jest pomijany)
    If f <> "" Then
        If Right(f, 1) <> "\" Then f = f & "\"
        For i = 1 To gN
            pk = LCase(Replace(CStr(gData(i, iPath)), "/", "\"))
            If Right(pk, 1) <> "\" Then pk = pk & "\"
            If pk = f Then
                inF(i) = 2: direct = direct + 1
            ElseIf Left(pk, Len(f)) = f Then
                inF(i) = 1: sub1 = sub1 + 1
            End If
        Next i
        If direct > 0 Then
            For i = 1 To gN
                If inF(i) = 1 Then inF(i) = 0
            Next i
        End If
    Else
        For i = 1 To gN: inF(i) = 1: Next i
    End If
    
    ' --- terminy wyszukiwania ---
    Dim nt As Long, tCol() As Long, tPat() As String, tWild() As Boolean, tNeg() As Boolean
    ReDim tCol(1 To 20): ReDim tPat(1 To 20): ReDim tWild(1 To 20): ReDim tNeg(1 To 20)
    Dim toks As Collection, k As Long, t As String, neg As Boolean, cpos As Long, fld As String, vtxt As String, fc As Long
    Set toks = Tokenize(Trim(srch))
    k = 1
    Do While k <= toks.Count And nt < 20
        t = toks(k): neg = False
        If Left(t, 1) = "-" And Len(t) > 1 Then neg = True: t = Mid(t, 2)
        cpos = InStr(t, ":")
        fc = 0
        If cpos > 1 Then
            fld = Left(t, cpos - 1): vtxt = Mid(t, cpos + 1)
            fc = ResolveField(fld)
            If fc >= 0 Then
                If vtxt = "" And k < toks.Count Then k = k + 1: vtxt = toks(k)
            Else
                fc = 0: vtxt = t                     ' np. pw://... - zwykly tekst
            End If
        Else
            vtxt = t
        End If
        If vtxt <> "" Then
            nt = nt + 1
            tCol(nt) = fc: tNeg(nt) = neg
            tPat(nt) = LCase(vtxt)
            tWild(nt) = (InStr(vtxt, "*") > 0 Or InStr(vtxt, "?") > 0)
            If tWild(nt) Then tPat(nt) = EscLike(tPat(nt))
        End If
        k = k + 1
    Loop
    
    Dim needHay As Boolean
    For j = 1 To nt
        If tCol(j) = 0 Then needHay = True
    Next j
    If needHay Then BuildHay
    
    Dim ok As Boolean, hit As Boolean, v As String
    For i = 1 To gN
        If inF(i) <> 0 Then
            ok = True
            For j = 1 To nt
                If tCol(j) = 0 Then
                    If tWild(j) Then hit = (gHay(i) Like "*" & tPat(j) & "*") Else hit = (InStr(gHay(i), tPat(j)) > 0)
                Else
                    v = LCase(CStr(gData(i, tCol(j))))
                    If tWild(j) Then hit = (v Like tPat(j)) Else hit = (InStr(v, tPat(j)) > 0)
                End If
                If tNeg(j) Then hit = Not hit
                If Not hit Then ok = False: Exit For
            Next j
            If ok Then gHits = gHits + 1: gIdx(gHits) = i
        End If
    Next i
    RunQuery = gHits
End Function

' Statystyki calej bazy do kart KPI
Public Sub DataStats(ByRef nFolders As Long, ByRef nFiles As Long, ByRef nChg As Long, ByRef nDel As Long)
    EnsureData
    nFiles = gN: nFolders = 0: nChg = 0: nDel = 0
    If gN = 0 Then Exit Sub
    Dim d As Object, i As Long, iP As Long, iSt As Long, s As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    iP = ColOf("PW Source Path"): iSt = ColOf(COL_STATUS)
    For i = 1 To gN
        d(CStr(gData(i, iP))) = 1
        If iSt > 0 Then
            s = CStr(gData(i, iSt))
            If s = "NEW" Or s = "CHANGED" Then nChg = nChg + 1
            If s = "DELETED" Then nDel = nDel + 1
        End If
    Next i
    nFolders = d.Count
End Sub
