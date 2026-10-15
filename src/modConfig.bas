Attribute VB_Name = "modConfig"
Option Explicit

' =======================================================================
' KONFIGURACJA SKOROSZYTU - jedno miejsce na nazwy arkuszy, układ i kolory.
' Zmiana nazwy zakładki w Excelu wymaga zmiany TUTAJ (i tylko tutaj).
' =======================================================================

' --- Wersja kodu: hash commita z repozytorium, z ktorego zaimportowano kod (widac go w arkuszu Instrukcja, B3) ---
Public Const BUILD_ID As String = "432f6a8"

' --- Nazwy arkuszy (zakładek) ---
Public Const SH_START As String = "Guide"            ' User guide
Public Const SH_FOLDERS As String = "Folders"       ' Lista folderów ProjectWise do skanowania
Public Const SH_ATTRS As String = "Attributes"     ' Lista atrybutów (kolumn) do pobrania
Public Const SH_DATA As String = "PW_Data"          ' Wynik: lista plików z ProjectWise
Public Const SH_DIAG As String = "Diagnostics"      ' Podgląd dostępnych atrybutów PW
Public Const SH_LOG As String = "Log"               ' Historia zmian
Public Const SH_VIEWS As String = "Views"            ' Definicje widokow (ktore kolumny i w jakiej kolejnosci)
Public Const SH_VIEW As String = "Dashboard"       ' Wynik: Baza_PW wyswietlona wg wybranego widoku

' --- Układ arkusza "Foldery": A = Aktywny (TAK/NIE), B = link PW, C = opis ---
Public Const FOLDERS_FIRST_ROW As Long = 6
Public Const FOLDERS_ACTIVE_COL As Long = 1
Public Const FOLDERS_LINK_COL As Long = 2

' --- Uklad arkusza "Atrybuty": A = nazwa, B = Pobierac (TAK/NIE), C = typ, D = przyklad, E = opis ---
Public Const ATTRS_FIRST_ROW As Long = 6
Public Const ATTRS_ENABLE_COL As Long = 2
Public Const ATTRS_TYPE_COL As Long = 3
Public Const ATTRS_EXAMPLE_COL As Long = 4
Public Const ATTRS_SYSTEM_DEFAULTS As String = "Version,VersionSequence,WorkflowState,DocumentStatus,FileSize,DocumentCreatorName,CreateDate" ' wlasciwosci systemowe domyslnie TAK

' --- Automat: po wklejeniu linku w arkuszu Foldery wykryj atrybuty, a potem pobierz dane ---
Public Const AUTO_DOWNLOAD_AFTER_PASTE As Boolean = True

' --- Widoki: arkusz Widoki (A = nazwa, B = kolumny, C = opis; od wiersza 6), arkusz Widok (wybor w B3, szukaj w B4, tabela od wiersza 8) ---
Public Const VIEWS_FIRST_ROW As Long = 6
Public Const VIEW_SELECTOR As String = "B5"
Public Const VIEW_HEADER_ROW As Long = 8          ' naglowek tabeli w arkuszu Widok; dane od wiersza 9
Public Const VIEW_SEARCH As String = "D5"           ' pole Szukaj w arkuszu Widok

' --- Układ arkusza "Baza_PW" ---
Public Const DATA_HEADER_ROW As Long = 11           ' Wiersz nagłówków; dane od następnego wiersza

' --- Pobieranie ---
Public Const FAST_FETCH As Boolean = True           ' True = szybkie pobieranie (DataTable, sekundy); przy bledzie samo przechodzi na wolny tryb
Public Const GET_ALL_VERSIONS As Boolean = True     ' True = kazda rewizja pliku (A, B, C...) w osobnym wierszu; False = tylko biezaca

' --- Kolory (wartości RGB policzone: R + G*256 + B*65536) ---
Public Const CLR_HEADER As Long = 6567967          ' granat  RGB(31,56,100)  - nagłówki
Public Const CLR_FOLDER As Long = 16247773         ' jasny niebieski RGB(221,235,247) - wiersz folderu
Public Const CLR_CHANGED As Long = 10284031        ' jasny żółty RGB(255,235,156) - zmieniona wartość
Public Const CLR_DELETED As Long = 13551615        ' jasny czerwony RGB(255,199,206) - plik usunięty z PW
Public Const CLR_WHITE As Long = 16777215
Public Const CLR_FOLDER_DARK As Long = 6044710     ' RGB(38,60,92)    - folder w motywie Ciemny
Public Const CLR_FOLDER_BLUE As Long = 16114125    ' RGB(205,225,245) - folder w motywie Niebieski
Public Const TARGET_CELL As String = "H5"          ' folder docelowy pobierania (Dashboard)
Public Const THEME_CELL As String = "F5"           ' wybor motywu w arkuszu Widok (Jasny / Ciemny / Niebieski)

' Czy kolor to tlo wiersza folderu w ktoryms z motywow (wykrywanie folderow nie zalezy od motywu)
Public Function IsFolderColor(ByVal c As Long) As Boolean
    IsFolderColor = (c = CLR_FOLDER Or c = CLR_FOLDER_DARK Or c = CLR_FOLDER_BLUE)
End Function

' Komunikat bez okna (nie blokuje Excela): pasek stanu + Immediate
Public Sub Notify(ByVal msg As String)
    Application.StatusBar = Left(Replace(msg, vbCrLf, " "), 250)
    Debug.Print msg
End Sub

' Wartosc wylaczona w kolumnach YES/NO (akceptuje tez stare polskie NIE)
Public Function IsNo(ByVal v As Variant) As Boolean
    Dim t As String
    t = UCase(Trim(CStr(v)))
    IsNo = (t = "NO" Or t = "NIE")
End Function
