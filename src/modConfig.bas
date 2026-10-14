Attribute VB_Name = "modConfig"
Option Explicit

' =======================================================================
' KONFIGURACJA SKOROSZYTU - jedno miejsce na nazwy arkuszy, układ i kolory.
' Zmiana nazwy zakładki w Excelu wymaga zmiany TUTAJ (i tylko tutaj).
' =======================================================================

' --- Nazwy arkuszy (zakładek) ---
Public Const SH_START As String = "Instrukcja"      ' Instrukcja obsługi
Public Const SH_FOLDERS As String = "Foldery"       ' Lista folderów ProjectWise do skanowania
Public Const SH_ATTRS As String = "Atrybuty"        ' Lista atrybutów (kolumn) do pobrania
Public Const SH_DATA As String = "Baza_PW"          ' Wynik: lista plików z ProjectWise
Public Const SH_DIAG As String = "Diagnostyka"      ' Podgląd dostępnych atrybutów PW
Public Const SH_LOG As String = "Log"               ' Historia zmian

' --- Układ arkusza "Foldery": A = Aktywny (TAK/NIE), B = link PW, C = opis ---
Public Const FOLDERS_FIRST_ROW As Long = 6
Public Const FOLDERS_ACTIVE_COL As Long = 1
Public Const FOLDERS_LINK_COL As Long = 2

' --- Układ arkusza "Atrybuty": A = nazwa atrybutu, B = opis, C = przykład ---
Public Const ATTRS_FIRST_ROW As Long = 6
Public Const ATTRS_SAMPLE_FOLDER As String = "B3"   ' Komórka z folderem próbnym dla Diagnostyki

' --- Układ arkusza "Baza_PW" ---
Public Const DATA_HEADER_ROW As Long = 11           ' Wiersz nagłówków; dane od następnego wiersza

' --- Pobieranie ---
Public Const GET_ALL_VERSIONS As Boolean = True     ' True = kazda rewizja pliku (A, B, C...) w osobnym wierszu; False = tylko biezaca

' --- Kolory (wartości RGB policzone: R + G*256 + B*65536) ---
Public Const CLR_HEADER As Long = 6567967          ' granat  RGB(31,56,100)  - nagłówki
Public Const CLR_FOLDER As Long = 16247773         ' jasny niebieski RGB(221,235,247) - wiersz folderu
Public Const CLR_CHANGED As Long = 10284031        ' jasny żółty RGB(255,235,156) - zmieniona wartość
Public Const CLR_DELETED As Long = 13551615        ' jasny czerwony RGB(255,199,206) - plik usunięty z PW
Public Const CLR_WHITE As Long = 16777215
