# PW_EXCEL – eksport zawartości ProjectWise do Excela

`PW_integration.xlsm` pobiera listę plików (z atrybutami) z folderów ProjectWise przez moduł PowerShell `PWPS_DAB`
i utrzymuje ją w `Sheet_1` jako **deltę** (nowe / zmienione / usunięte pliki).

## Arkusze
| Arkusz | Rola |
|---|---|
| `Instrukcja` | Opis obsługi, przykłady linków do folderów PW, opis kolumn i kolorów, rozwiązywanie problemów |
| `Foldery` | Linki `pw:\\…` do skanowanych folderów (A: TAK/NIE, B: link, C: opis; dane od wiersza 6) |
| `Atrybuty` | Nazwy atrybutów PW do pobrania (kolumna A od wiersza 6) + folder próbny (B3) i przycisk „Zbadaj atrybuty” |
| `Baza_PW` | Wynik: pliki z atrybutami (nagłówki w wierszu 11, dane od 12), przyciski, legenda, wyszukiwarka |
| `Diagnostyka` | Atrybuty jednego pliku – ściąga nazw do arkusza `Atrybuty` |
| `Log` | Historia zmian (ADDED / CHANGED / DELETED / WARNING) |

Dawne nazwy: `Sheet_1`→`Baza_PW`, `Sheet_2`→`Foldery`, `Sheet_3`→`Atrybuty`, `PW_Diagnostyka`→`Diagnostyka`.
Nazwy arkuszy, układ i kolory są w jednym miejscu: `src/modConfig.bas`.

## Jak to działa
1. `Foldery` → linki, `Atrybuty` → atrybuty.
2. `AKTUALIZUJ DANE` (`DownloadFromProjectWise_PWPS`, Module1) generuje `PW_GetFiles.ps1`, uruchamia go, wczytuje CSV i porównuje z `Baza_PW`:
   - nowy plik → dodany + `ADDED` w `Log`
   - zmiana atrybutu → żółta komórka, komentarz ze starą wartością, `CHANGED`
   - plik zniknął z PW → czerwony przekreślony wiersz, `DELETED`

## Plik główny i wdrożenie
`PW_integration.xlsm` – nowy układ (V1.6). Kod VBA w środku jest jeszcze stary, więc **raz** uruchom `Import-VBA.bat`
(Excel zamknięty, na kopii pliku; wymaga włączonego „Ufaj dostępowi do modelu obiektowego projektu VBA”) – wgra `src/` do skoroszytu.

## Automat (V1.6)
- Wklejenie linku w `Foldery!B` → wykrycie atrybutów (`Diagnostyka`) → dopisanie do `Atrybuty` (TAK/NIE, typ, przykład) → pobranie danych
  (`AUTO_DOWNLOAD_AFTER_PASTE` w `modConfig`).
- `Widoki` (definicje kolumn: nazwy lub litery, np. `A, B, F, L, TK`) + `Widok` (wybór w B3; kopia `Baza_PW` w wybranej kolejności).

## Pliki w repo
- `src/` – kod VBA wyciągnięty ze skoroszytu (podgląd/diffy; formularze bez układu graficznego).
- `tools/build_workbook.py` – skrypt, który z oryginału (`tools/original/`) buduje nowy układ arkuszy na poziomie XML
  (`python3 tools/build_workbook.py tools/original/PW_integrationg_original.xlsm PW_integration.xlsm`).
- `Import-VBA.ps1` – wgrywa `src/` do skoroszytu (tylko po zmianach w kodzie).

## Zmiany w V1.4 (UI)
- Czytelne nazwy arkuszy, arkusz `Instrukcja`, nowy wygląd (nagłówki, kolory, listy TAK/NIE, walidacja linków, legenda).
- Kolumna „Aktywny” w `Foldery`; działająca wyszukiwarka w `Baza_PW`; nagłówki w `Log`.
- Kolory zmian/usunięć stonowane (stałe w `modConfig`).

## Zmiany w V1.3
- **Bezpieczne usuwanie**: PowerShell zapisuje `[SCOPE_OK]` dla folderów faktycznie przeskanowanych. Wcześniej błąd logowania/nieznaleziony
  folder powodował oznaczenie WSZYSTKICH jego plików jako usunięte. Teraz takie foldery są pomijane i logowane jako `WARNING`.
- **Brak starych danych**: przed uruchomieniem PS kasowane są stare `PW_Output.csv` i log struktury (przy awarii PS delta nie liczy się już na starych plikach).
- **Odporność makra**: obsługa błędów, wyłączenie `ScreenUpdating/EnableEvents/Calculation` na czas pracy (zawsze przywracane), błąd trafia do `Log`.
- `CopyHeadersFromSheet3`: nie nadpisuje już nagłówków pustymi komórkami i kopiuje tylko faktyczne atrybuty (kolejność = CSV).
- `RemoveEmptyRows`: pyta tylko, gdy są puste wiersze, i podaje ich liczbę.
- Usunięto wykluczanie zmian użytkownika „Starszy Specjalista ds. kadr i płac, HR, Polska” – makro przetwarza teraz zmiany wszystkich osób.

> Zmiany w VBA nie były uruchamiane w Excelu (środowisko bez Excela/ProjectWise) – przetestuj na kopii pliku.

## Pomysły na kolejne kroki
- Wstawianie wierszy partiami zamiast `Rows.Insert` per plik (główne wąskie gardło przy tysiącach plików).
- Konfiguracja (wykluczeni użytkownicy, wiersz nagłówków, kolory) w osobnym arkuszu `Config`.
- Pasek postępu w Excelu i odczyt kodu wyjścia PowerShella; usunięcie `Start-Sleep 10` i oddzielne okno tylko przy błędzie.
- Dopasowanie zakresu usunięć po prefiksie ścieżki zamiast `InStr` (dziś podfolder o podobnej nazwie może dać fałszywe trafienie).
- Obcięcie komentarzy-historii w komórkach (rosną bez limitu).

## Rewizje (V1.5)
- `GET_ALL_VERSIONS` w `src/modConfig.bas` (domyślnie `True`): wyszukiwanie dostaje `-GetVersionsToo`, więc każda rewizja pliku (A…P) trafia do osobnego wiersza.
- Starsza rewizja o tym samym `DocumentURN` dostaje w `PW Link` dopisek `?version=<VersionSequence>` (klucz musi być unikalny).
- Dopasowanie po samej nazwie pliku działa tylko dla wierszy bez linku.
- W arkuszu `Atrybuty` dodaj: `Version`, `VersionSequence` (rewizja/numer wersji), `WorkflowState` (nazwa stanu), opcjonalnie `Revision_number`. Atrybut `State` nie istnieje w tym PW.
