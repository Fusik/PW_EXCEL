# PW_EXCEL – eksport zawartości ProjectWise do Excela

`PW_integrationg.xlsm` pobiera listę plików (z atrybutami) z folderów ProjectWise przez moduł PowerShell `PWPS_DAB`
i utrzymuje ją w `Sheet_1` jako **deltę** (nowe / zmienione / usunięte pliki).

## Jak to działa
1. `Sheet_2` kol. B – linki `pw:\\...` do skanowanych folderów.
2. `Sheet_3` kol. A – nazwy atrybutów PW do pobrania (wiersz 1 = nagłówek).
3. Makro `DownloadFromProjectWise_PWPS` (Module1) generuje `PW_GetFiles.ps1`, uruchamia je, wczytuje CSV i porównuje z `Sheet_1`
   (od wiersza 12; nagłówki w wierszu 11; zielone wiersze = foldery).
   - nowy plik → dodany + wpis `ADDED` w `Log`
   - zmiana atrybutu → żółta komórka, komentarz ze starą wartością, wpis `CHANGED`
   - plik zniknął z PW → czerwony wiersz + przekreślenie, wpis `DELETED`
4. `PW_Diagnostyka` / `PW_diagnostyka` – sprawdzenie środowiska, `ZbadajAtrybutyPW` – podgląd dostępnych atrybutów.

## Kod źródłowy (`src/`)
Kod VBA wyciągnięty z xlsm, żeby dało się go przeglądać w diffach. Moduły standardowe (`Module1-3`) wgrywa do skoroszytu
`Import-VBA.ps1` (Windows + Excel, włączony dostęp do modelu VBA). Formularze i kod arkuszy (`*.frm`, `*.cls`) są tylko do podglądu.

## Zmiany w tej wersji (V1.3)
- **Bezpieczne usuwanie**: PowerShell zapisuje `[SCOPE_OK]` dla folderów faktycznie przeskanowanych. Wcześniej błąd logowania/nieznaleziony
  folder powodował oznaczenie WSZYSTKICH jego plików jako usunięte. Teraz takie foldery są pomijane i logowane jako `WARNING`.
- **Brak starych danych**: przed uruchomieniem PS kasowane są stare `PW_Output.csv` i log struktury (przy awarii PS delta nie liczy się już na starych plikach).
- **Odporność makra**: obsługa błędów, wyłączenie `ScreenUpdating/EnableEvents/Calculation` na czas pracy (zawsze przywracane), błąd trafia do `Log`.
- `CopyHeadersFromSheet3`: nie nadpisuje już nagłówków pustymi komórkami i kopiuje tylko faktyczne atrybuty (kolejność = CSV).
- `RemoveEmptyRows`: pyta tylko, gdy są puste wiersze, i podaje ich liczbę.
- Wykluczany użytkownik (`IGNORED_UPDATER`) wyniesiony do stałej na górze modułu.

> Zmiany w VBA nie były uruchamiane w Excelu (środowisko bez Excela/ProjectWise) – przetestuj na kopii pliku.

## Pomysły na kolejne kroki
- Wstawianie wierszy partiami zamiast `Rows.Insert` per plik (główne wąskie gardło przy tysiącach plików).
- Konfiguracja (wykluczeni użytkownicy, wiersz nagłówków, kolory) w osobnym arkuszu `Config`.
- Pasek postępu w Excelu i odczyt kodu wyjścia PowerShella; usunięcie `Start-Sleep 10` i oddzielne okno tylko przy błędzie.
- Dopasowanie zakresu usunięć po prefiksie ścieżki zamiast `InStr` (dziś podfolder o podobnej nazwie może dać fałszywe trafienie).
- Obcięcie komentarzy-historii w komórkach (rosną bez limitu).
