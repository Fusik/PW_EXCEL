@echo off
setlocal
chcp 65001 >nul
rem ==========================================================================
rem  Pobiera najnowsza wersje repozytorium fusik/pw_excel do folderu, w ktorym
rem  lezy ten plik .bat. Istniejace pliki o tych samych nazwach sa NADPISYWANE
rem  (zrob kopie PW_integrationg.xlsm, jesli masz w nim wlasne dane).
rem ==========================================================================
set "REPO=fusik/pw_excel"
set "BRANCH=claude/ecstatic-planck-3kaiet"
set "DEST=%~dp0"
set "TMP_DIR=%TEMP%\pw_excel_download"

echo Repozytorium: %REPO%  (galaz: %BRANCH%)
echo Folder docelowy: %DEST%
echo.
echo UWAGA: pliki w tym folderze zostana nadpisane wersja z GitHuba.
choice /m "Kontynuowac"
if errorlevel 2 exit /b 0

if exist "%TMP_DIR%" rmdir /s /q "%TMP_DIR%"

where git >nul 2>nul
if %errorlevel%==0 (
    echo [git] Pobieranie...
    git clone --depth 1 --branch "%BRANCH%" "https://github.com/%REPO%.git" "%TMP_DIR%"
    if errorlevel 1 goto :fail
    rmdir /s /q "%TMP_DIR%\.git"
) else (
    echo [git niedostepny] Pobieranie archiwum ZIP przez PowerShell - dziala tylko dla repozytorium publicznego...
    mkdir "%TMP_DIR%"
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest 'https://github.com/%REPO%/archive/refs/heads/%BRANCH%.zip' -OutFile '%TMP_DIR%\repo.zip'; Expand-Archive '%TMP_DIR%\repo.zip' -DestinationPath '%TMP_DIR%\unzip' -Force; Move-Item (Get-ChildItem '%TMP_DIR%\unzip' | Select-Object -First 1).FullName '%TMP_DIR%\content'"
    if errorlevel 1 goto :fail
    rename "%TMP_DIR%" "pw_excel_download_zip" 2>nul
    set "TMP_DIR=%TEMP%\pw_excel_download_zip\content"
)

echo Kopiowanie plikow do %DEST% ...
xcopy "%TMP_DIR%\*" "%DEST%" /E /I /Y /Q >nul
if errorlevel 1 goto :fail

rmdir /s /q "%TMP_DIR%" 2>nul
rmdir /s /q "%TEMP%\pw_excel_download_zip" 2>nul
echo.
echo Gotowe.
pause
exit /b 0

:fail
echo.
echo BLAD: nie udalo sie pobrac repozytorium. Jesli jest prywatne, zainstaluj Git
echo i zaloguj sie do GitHuba (git ssam zapyta o dane logowania).
pause
exit /b 1
