@echo off
rem Uruchamia Import-VBA.ps1 z pominieciem blokady skryptow (tylko dla tego jednego uruchomienia).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Import-VBA.ps1" %*
pause
