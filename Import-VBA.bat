@echo off
rem Wgrywa kod VBA z src\ do PW_integration.xlsm (uruchom na kopii; Excel zamkniety).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Import-VBA.ps1" %*
pause
