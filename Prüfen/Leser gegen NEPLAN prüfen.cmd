@echo off
rem Prueft den Export ohne NEPLAN gegen einen echten NEPLAN-Export.
rem Eine Datenbank, die NEPLAN exportiert hat, auf diese Datei ziehen
rem (oder doppelklicken und die Datenbank auswaehlen).
set "PS=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0..\werkzeug\leser_pruefen.ps1" "%~1"
echo.
pause
