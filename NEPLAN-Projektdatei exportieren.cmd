@echo off
rem Der einfachste Weg: NEPLAN-Projektdatei(en) auswaehlen, je Ort entsteht eine Datenbank.
rem Braucht nur Windows. 32-Bit-PowerShell wegen der Jet-Schnittstelle fuer Access.
set "PS=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell.exe"
start "" "%PS%" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0werkzeug\exportieren.ps1"
