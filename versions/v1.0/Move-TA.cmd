@echo off
rem Abre a tool Move TA (GUI). Requer PowerShell 5.1 (Windows) e acesso ao SQL Server do IPS.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Move-TA.ps1" %*
