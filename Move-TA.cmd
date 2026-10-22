@echo off
rem Opens the Move TA tool (GUI). Requires PowerShell 5.1 (Windows) and access to the IPS SQL Server.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Move-TA.ps1" %*
