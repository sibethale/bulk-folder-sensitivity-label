@echo off
REM Launches Set-FolderSensitivityLabel.ps1 under Windows PowerShell 5.1, which is the only
REM host where the Purview Information Protection client module loads correctly.
REM All arguments are forwarded, e.g.:
REM     Run-LabelChanger.cmd -Path "C:\Data\Reports" -Recurse -ShowCurrentLabels
powershell.exe -NoProfile -ExecutionPolicy Bypass -NoLogo -Command "& '%~dp0Set-FolderSensitivityLabel.ps1' %*"
