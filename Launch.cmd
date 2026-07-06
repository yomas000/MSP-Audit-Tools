@echo off
REM Launches the MSP Audit Toolkit GUI with no visible console window.
REM -STA is required for WinForms. -WindowStyle Hidden suppresses the console.
REM %~dp0 resolves to the folder this launcher is extracted into by IExpress.

powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0MSPAuditToolkit.ps1"
