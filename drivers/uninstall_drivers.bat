@echo off
setlocal

echo.
echo ******************************
echo * Uninstalling MOXA and PEAK *
echo ******************************
echo.

REM Log file to capture output and errors
set logfile=%temp%\uninstall_log.txt
echo Logging to %logfile%

REM Uninstall all software provided by MOXA Inc.
echo Searching for MOXA Inc. software...
powershell -command "Get-WmiObject -Query \"SELECT * FROM Win32_Product WHERE Vendor LIKE 'Moxa Inc.%'\" | ForEach-Object { $_.Uninstall() }" >> %logfile% 2>&1

REM Uninstall all software provided by PEAK-System Technik GmbH
echo Searching for PEAK-System Technik GmbH software...
powershell -command "Get-WmiObject -Query \"SELECT * FROM Win32_Product WHERE Vendor LIKE 'PEAK-System Technik GmbH%'\" | ForEach-Object { $_.Uninstall() }" >> %logfile% 2>&1

echo.
echo ******************************
echo * Uninstallation Completed   *
echo ******************************
echo.

echo Please review the log file at %logfile% if there were any issues.
pause
endlocal
