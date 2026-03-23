@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Printer Audit and Cleanup Utility

color 09
chcp 65001 >nul

:: Check for administrator rights
net session >nul 2>&1
if %errorlevel%==0 (
  set "ISADMIN=1"
) else (
  set "ISADMIN=0"
)

:: Report folder on desktop
set "REPORTROOT=%USERPROFILE%\Desktop\PrinterReports"
if not exist "%REPORTROOT%" md "%REPORTROOT%" >nul 2>&1

:: Neutral search term placeholder
set "SEARCHTERM="

:MAIN
cls
echo ============================================================
echo Printer Audit and Cleanup by complicatiion
echo ============================================================
echo.
if "%ISADMIN%"=="1" (
  echo Admin status: YES
) else (
  echo Admin status: NO
)
echo Report folder: %REPORTROOT%
if defined SEARCHTERM (
  echo Search term: %SEARCHTERM%
) else (
  echo Search term: ^<nicht gesetzt^>
)
echo.
echo [1] List installed printers
echo [2] List printer ports
echo [3] List printer drivers
echo [4] Search printer / port / driver
echo [5] Check print server connections
echo [6] Check registry connections
echo [7] Check spooler status
echo [8] Generate full report
echo [9] Enter printer name / port / IP / driver term
echo [A] Delete printer by exact name                 [Admin]
echo [B] Delete printer port by exact name             [Admin]
echo [C] Delete network printer connection by UNC
echo [D] Delete computer printer connection by UNC        [Admin]
echo [E] Restart spooler                     [Admin]
echo [F] Report folder oeffnen
echo [0] Exit
echo.
set /p CHO="Selection: "

if "%CHO%"=="1" goto :LISTPRINTERS
if "%CHO%"=="2" goto :LISTPORTS
if "%CHO%"=="3" goto :LISTDRIVERS
if "%CHO%"=="4" goto :SEARCH
if "%CHO%"=="5" goto :SERVERCONN
if "%CHO%"=="6" goto :REGCHECK
if "%CHO%"=="7" goto :SPOOLER
if "%CHO%"=="8" goto :REPORT
if "%CHO%"=="9" goto :SETSEARCH
if /I "%CHO%"=="A" goto :DELPRINTER
if /I "%CHO%"=="B" goto :DELPORT
if /I "%CHO%"=="C" goto :DELUNCUSER
if /I "%CHO%"=="D" goto :DELUNCCOMP
if /I "%CHO%"=="E" goto :RESTARTSPOOLER
if /I "%CHO%"=="F" goto :OPENFOLDER
if "%CHO%"=="0" goto :END
goto :MAIN

:SETSEARCH
cls
echo ============================================================
echo Search term festlegen
echo ============================================================
echo.
echo Examples:
echo - ND8007
echo - ND 8007
echo - HP
echo - Lexmark
echo - 10.10.10.50
echo - PRINTSERVER01
echo.
set /p SEARCHTERM="Enter printer name, port, IP or driver term: "
goto :MAIN

:LISTPRINTERS
cls
echo ============================================================
echo Installed printers
echo ============================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Printer -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object Name, ComputerName, Type, DriverName, PortName, Shared, Published, Default, Comment | Format-Table -AutoSize"
echo.
pause
goto :MAIN

:LISTPORTS
cls
echo ============================================================
echo Printer ports
echo ============================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-PrinterPort -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object Name, PrinterHostAddress, PortNumber, PortMonitor, Description | Format-Table -AutoSize"
echo.
pause
goto :MAIN

:LISTDRIVERS
cls
echo ============================================================
echo Printer drivers
echo ============================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-PrinterDriver -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object Name, MajorVersion, DriverVersion, InfPath, Manufacturer | Format-Table -AutoSize"
echo.
pause
goto :MAIN

:SEARCH
cls
echo ============================================================
echo Search printer / port / driver
echo ============================================================
echo.
if not defined SEARCHTERM (
  echo Es ist noch kein Search term gesetzt.
  echo Please use option [9] first.
  echo.
  pause
  goto :MAIN
)
echo Search term: %SEARCHTERM%
echo.
echo [1] Matching printers
powershell -NoProfile -ExecutionPolicy Bypass -Command "$term='%SEARCHTERM%'; Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ('*' + $term + '*') -or $_.PortName -like ('*' + $term + '*') -or $_.DriverName -like ('*' + $term + '*') -or $_.Comment -like ('*' + $term + '*') } | Select-Object Name, Type, DriverName, PortName, Shared, Default, Comment | Format-Table -AutoSize"
echo.
echo [2] Matching ports
powershell -NoProfile -ExecutionPolicy Bypass -Command "$term='%SEARCHTERM%'; Get-PrinterPort -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ('*' + $term + '*') -or $_.PrinterHostAddress -like ('*' + $term + '*') -or $_.Description -like ('*' + $term + '*') } | Select-Object Name, PrinterHostAddress, PortNumber, PortMonitor, Description | Format-Table -AutoSize"
echo.
echo [3] Matching drivers
powershell -NoProfile -ExecutionPolicy Bypass -Command "$term='%SEARCHTERM%'; Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ('*' + $term + '*') -or $_.Manufacturer -like ('*' + $term + '*') -or $_.InfPath -like ('*' + $term + '*') } | Select-Object Name, MajorVersion, DriverVersion, InfPath, Manufacturer | Format-Table -AutoSize"
echo.
pause
goto :MAIN

:SERVERCONN
cls
echo ============================================================
echo Check print server connections
echo ============================================================
echo.
echo [1] Installed printers mit Serverbezug
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.ComputerName -or $_.Type -match 'Connection' } | Sort-Object Name | Select-Object Name, ComputerName, Type, DriverName, PortName | Format-Table -AutoSize"
echo.
if defined SEARCHTERM (
  echo [2] Search term in Druckserver-Verbindungen
  powershell -NoProfile -ExecutionPolicy Bypass -Command "$term='%SEARCHTERM%'; Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ('*' + $term + '*') -or $_.ComputerName -like ('*' + $term + '*') -or $_.PortName -like ('*' + $term + '*') } | Select-Object Name, ComputerName, Type, DriverName, PortName | Format-Table -AutoSize"
  echo.
)
echo [3] Note
echo Printers deployed centrally by print server or GPO may reappear after deletion.
echo In that case, the source must be cleaned on print server or GPO level.
echo.
pause
goto :MAIN

:REGCHECK
cls
echo ============================================================
echo Check registry connections
echo ============================================================
echo.
echo [1] User-based printer connections
reg query "HKCU\Printers\Connections"
echo.
echo [2] Computer-based printer connections
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections"
echo.
echo [3] Local printer objects
reg query "HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers"
echo.
pause
goto :MAIN

:SPOOLER
cls
echo ============================================================
echo Check spooler status
echo ============================================================
echo.
sc query spooler
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Service spooler | Select-Object Name, Status, StartType | Format-Table -AutoSize"
echo.
pause
goto :MAIN

:DELPRINTER
if not "%ISADMIN%"=="1" goto :NEEDADMIN
cls
echo ============================================================
echo Delete printer by exact name
echo ============================================================
echo.
set /p DELPRN="Enter exact printer name: "
if "%DELPRN%"=="" goto :MAIN
echo.
echo Target: %DELPRN%
set /p CONFIRM="Type YES to confirm: "
if /I not "%CONFIRM%"=="JA" goto :MAIN
powershell -NoProfile -ExecutionPolicy Bypass -Command "Remove-Printer -Name '%DELPRN%' -ErrorAction Stop; Write-Host 'Printer was deleted.'"
echo.
pause
goto :MAIN

:DELPORT
if not "%ISADMIN%"=="1" goto :NEEDADMIN
cls
echo ============================================================
echo Delete printer port by exact name
echo ============================================================
echo.
set /p DELPORTNAME="Enter exact port name: "
if "%DELPORTNAME%"=="" goto :MAIN
echo.
echo Target: %DELPORTNAME%
set /p CONFIRM="Type YES to confirm: "
if /I not "%CONFIRM%"=="JA" goto :MAIN
powershell -NoProfile -ExecutionPolicy Bypass -Command "Remove-PrinterPort -Name '%DELPORTNAME%' -ErrorAction Stop; Write-Host 'Port was deleted.'"
echo.
pause
goto :MAIN

:DELUNCUSER
cls
echo ============================================================
echo Delete network printer connection by UNC
echo ============================================================
echo.
echo Example: \\PRINTSERVER\PrinterName
set /p UNCNAME="Enter UNC printer path: "
if "%UNCNAME%"=="" goto :MAIN
echo.
echo Target: %UNCNAME%
set /p CONFIRM="Type YES to confirm: "
if /I not "%CONFIRM%"=="JA" goto :MAIN
rundll32 printui.dll,PrintUIEntry /dn /n%UNCNAME%
echo.
echo Command executed.
echo.
pause
goto :MAIN

:DELUNCCOMP
if not "%ISADMIN%"=="1" goto :NEEDADMIN
cls
echo ============================================================
echo Delete computer printer connection by UNC
echo ============================================================
echo.
echo Example: \\PRINTSERVER\PrinterName
set /p UNCNAME="Enter UNC printer path: "
if "%UNCNAME%"=="" goto :MAIN
echo.
echo Target: %UNCNAME%
set /p CONFIRM="Type YES to confirm: "
if /I not "%CONFIRM%"=="JA" goto :MAIN
rundll32 printui.dll,PrintUIEntry /gd /n%UNCNAME%
echo.
echo Command executed.
echo.
pause
goto :MAIN

:RESTARTSPOOLER
if not "%ISADMIN%"=="1" goto :NEEDADMIN
cls
echo ============================================================
echo Restart spooler
echo ============================================================
echo.
net stop spooler
net start spooler
echo.
pause
goto :MAIN

:REPORT
cls
echo [*] Creating report...
echo.
for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set STAMP=%%I
set "OUTFILE=%REPORTROOT%\Printer_Audit_Report_%STAMP%.txt"

(
echo ============================================================
echo Printer Audit and Cleanup Report
echo ============================================================
echo Date: %DATE% %TIME%
echo Computer: %COMPUTERNAME%
echo User: %USERNAME%
echo Admin: %ISADMIN%
if defined SEARCHTERM (echo Search term: %SEARCHTERM%) else (echo Search term: not set)
echo ============================================================
echo.

echo [1] Installed printers
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Printer -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object Name, ComputerName, Type, DriverName, PortName, Shared, Published, Default, Comment | Format-Table -AutoSize"
echo.

if defined SEARCHTERM (
echo [2] Matching printers for search term
powershell -NoProfile -ExecutionPolicy Bypass -Command "$term='%SEARCHTERM%'; Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ('*' + $term + '*') -or $_.PortName -like ('*' + $term + '*') -or $_.DriverName -like ('*' + $term + '*') -or $_.Comment -like ('*' + $term + '*') } | Select-Object Name, ComputerName, Type, DriverName, PortName, Shared, Default, Comment | Format-Table -AutoSize"
echo.
)

echo [3] Installed printer ports
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-PrinterPort -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object Name, PrinterHostAddress, PortNumber, PortMonitor, Description | Format-Table -AutoSize"
echo.

if defined SEARCHTERM (
echo [4] Matching ports for search term
powershell -NoProfile -ExecutionPolicy Bypass -Command "$term='%SEARCHTERM%'; Get-PrinterPort -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ('*' + $term + '*') -or $_.PrinterHostAddress -like ('*' + $term + '*') -or $_.Description -like ('*' + $term + '*') } | Select-Object Name, PrinterHostAddress, PortNumber, PortMonitor, Description | Format-Table -AutoSize"
echo.
)

echo [5] Installed printer drivers
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-PrinterDriver -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object Name, MajorVersion, DriverVersion, InfPath, Manufacturer | Format-Table -AutoSize"
echo.

if defined SEARCHTERM (
echo [6] Matching drivers for search term
powershell -NoProfile -ExecutionPolicy Bypass -Command "$term='%SEARCHTERM%'; Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ('*' + $term + '*') -or $_.Manufacturer -like ('*' + $term + '*') -or $_.InfPath -like ('*' + $term + '*') } | Select-Object Name, MajorVersion, DriverVersion, InfPath, Manufacturer | Format-Table -AutoSize"
echo.
)

echo [7] Printer server connections
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.ComputerName -or $_.Type -match 'Connection' } | Sort-Object Name | Select-Object Name, ComputerName, Type, DriverName, PortName | Format-Table -AutoSize"
echo.

echo [8] Spooler service
sc query spooler
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Service spooler | Select-Object Name, Status, StartType | Format-Table -AutoSize"
echo.

echo [9] User printer connections in registry
reg query "HKCU\Printers\Connections"
echo.

echo [10] Computer printer connections in registry
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections"
echo.

echo [11] Local printer objects in registry
reg query "HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers"
echo.

echo [12] Recommendations
echo - Search for ghost printers by exact name, port and driver.
echo - Remove stale queue first, then stale port if needed.
echo - If printer is deployed by print server or GPO, clean the central source as well.
echo - Restart spooler after cleanup if objects still appear inconsistent.
echo.
) > "%OUTFILE%" 2>&1

echo Report created:
echo %OUTFILE%
echo.
pause
goto :MAIN

:OPENFOLDER
start "" explorer.exe "%REPORTROOT%"
goto :MAIN

:NEEDADMIN
cls
echo Administrator rights are required for this action.
echo.
pause
goto :MAIN

:END
endlocal
exit /b
