@echo off
echo ==========================================
echo Excel Sync Add-in Unregistration
echo ==========================================
echo.

:: Check for admin rights
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script must be run as Administrator.
    echo Right-click and select "Run as administrator"
    pause
    exit /b 1
)

:: Set paths
set ADDIN_PATH=%~dp0..\src\ExcelSyncAddin\bin\Release\net48\ExcelSyncAddin.dll
set REGASM=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

:: Check if DLL exists
if not exist "%ADDIN_PATH%" (
    echo WARNING: ExcelSyncAddin.dll not found at expected path.
    echo Will attempt to unregister anyway...
)

echo Unregistering COM components...
"%REGASM%" "%ADDIN_PATH%" /unregister
if %errorLevel% neq 0 (
    echo WARNING: Unregistration may have had issues.
)

echo.
echo ==========================================
echo Unregistration complete!
echo ==========================================
echo.
echo Please restart Excel if it was running.
echo.
pause
