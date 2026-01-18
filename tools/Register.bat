@echo off
echo ==========================================
echo Excel Sync Add-in Registration
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
    echo ERROR: ExcelSyncAddin.dll not found.
    echo Please build the project first in Release mode.
    echo Expected path: %ADDIN_PATH%
    pause
    exit /b 1
)

:: Check if RegAsm exists
if not exist "%REGASM%" (
    echo ERROR: RegAsm.exe not found at %REGASM%
    echo Please ensure .NET Framework 4.8 is installed.
    pause
    exit /b 1
)

echo Registering COM components...
"%REGASM%" "%ADDIN_PATH%" /codebase /tlb
if %errorLevel% neq 0 (
    echo ERROR: Registration failed.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo Registration successful!
echo ==========================================
echo.
echo To use in Excel:
echo 1. Open Excel
echo 2. For RTD: =RTD("ExcelSync.RtdServer", "", "your_key")
echo 3. Restart Excel if it was open during registration
echo.
pause
