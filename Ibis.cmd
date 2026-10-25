@echo off
setlocal EnableExtensions

rem Ibis launcher.
rem
rem Double-click this file to start Ibis. It saves the analyst from opening a
rem terminal, changing the execution policy, and typing the script path.
rem
rem Ibis needs Administrator rights for Microsoft Defender exclusions and for
rem Windows long path support, so this launcher asks for elevation first. The
rem analyst stays in control: if the UAC prompt is declined, Ibis still starts
rem without those two abilities. Pass /noadmin to skip the UAC prompt.

set "IBIS_ROOT=%~dp0"
set "IBIS_SCRIPT=%IBIS_ROOT%Run-Ibis.ps1"
set "IBIS_LAUNCHER=%~f0"

if not exist "%IBIS_SCRIPT%" (
    echo [Ibis] Run-Ibis.ps1 was not found next to this launcher.
    echo [Ibis] Expected: %IBIS_SCRIPT%
    echo.
    pause
    exit /b 1
)

set "IBIS_ELEVATED=0"
set "IBIS_SKIP_ADMIN=0"
:parse_args
if "%~1"=="" goto args_done
if /i "%~1"=="/elevated" set "IBIS_ELEVATED=1"
if /i "%~1"=="/noadmin" set "IBIS_SKIP_ADMIN=1"
shift
goto parse_args
:args_done

rem fltmc needs Administrator rights and changes nothing, so it is a safe test.
fltmc >nul 2>&1
if not errorlevel 1 goto start_ibis

if "%IBIS_SKIP_ADMIN%"=="1" (
    echo [Ibis] Starting without Administrator rights, as requested.
    goto start_ibis
)

if "%IBIS_ELEVATED%"=="1" (
    echo [Ibis] The Administrator prompt was declined or did not apply.
    echo [Ibis] Ibis will start, but it cannot add Defender exclusions or change
    echo [Ibis] Windows long path support in this session.
    echo.
    goto start_ibis
)

echo [Ibis] Asking for Administrator rights...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "try { Start-Process -FilePath $env:IBIS_LAUNCHER -ArgumentList '/elevated' -Verb RunAs -ErrorAction Stop; exit 0 } catch { exit 1 }"
if not errorlevel 1 exit /b 0

echo [Ibis] Administrator rights were not granted.
echo [Ibis] Ibis will start, but it cannot add Defender exclusions or change
echo [Ibis] Windows long path support in this session.
echo.

:start_ibis
rem -ExecutionPolicy Bypass applies to this process only. It does not change the
rem machine or user execution policy.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%IBIS_SCRIPT%"
set "IBIS_EXIT=%ERRORLEVEL%"

if not "%IBIS_EXIT%"=="0" (
    echo.
    echo [Ibis] Ibis exited with code %IBIS_EXIT%.
    echo [Ibis] Session logs are under: %IBIS_ROOT%logs
    echo.
    pause
)

endlocal & exit /b %IBIS_EXIT%
