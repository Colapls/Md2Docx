@echo off
chcp 65001 >nul
setlocal

echo ===================================
echo   Markdown to DOCX Builder
echo ===================================
echo.

REM Always tee everything to a log file so we can debug if the
REM console closes unexpectedly.
set "LOG=%~dp0build.log"
echo [log] Output is being saved to: %LOG%
echo.

REM ---- Pick a builder. Prefer `python -m PyInstaller` because
REM it does not depend on pyinstaller.exe being on PATH. ----
set "BUILDER="
set "BUILD_CMD="

python -m PyInstaller --version >nul 2>nul
if %errorlevel%==0 (
    set "BUILDER=PyInstaller (python -m)"
    set "BUILD_CMD=python -m PyInstaller --noconfirm --clean build.spec"
    goto :do_build
)

where pyinstaller >nul 2>nul
if %errorlevel%==0 (
    set "BUILDER=PyInstaller"
    set "BUILD_CMD=pyinstaller --noconfirm --clean build.spec"
    goto :do_build
)

python -m nuitka --version >nul 2>nul
if %errorlevel%==0 (
    set "BUILDER=Nuitka (python -m)"
    set "BUILD_CMD=python -m nuitka --onefile --windows-disable-console --output-filename=Markdown2Docx.exe --output-dir=dist main.py"
    goto :do_build
)

where nuitka >nul 2>nul
if %errorlevel%==0 (
    set "BUILDER=Nuitka"
    set "BUILD_CMD=nuitka --onefile --windows-disable-console --output-filename=Markdown2Docx.exe --output-dir=dist main.py"
    goto :do_build
)

echo [ERROR] Neither PyInstaller nor Nuitka is installed. 1>&2
echo         Run:  pip install pyinstaller 1>&2
echo. 1>&2
echo Press any key to close... 1>&2
pause >nul
exit /b 1

:do_build
echo Using builder: %BUILDER%
echo Build command: %BUILD_CMD%
echo.

REM ---- Clean previous artifacts ----
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

REM ---- Run the build, capture log ----
echo Building...
%BUILD_CMD% > "%LOG%" 2>&1
set "RC=%errorlevel%"

REM Show the log tail in the console so the user sees what happened.
echo.
echo ---- Last 40 lines of build log ----
if exist "%LOG%" (
    powershell -NoProfile -Command "Get-Content -Path '%LOG%' -Tail 40"
) else (
    echo [WARN] No log produced.
)
echo ---- end ----
echo.

if not "%RC%"=="0" (
    echo [ERROR] Build failed with code %RC%. 1>&2
    echo See log: %LOG% 1>&2
    echo. 1>&2
    echo Press any key to close... 1>&2
    pause >nul
    exit /b %RC%
)

if not exist dist\Markdown2Docx.exe (
    echo [ERROR] dist\Markdown2Docx.exe was not produced. 1>&2
    echo See log: %LOG% 1>&2
    echo. 1>&2
    echo Press any key to close... 1>&2
    pause >nul
    exit /b 1
)

echo ===================================
echo   Build complete!
echo ===================================
echo Executable: %CD%\dist\Markdown2Docx.exe
echo.
pause
endlocal
