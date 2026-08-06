@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "NO_PAUSE=0"
set "PY_ARGS="
:parse_args
if "%~1"=="" goto args_done
if /I "%~1"=="--nopause" (
  set "NO_PAUSE=1"
  shift
  goto parse_args
)
if /I "%~1"=="--no-pause" (
  set "NO_PAUSE=1"
  shift
  goto parse_args
)
set "PY_ARGS=%PY_ARGS% %1"
shift
goto parse_args
:args_done

if not exist "logs" mkdir "logs"

echo ============================================================
echo  CC6 Outlook Report
echo  Start: %DATE% %TIME%
echo  Log dir: %CD%\logs
echo ============================================================
echo.

:: Try multiple known Python locations
set "PYTHON_EXE="

:: 1) Try python via PATH (verify it actually runs)
for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON_EXE=%%i"
if defined PYTHON_EXE goto found_python

:: 2) Try python3
for /f "delims=" %%i in ('python3 -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON_EXE=%%i"
if defined PYTHON_EXE goto found_python

:: 3) Try py launcher
for /f "delims=" %%i in ('py -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON_EXE=%%i"
if defined PYTHON_EXE goto found_python

:: 4) Fallback: .workbuddy bundled Python
set "FALLBACK=%USERPROFILE%\.workbuddy\binaries\python\versions\3.14.3\python.exe"
if exist "%FALLBACK%" (
  set "PYTHON_EXE=%FALLBACK%"
  goto found_python
)

:: 5) Common install locations
for %%p in (
  "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
  "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
  "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
  "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
  "C:\Python313\python.exe"
  "C:\Python312\python.exe"
  "C:\Python311\python.exe"
) do (
  if exist %%p (
    set "PYTHON_EXE=%%p"
    goto found_python
  )
)

echo [FAILED] Python not found.
echo [FAILED] Python not found.> "logs\last_status.txt"
echo status=FAILED>> "logs\last_status.txt"
echo finished_at=%DATE% %TIME%>> "logs\last_status.txt"
goto finish_fail

:found_python
echo Using Python: %PYTHON_EXE%
echo.

"%PYTHON_EXE%" main.py %PY_ARGS%
set EXITCODE=%ERRORLEVEL%

echo.
echo ============================================================
if %EXITCODE% neq 0 (
  echo  RESULT: FAILED  exit_code=%EXITCODE%
  echo  See: logs\last_status.txt
  echo  See: logs\run.log  and latest logs\run_*.log
  echo ============================================================
  goto finish_fail
)

echo  RESULT: SUCCESS
echo  See: logs\last_status.txt
echo  See: logs\run.log  and latest logs\run_*.log
echo ============================================================
if "%NO_PAUSE%"=="0" (
  echo.
  pause
)
exit /b 0

:finish_fail
if "%NO_PAUSE%"=="0" (
  echo.
  echo Press any key to close...
  pause >nul
)
exit /b %EXITCODE%
