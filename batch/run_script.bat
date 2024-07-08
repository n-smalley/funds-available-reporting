@echo off
setlocal

set "python_exe=c:/Users/nathansmalley/Desktop/Local-Files/Coding/funds-available-reporting/.venv/Scripts/python.exe"
set "script_path=c:/Users/nathansmalley/Desktop/Local-Files/Coding/funds-available-reporting/getReports.py"
set "log_file=c:/Users/nathansmalley/Desktop/Local-Files/Coding/funds-available-reporting/log.txt"

:: Run the Python script silently and redirect output to a log file
start /b "" "%python_exe%" "%script_path%" > "%log_file%" 2>&1

:: Check for errors
if errorlevel 1 (
    echo An error occurred. See log file for details.
    type "%log_file%"
    pause
)

endlocal
