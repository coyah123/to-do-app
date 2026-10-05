@echo off
REM Launches the To-Do app with no console window.
REM pythonw runs without a terminal; falls back to python if needed.
start "" pythonw "%~dp0todo_app.py" 2>nul || start "" python "%~dp0todo_app.py"
