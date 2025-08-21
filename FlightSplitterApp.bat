@echo off
REM Chemin relatif vers le pythonw de l'env
set PYTHON_PATH=%~dp0\mini-flightsplitter\pythonw.exe
"%PYTHON_PATH%" "%~dp0\FlightSplitter.py"
