@echo off
cd /d "%~dp0"
echo Ejecutando Reseguimientos Amgen...
python scripts\consolidar_data.py
echo.
pause
