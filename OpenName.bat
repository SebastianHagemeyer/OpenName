@echo off
cd /d "%~dp0"
python OpenName.py
if errorlevel 1 pause
