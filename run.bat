@echo off
cd /d "%~dp0"
python meeting_assistant.py
if errorlevel 1 pause
