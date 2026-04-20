@echo off
chcp 65001 >nul
start "" pythonw "%~dp0report_check_gui.py"
exit