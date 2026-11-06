@echo off
chcp 65001 >nul
set PYTHONDONTWRITEBYTECODE=1
cd /d "%~dp0"
set PY=python
%PY% --version >nul 2>&1
if errorlevel 1 set PY=py -3
%PY% tools\sua_lich_chay.py
pause
