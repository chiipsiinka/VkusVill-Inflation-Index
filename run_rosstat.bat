@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
chcp 65001 > nul
title Rosstat weekly loader
python -m pip install -q --disable-pip-version-check -r requirements.txt
if "%~1"=="" (
    python -u rosstat.py --force --log log.txt
) else (
    python -u rosstat.py "%~1" --log log.txt
)
echo.
pause
