@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
chcp 65001 > nul
title VkusVill price scraper

echo Checking Python packages...
python -m pip install -q --disable-pip-version-check -r requirements.txt

echo Collecting prices...
python -u scraper.py --log log.txt
set RC=%errorlevel%

echo.
echo Checking Rosstat weekly release...
python -u rosstat.py --log log.txt
set RC2=%errorlevel%

echo.
if not "%RC%"=="0" (
    echo PRICES: finished with errors. See the messages above or log.txt
) else (
    echo PRICES: done.
)
if not "%RC2%"=="0" (
    echo ROSSTAT: could not load the release. See log.txt, download nedel_Ipc.xlsx by hand and run run_rosstat.bat
)
echo This window closes in 60 seconds (press any key to close now).
timeout /t 60 > nul
