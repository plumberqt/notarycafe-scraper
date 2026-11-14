@echo off
title NotaryCafe Directory Scraper
echo ========================================================
echo        NotaryCafe Automated Directory Scraper
echo ========================================================
echo.
echo Make sure your VPN (USA location) is connected!
echo Starting scraper...
echo.
python "%~dp0scrape_notarycafe_complete.py"
echo.
pause
