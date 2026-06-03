@echo off
:loop
python C:\fedan_rate\fedan_scraper.py
timeout /t 10
goto loop