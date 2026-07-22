@echo off
rem Reyart gunluk 50 mail — Gorev Zamanlayici bu dosyayi her gun 09:30'da calistirir
cd /d "C:\Users\Home\OneDrive\Desktop\obsidian\01-Projects\scraper"
set PYTHONIOENCODING=utf-8
"C:\Users\Home\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\python.exe" send_daily_50.py >> output\_daily_mail.log 2>&1
