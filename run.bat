@echo off
REM Запуск программы без сборки (нужен установленный Python)
python app.py
if errorlevel 1 pause
