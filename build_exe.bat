@echo off
REM ============================================
REM  Сборка запускаемого EXE-файла для Windows
REM  Требуется: Python 3.10+ и выполненный
REM  pip install -r requirements.txt
REM ============================================
python -m PyInstaller --noconfirm --onefile --windowed ^
  --name "Обработчик_Комплекс" ^
  app.py
echo.
echo Готово! Файл находится в папке dist\Обработчик_Комплекс.exe
pause
