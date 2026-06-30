@echo off
echo ========================================
echo  ������� - Сборка .exe
echo ========================================
echo.

echo Установка зависимостей...
pip install openpyxl pywin32 pyinstaller

echo.
echo Сборка .exe файла...
pyinstaller --onefile --windowed --name "�������" gui_app.py

echo.
echo Готово! Файл: dist\�������.exe
pause
