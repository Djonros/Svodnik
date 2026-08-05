@echo off
chcp 65001 >nul
echo ==========================================
echo   Установка зависимостей для Сводник v1.4.0
echo ==========================================
echo.

:: Проверяем Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден!
    echo Установите Python 3.10+ с https://python.org
    echo Не забудьте поставить галочку "Add Python to PATH"
    pause
    exit /b 1
)

echo [1/4] Установка openpyxl (работа с Excel)...
pip install openpyxl --quiet
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить openpyxl
    pause
    exit /b 1
)

echo [2/4] Установка pywin32 (COM-подключение к КОМПАС)...
pip install pywin32 --quiet
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить pywin32
    pause
    exit /b 1
)

echo [3/4] Установка python-docx (работа с Word)...
pip install python-docx --quiet
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить python-docx
    pause
    exit /b 1
)

echo [4/4] Установка fpdf2 (работа с PDF)...
pip install fpdf2 --quiet
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить fpdf2
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   Все зависимости установлены!
echo ==========================================
echo.
echo Запуск: python app\gui_app.py
echo.
pause
