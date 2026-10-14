@echo off
chcp 65001 >nul
setlocal EnableExtensions
rem Публикация релиза Сводника на GitHub (Djonros/Svodnik).
rem   release.bat         собрать exe (по вопросу), закоммитить, поставить тег, опубликовать релиз
rem   release.bat check   только показать, что будет сделано
rem Версия берется из app\gui_app.py (VERSION = "X.Y.Z"),
rem описание релиза из release_notes.md, файл релиза dist\Svodnik.exe.

cd /d "%~dp0"
set "REPO=Djonros/Svodnik"
set "NOTES=release_notes.md"
set "EXE=dist\Svodnik.exe"

for /f "tokens=2 delims==" %%v in ('findstr /r /c:"^ *VERSION = " app\gui_app.py') do set "VER=%%~v"
if defined VER set "VER=%VER: =%"
if defined VER set "VER=%VER:"=%"
if not defined VER (
  echo [ОШИБКА] Не нашел VERSION в app\gui_app.py
  goto fail
)
set "TAG=v%VER%"

where gh >nul 2>&1 || (echo [ОШИБКА] Не установлен GitHub CLI: winget install GitHub.cli & goto fail)
gh auth status >nul 2>&1 || (echo [ОШИБКА] gh не авторизован: выполните gh auth login & goto fail)
if not exist "%NOTES%" (echo [ОШИБКА] Нет файла %NOTES% с описанием релиза & goto fail)

for /f "usebackq delims=" %%s in (`gh repo view %REPO% --json visibility -q .visibility`) do set "VIS=%%s"
set "EXISTS="
gh release view %TAG% -R %REPO% >nul 2>&1 && set "EXISTS=1"

echo.
echo   Версия:      %VER%  (тег %TAG%)
echo   Репозиторий: %REPO%  (%VIS%)
if defined EXISTS (echo   Релиз %TAG% уже есть: будет заменен файл Svodnik.exe и описание) else (echo   Релиз %TAG% будет создан)
echo.
if /i "%VIS%"=="PRIVATE" echo   [!] Репозиторий закрыт: программы клиентов не увидят обновление. Откройте его: Settings ^> Danger Zone ^> Make public
if /i "%~1"=="check" goto done

choice /c YN /n /m "Собрать exe заново перед публикацией? [Y/N] "
if errorlevel 2 goto skipbuild
python -m PyInstaller --noconfirm Svodnik.spec || (echo [ОШИБКА] Сборка не удалась & goto fail)
:skipbuild
if not exist "%EXE%" (echo [ОШИБКА] Нет файла %EXE% & goto fail)

rem Коммит собранного exe (dist в .gitignore, поэтому -f)
git add -f "%EXE%"
git diff --cached --quiet || git commit -m "Сводник %VER%: exe" || goto fail

git rev-parse -q --verify "refs/tags/%TAG%" >nul || git tag "%TAG%" || goto fail
git push origin HEAD || goto fail
git push origin "%TAG%" || goto fail

if defined EXISTS (
  gh release upload %TAG% "%EXE%" -R %REPO% --clobber || goto fail
  gh release edit %TAG% -R %REPO% --title "Сводник %VER%" --notes-file "%NOTES%" --latest || goto fail
) else (
  gh release create %TAG% "%EXE%" -R %REPO% --title "Сводник %VER%" --notes-file "%NOTES%" --latest || goto fail
)

echo.
echo [OK] Релиз опубликован: https://github.com/%REPO%/releases/tag/%TAG%
:done
echo.
pause
exit /b 0
:fail
echo.
echo Релиз не опубликован.
pause
exit /b 1
