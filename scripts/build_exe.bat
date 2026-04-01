@echo off
chcp 65001 > nul
echo ============================================
echo  Сборка EXE - AnalyticsUI ОГПЗ/ОНГКМ
echo ============================================
echo.

cd /d "%~dp0.."

echo [1/3] Проверка PyInstaller...
.venv\Scripts\python -m pip show pyinstaller > nul 2>&1
if errorlevel 1 (
    echo Установка PyInstaller...
    .venv\Scripts\pip install pyinstaller
)
echo PyInstaller готов.
echo.

echo [2/3] Запуск сборки...
.venv\Scripts\pyinstaller --clean build_exe.spec
echo.

echo [3/3] Результат:
if exist "dist\AnalyticsUI_OGPZ.exe" (
    echo  УСПЕШНО: dist\AnalyticsUI_OGPZ.exe
    for %%A in ("dist\AnalyticsUI_OGPZ.exe") do echo  Размер: %%~zA байт
) else (
    echo  ОШИБКА: файл не создан, проверьте вывод выше
)

echo.
pause
