@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
setlocal EnableDelayedExpansion

set "PORT=8502"
set "PY_EXE="
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"

if exist ".venv\Scripts\python.exe" (
    set "PY_EXE=.venv\Scripts\python.exe"
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set "PY_EXE=py -3"
    ) else (
        set "PY_EXE=python"
    )
)

echo [1/3] Проверка Python...
%PY_EXE% --version >nul 2>nul
if errorlevel 1 (
    echo Не удалось найти Python. Установите Python 3.10+ и повторите запуск.
    pause
    exit /b 1
)

echo [2/3] Проверка зависимостей...
%PY_EXE% -m pip show streamlit >nul 2>nul
if errorlevel 1 (
    echo Устанавливаю зависимости из requirements.txt...
    %PY_EXE% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Ошибка установки зависимостей.
        pause
        exit /b 1
    )
)

echo [3/3] Запуск приложения...
start "" "http://localhost:%PORT%"
%PY_EXE% -m streamlit run streamlit_app.py --server.port %PORT%

if errorlevel 1 (
    echo Не удалось запустить на порту !PORT!. Пробую порт 8503...
    set "PORT=8503"
    start "" "http://localhost:!PORT!"
    %PY_EXE% -m streamlit run streamlit_app.py --server.port !PORT!
)

if errorlevel 1 (
    echo Приложение завершилось с ошибкой.
    pause
    exit /b 1
)

endlocal
