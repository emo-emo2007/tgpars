@echo off
setlocal
cd /d "%~dp0"

set "WAIT_AT_END=0"
if "%~1"=="" set "WAIT_AT_END=1"

if not exist ".venv\Scripts\python.exe" (
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3 -m venv .venv
    ) else (
        python -m venv .venv
    )
    if errorlevel 1 goto setup_failed
)

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto setup_failed

if "%~1"=="" (
    if not exist "links.txt" (
        echo Не найден links.txt. Создайте файл со ссылками или запустите:
        echo run.bat links.txt -o posts.docx
        set "EXIT_CODE=1"
        goto finish
    )
    ".venv\Scripts\python.exe" telegram_to_word.py links.txt -o posts.docx
) else (
    ".venv\Scripts\python.exe" telegram_to_word.py %*
)
set "EXIT_CODE=%ERRORLEVEL%"
goto finish

:setup_failed
echo Не удалось подготовить Python-окружение или установить зависимости.
set "EXIT_CODE=1"

:finish
if "%WAIT_AT_END%"=="1" pause
exit /b %EXIT_CODE%