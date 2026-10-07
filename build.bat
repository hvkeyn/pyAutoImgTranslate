@echo off
REM Сборка релиза pyAutoImgTranslate под Windows (из корня проекта).
REM Результат: dist\pyAutoImgTranslate.exe

setlocal
if not exist .venv\Scripts\python.exe (
    echo Не найден .venv. Создайте окружение: python -m venv .venv
    exit /b 1
)

echo == Установка зависимостей сборки ==
.venv\Scripts\python.exe -m pip install -q --upgrade pip
.venv\Scripts\python.exe -m pip install -q -r requirements.txt

echo == Сборка через PyInstaller ==
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean main.spec

echo == Готово: dist\ ==
dir dist
endlocal
