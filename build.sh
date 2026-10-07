#!/usr/bin/env bash
# Сборка релиза pyAutoImgTranslate под текущую ОС.
#   Windows (Git Bash / WSL): сборка Windows-версии
#   Linux (Debian/Ubuntu):    сборка Linux-версии
#
# Результат: dist/pyAutoImgTranslate(.exe)
set -e

echo "== Установка зависимостей сборки =="
python3 -m pip install -q --upgrade pip
python3 -m pip install -q -r requirements.txt

echo "== Сборка через PyInstaller =="
python3 -m PyInstaller --noconfirm --clean main.spec

echo "== Готово: dist/ =="
ls -la dist/
