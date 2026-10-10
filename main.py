"""
pyAutoImgTranslate — перевод текста с экрана.

Возможности:
  * Захват области экрана — на Windows Win+Shift+S (штатный оверлей), на Linux
    Ctrl+Alt+S (вызывает установленный инструмент: grim/gnome-screenshot/spectacle/
    scrot/maim/flameshot). Далее OCR (Tesseract) или vision-нейросеть и перевод.
  * Поддержка нескольких провайдеров перевода: локальный LM Studio (OpenAI-совместимый
    сервер) и DeepSeek (облачный API). Можно добавлять свои модели.
  * Автозапуск при входе в систему (Windows — реестр, Linux — ~/.config/autostart),
    работа в системном трее, логирование в файл.
  * Ненавязчивое окно результата с картинкой и аккуратным блоком перевода.
"""

from __future__ import annotations

import ctypes
import base64
import hashlib
import io
import itertools
import json
import logging
import logging.handlers
import os
import queue
import re
import shutil
import struct
import subprocess
import sys
import threading
import time
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

import cv2
import numpy as np
import pytesseract
import requests

from PIL import Image, ImageDraw, ImageEnhance, ImageGrab, ImageTk

APP_VERSION = "1.7.0"

IS_WINDOWS = sys.platform.startswith("win")
IS_LINUX = sys.platform.startswith("linux")

# Windows-only библиотеки подключаем только на Windows.
if IS_WINDOWS:
    import win32clipboard  # type: ignore
    import win32con  # type: ignore

# Глобальные горячие клавиши: на Windows — keyboard; на Linux — pynput (если доступен).
if IS_WINDOWS:
    import keyboard  # type: ignore
else:
    keyboard = None  # type: ignore

# ---------------------------------------------------------------------------
# Пути и окружение (работает и из исходников, и из собранного .exe)
# ---------------------------------------------------------------------------
def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = _app_dir()
LOG_PATH = APP_DIR / "pyAutoImgTranslate.log"
CONFIG_PATH = APP_DIR / "config.json"


def resource_path(name: str) -> Path:
    """Путь к ресурсу: сначала рядом с exe/скриптом, потом в распакованном _MEIPASS."""
    candidates = [APP_DIR / name]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / name)
    for path in candidates:
        if path.is_file():
            return path
    return candidates[0]


# ---------------------------------------------------------------------------
# Логирование
# ---------------------------------------------------------------------------
logger = logging.getLogger("pyAutoImgTranslate")


def setup_logging() -> None:
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    try:
        handler = logging.handlers.RotatingFileHandler(
            LOG_PATH, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
    except Exception:
        handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(fmt)
    logger.addHandler(handler)

    # Ненавязчивый вывод в консоль, если она есть (не мешает при запуске без окна).
    if sys.stdout is not None and sys.stdout.isatty():
        stream = logging.StreamHandler(sys.stdout)
        stream.setFormatter(fmt)
        logger.addHandler(stream)


# ---------------------------------------------------------------------------
# Конфигурация / настройки
# ---------------------------------------------------------------------------
# DeepSeek API-ключ по умолчанию. Его можно переопределить в config.json
# или через переменную окружения DEEPSEEK_API_KEY.
# ВАЖНО: значение только ASCII — не-ASCII ключ ломает HTTP-заголовки (latin-1).
DEFAULT_DEEPSEEK_API_KEY = "sk-REPLACE-WITH-YOUR-KEY"

# Схема конфига полностью управляется данными: провайдеры лежат списком в
# CONFIG["providers"], у каждого — свой base_url, ключ и список моделей.
def _default_hotkey() -> str:
    # На Windows удобно использовать штатный Win+Shift+S, на Linux — Ctrl+Alt+S.
    return "windows+shift+s" if IS_WINDOWS else "ctrl+alt+s"


def _default_live_hotkey() -> str:
    return "windows+shift+a" if IS_WINDOWS else "ctrl+alt+a"


def _default_builtin_hotkey() -> str:
    """Клавиша для своего оверлея выделения.

    Win+Shift+S занимает сама Windows: нажатие открывает системные «Ножницы»,
    и перехватить его нельзя, поэтому для своего выделения берём свободное
    сочетание рядом.
    """
    return "windows+shift+x" if IS_WINDOWS else "ctrl+alt+x"


# Сочетания, которые Windows обрабатывает сама (их нельзя занять).
SYSTEM_CAPTURE_HOTKEYS = {"windows+shift+s", "win+shift+s", "cmd+shift+s", "super+shift+s"}


def _default_tesseract_cmd() -> str:
    if IS_WINDOWS:
        return r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    return shutil.which("tesseract") or "/usr/bin/tesseract"


DEFAULT_CONFIG = {
    "active_provider": "deepseek",     # id активного провайдера
    "active_model": "deepseek-flash",    # выбранная модель
    "ocr_lang": "eng+rus",
    "ocr_mode": "ai",    # "ai" (нейросеть читает картинку — точнее) | "tesseract" (локальный OCR)
    "tesseract_cmd": _default_tesseract_cmd(),
    "source_lang": "авто",
    "target_lang": "русский",
    "request_timeout": 120,
    "min_ocr_confidence": 40.0,  # минимальная средняя уверенность OCR (режим Tesseract)
    "add_to_autostart": False,
    "keep_image_on_top": False,
    "live_on_top": True,        # окно режима «на лету» показывать поверх других окон
    "skip_no_text": True,       # не тратить запрос, если на снимке нет текста
    "stream_results": True,     # показывать перевод по мере генерации (потоком)
    "capture_timeout": 60,      # сколько секунд ждать выделения области
    # Как выделять область: "builtin" — своим оверлеем (как Win+Shift+S),
    # "system" — системным средством (Win+Shift+S на Windows, grim/scrot и т.п. на Linux).
    "capture_mode": "builtin" if IS_WINDOWS else "system",
    "capture_dim": 0.45,        # затемнение экрана при своём выделении (0 — без затемнения)
    "log_recognized_text": False,   # писать распознанный текст в журнал (по умолчанию нет)
    "hotkey": _default_hotkey(),
    "live_hotkey": _default_live_hotkey(),
    "providers": [
        {
            "id": "deepseek",
            "name": "DeepSeek (облако)",
            "base_url": "https://api.deepseek.com",
            "api_key": "",                 # пусто -> env DEEPSEEK_API_KEY -> значение по умолчанию
            "api_key_env": "DEEPSEEK_API_KEY",
            "models": ["deepseek-flash", "deepseek-v4-pro", "deepseek-chat", "deepseek-reasoner"],
        },
        {
            "id": "local",
            "name": "LM Studio (локально)",
            "base_url": "http://127.0.0.1:1234/v1",
            "api_key": "lm-studio",
            "api_key_env": "",
            "models": [
                "qwen2.5-7b-instruct",
                "qwen2.5-14b-instruct",
                "llama-3.1-8b-instruct",
                "gemma-2-9b-it",
                "mistral-nemo-instruct-2407",
            ],
        },
        {
            "id": "custom",
            "name": "Свой сервер",
            "base_url": "https://api.example.com/v1",
            "api_key": "",
            "api_key_env": "OPENAI_API_KEY",
            "models": ["gpt-4o-mini", "gpt-4o", "claude-3-5-sonnet-latest"],
        },
    ],
}

# Псевдоним для обратной совместимости (заполняется из providers).
PROVIDER_LABELS: dict[str, str] = {}


def refresh_provider_labels() -> None:
    PROVIDER_LABELS.clear()
    for item in get_providers():
        PROVIDER_LABELS[item["id"]] = item.get("name", item["id"])


def _deep_merge(base: dict, override: dict) -> dict:
    result = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_config() -> dict:
    config = {k: (list(v) if isinstance(v, list) else dict(v) if isinstance(v, dict) else v)
              for k, v in DEFAULT_CONFIG.items()}
    if CONFIG_PATH.is_file():
        try:
            with CONFIG_PATH.open("r", encoding="utf-8") as fh:
                config = _deep_merge(config, json.load(fh))
        except Exception as exc:  # noqa: BLE001
            logger.warning("Не удалось прочитать config.json: %s", exc)
    return migrate_config(config)


def migrate_config(config: dict) -> dict:
    """Приводит старую схему (provider+deepseek/local/custom) к новой (providers)."""
    # Старый одиночный провайдер -> элемент списка providers.
    if not isinstance(config.get("providers"), list) or not config["providers"]:
        providers = []
        for pid in ("deepseek", "local", "custom"):
            section = config.get(pid)
            if isinstance(section, dict):
                providers.append({
                    "id": pid,
                    "name": section.get("name", pid),
                    "base_url": section.get("base_url", ""),
                    "api_key": section.get("api_key", ""),
                    "api_key_env": section.get("api_key_env", ""),
                    "models": list(section.get("models") or ([section["model"]] if section.get("model") else [])),
                })
        if providers:
            config["providers"] = providers
        else:
            config["providers"] = [dict(p) for p in DEFAULT_CONFIG["providers"]]

    # Активный провайдер: старое поле provider -> active_provider.
    if not config.get("active_provider"):
        config["active_provider"] = config.get("provider") or config["providers"][0]["id"]
    # Активная модель: старое translator_model -> active_model.
    if not config.get("active_model"):
        config["active_model"] = config.get("translator_model") or ""

    # Нормализуем записи провайдеров.
    seen = set()
    normalized = []
    for item in config["providers"]:
        if not isinstance(item, dict):
            continue
        pid = str(item.get("id") or item.get("name") or "provider").strip()
        if not pid or pid in seen:
            continue
        seen.add(pid)
        item["id"] = pid
        item.setdefault("name", pid)
        item.setdefault("base_url", "")
        item.setdefault("api_key", "")
        item.setdefault("api_key_env", "")
        models = item.get("models")
        item["models"] = [str(m) for m in models if str(m).strip()] if isinstance(models, list) else []
        normalized.append(item)
    config["providers"] = normalized

    # Гарантируем, что active_provider существует, и модель ему соответствует.
    ids = [p["id"] for p in normalized]
    if config["active_provider"] not in ids:
        config["active_provider"] = ids[0] if ids else ""
    section = get_provider(config["active_provider"], config)
    if section and config["active_model"] not in section.get("models", []):
        config["active_model"] = (section.get("models") or [config["active_model"]])[0]

    # Убираем устаревшие ключи, чтобы конфиг не разрастался.
    for stale in ("provider", "translator_model", "deepseek", "local", "custom"):
        config.pop(stale, None)
    return config


def get_providers(config: dict | None = None) -> list[dict]:
    return list((config or CONFIG).get("providers") or [])


def get_provider(provider_id: str, config: dict | None = None) -> dict | None:
    for item in get_providers(config):
        if item.get("id") == provider_id:
            return item
    return None


# Модели без поддержки изображений (vision). Проверено на практике:
# deepseek-v4-pro изображения не принимает. При свободном выборе модели
# предупреждаем пользователя, но не запрещаем.
NON_VISION_MODEL_HINTS = ("deepseek-v4-pro",)


def model_supports_vision(provider_id: str, model: str) -> bool:
    """Грубая оценка поддержки изображений выбранной моделью."""
    model_l = (model or "").strip().lower()
    if not model_l:
        return False
    if any(hint in model_l for hint in NON_VISION_MODEL_HINTS):
        return False
    # Явные vision-признаки.
    if any(k in model_l for k in ("vision", "vl", "llava", "gpt-4o", "gpt-4.1", "gemini", "claude-3", "qwen2-vl", "qwen2.5-vl")):
        return True
    # DeepSeek: flash/chat (и legacy reasoner) принимают картинки.
    if "deepseek" in model_l:
        return True
    # Для неизвестных моделей — считаем, что может не поддерживать (предупреждаем мягко).
    return False


def save_config(config: dict) -> None:
    try:
        with CONFIG_PATH.open("w", encoding="utf-8") as fh:
            json.dump(config, fh, ensure_ascii=False, indent=2)
    except Exception as exc:  # noqa: BLE001
        logger.warning("Не удалось сохранить config.json: %s", exc)


CONFIG = load_config()
refresh_provider_labels()


def env_var(name: str) -> str:
    """Значение переменной окружения, при необходимости — из реестра Windows.

    Переменную может не видеть сам процесс: приложение запущено не из проводника,
    переменную задали уже после входа в систему или запуск идёт из задачи
    планировщика. В этом случае читаем значение прямо из окружения пользователя,
    чтобы ключ находился без перезагрузки и без записи ключа в файлы.
    """
    if not name:
        return ""
    if name in _ENV_CACHE:
        return _ENV_CACHE[name]
    value = os.environ.get(name, "").strip()
    if not value and IS_WINDOWS:
        value = _env_var_from_registry(name)
    _ENV_CACHE[name] = value
    return value


def _env_var_from_registry(name: str) -> str:
    """Ищет переменную окружения пользователя и системы в реестре Windows."""
    try:
        import winreg

        locations = (
            (winreg.HKEY_CURRENT_USER, "Environment"),
            (winreg.HKEY_LOCAL_MACHINE,
             r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
        )
        for root, path in locations:
            try:
                with winreg.OpenKey(root, path, 0, winreg.KEY_READ) as key:
                    raw, _kind = winreg.QueryValueEx(key, name)
            except FileNotFoundError:
                continue
            except Exception as exc:  # noqa: BLE001
                logger.debug("Не удалось прочитать %s из реестра: %s", name, exc)
                continue
            if isinstance(raw, str) and raw.strip():
                resolved = os.path.expandvars(raw).strip()
                if resolved:
                    logger.debug("Переменная %s взята из реестра Windows.", name)
                    return resolved
    except Exception as exc:  # noqa: BLE001
        logger.debug("Реестр недоступен: %s", exc)
    return ""


# Значения переменных окружения кэшируем: ключ спрашивается на каждый запрос,
# а обращение к реестру — не бесплатное.
_ENV_CACHE: dict[str, str] = {}


def _provider_env_name(provider_id: str) -> str:
    section = get_provider(provider_id) or {}
    return (section.get("api_key_env") or "").strip() or (
        "DEEPSEEK_API_KEY" if provider_id == "deepseek" else "")


def get_api_key(provider_id: str) -> str:
    section = get_provider(provider_id) or {}
    key = (section.get("api_key") or "").strip()
    if key:
        return key
    env_name = _provider_env_name(provider_id)
    if env_name:
        env = env_var(env_name)
        if env:
            return env
    if provider_id == "deepseek":
        env = env_var("DEEPSEEK_API_KEY")
        if env:
            return env
        return DEFAULT_DEEPSEEK_API_KEY
    return ""


def key_source_hint(provider_id: str) -> str:
    """Откуда взялся ключ провайдера — для журнала (сам ключ не показываем)."""
    section = get_provider(provider_id) or {}
    if (section.get("api_key") or "").strip():
        return "из настроек (config.json)"
    env_name = _provider_env_name(provider_id)
    if env_name and env_var(env_name):
        return f"из переменной окружения {env_name}"
    if provider_id == "deepseek" and env_var("DEEPSEEK_API_KEY"):
        return "из переменной окружения DEEPSEEK_API_KEY"
    if provider_id == "deepseek":
        return "НЕ НАЙДЕН (используется плейсхолдер — перевода не будет)"
    return "НЕ НАЙДЕН"


# ---------------------------------------------------------------------------
# Захват области: свой оверлей или системное средство
# ---------------------------------------------------------------------------
def capture_mode() -> str:
    """Режим выделения области: "builtin" (свой оверлей) или "system"."""
    return capture_mode_of(CONFIG)


def capture_mode_of(config: dict) -> str:
    """То же, но по конкретному словарю настроек (в окне настроек — черновик)."""
    mode = str(config.get("capture_mode") or "").strip().lower()
    if mode in ("builtin", "system"):
        return mode
    return "builtin" if IS_WINDOWS else "system"


def hotkey_hint(mode: str, hotkey: str) -> str:
    """Строка-подсказка о действующей клавише захвата."""
    hotkey = (hotkey or "").strip() or _default_hotkey()
    if mode == "builtin" and hotkey.lower() in SYSTEM_CAPTURE_HOTKEYS:
        return f"клавиша: {_default_builtin_hotkey()} — Win+Shift+S занята системой"
    return f"клавиша: {hotkey}"


def effective_hotkey() -> str:
    """Действующая клавиша захвата с учётом режима выделения.

    В своём режиме Win+Shift+S использовать нельзя: её обрабатывает Windows,
    и вместе с нашим оверлеем открылись бы системные «Ножницы».
    """
    hotkey = (CONFIG.get("hotkey") or _default_hotkey()).strip()
    if capture_mode() == "builtin" and hotkey.lower() in SYSTEM_CAPTURE_HOTKEYS:
        return _default_builtin_hotkey()
    return hotkey


def _enable_dpi_awareness() -> None:
    """Физические пиксели вместо виртуальных: снимок и координаты совпадают.

    Без этого на экранах с масштабированием 125–200 % картинка захвата и окно
    выделения живут в разных системах координат, и область «съезжает».
    """
    if not IS_WINDOWS:
        return
    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except Exception:  # noqa: BLE001
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)     # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:  # noqa: BLE001
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:  # noqa: BLE001
            logger.debug("Не удалось включить DPI-awareness.")


# ---------------------------------------------------------------------------
# Автозапуск при входе в систему
#   Windows — реестр HKCU\...\Run (без прав администратора)
#   Linux   — ~/.config/autostart/pyAutoImgTranslate.desktop (freedesktop)
# ---------------------------------------------------------------------------
AUTOSTART_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_REG_NAME = "pyAutoImgTranslate"
AUTOSTART_DESKTOP = Path.home() / ".config" / "autostart" / "pyAutoImgTranslate.desktop"


def _launch_argv() -> list[str]:
    """Аргументы командной строки для запуска приложения."""
    if getattr(sys, "frozen", False):
        return [str(Path(sys.executable).resolve())]
    return [str(Path(sys.executable).resolve()), str(Path(__file__).resolve())]


def _autostart_command() -> str:
    argv = _launch_argv()
    if IS_WINDOWS:
        python = Path(argv[0])
        pythonw = python.with_name("pythonw.exe")
        if len(argv) == 2 and pythonw.is_file():
            argv[0] = str(pythonw)
        return " ".join(f'"{a}"' for a in argv)
    # Linux: экранируем пробелы для Exec=.
    return " ".join(f'"{a}"' for a in argv)


def _registered_autostart_command() -> str:
    """Команда автозапуска, записанная в системе (реестр / .desktop). Пусто — если нет."""
    if IS_LINUX:
        try:
            for line in AUTOSTART_DESKTOP.read_text(encoding="utf-8").splitlines():
                if line.startswith("Exec="):
                    return line[5:].strip()
        except FileNotFoundError:
            return ""
        except Exception as exc:  # noqa: BLE001
            logger.warning("Не удалось прочитать %s: %s", AUTOSTART_DESKTOP, exc)
        return ""
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, AUTOSTART_REG_NAME)
            return str(value or "")
    except FileNotFoundError:
        return ""
    except Exception as exc:  # noqa: BLE001
        logger.warning("Не удалось проверить автозапуск: %s", exc)
        return ""


def is_autostart_enabled() -> bool:
    return bool(_registered_autostart_command().strip())


def autostart_is_stale() -> bool:
    """True, если запись автозапуска ведёт на несуществующий файл.

    Так чинится только реально сломанная запись (приложение переехало или
    удалено). Запуск из исходников при этом не переписывает автозапуск
    собранного .exe — команды просто разные.
    """
    current = _registered_autostart_command().strip()
    if not current:
        return False
    return not autostart_command_is_valid(current)


def autostart_command_is_valid(command: str) -> bool:
    """Все ли файлы из команды автозапуска существуют на диске."""
    targets = re.findall(r'"([^"]+)"', command) or command.split()
    files = [Path(p) for p in targets if p.lower().endswith((".exe", ".py", ".com", ".bat", ".cmd"))]
    if not files:
        return False
    return all(path.is_file() for path in files)


def set_autostart(enabled: bool) -> bool:
    try:
        if IS_LINUX:
            if enabled:
                AUTOSTART_DESKTOP.parent.mkdir(parents=True, exist_ok=True)
                AUTOSTART_DESKTOP.write_text(
                    "[Desktop Entry]\n"
                    "Type=Application\n"
                    "Name=pyAutoImgTranslate\n"
                    "Comment=Перевод текста с экрана\n"
                    f"Exec={_autostart_command()}\n"
                    "X-GNOME-Autostart-enabled=true\n"
                    "Terminal=false\n",
                    encoding="utf-8",
                )
            else:
                try:
                    AUTOSTART_DESKTOP.unlink()
                except FileNotFoundError:
                    pass
        else:
            import winreg

            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, AUTOSTART_REG_PATH, 0, winreg.KEY_SET_VALUE
            ) as key:
                if enabled:
                    winreg.SetValueEx(key, AUTOSTART_REG_NAME, 0, winreg.REG_SZ, _autostart_command())
                else:
                    try:
                        winreg.DeleteValue(key, AUTOSTART_REG_NAME)
                    except FileNotFoundError:
                        pass
        CONFIG["add_to_autostart"] = enabled
        save_config(CONFIG)
        logger.info("Автозапуск %s (команда: %s).",
                    "включён" if enabled else "выключен", _autostart_command())
        return True
    except Exception as exc:  # noqa: BLE001
        logger.error("Не удалось изменить автозапуск: %s", exc)
        return False


def ensure_autostart_consistency() -> None:
    """Синхронизирует автозапуск с настройкой в config.json при старте.

    Кроме включения/выключения чиним устаревшую запись: если приложение
    переехало (новый путь к .exe или к main.py), команда автозапуска
    перезаписывается актуальной.
    """
    want = bool(CONFIG.get("add_to_autostart"))
    if want and not is_autostart_enabled():
        set_autostart(True)
    elif want and autostart_is_stale():
        logger.info("Команда автозапуска устарела — перезаписываю.")
        set_autostart(True)
    elif not want and is_autostart_enabled():
        set_autostart(False)


# ---------------------------------------------------------------------------
# Выравнивание изображения (deskew) и OCR
# ---------------------------------------------------------------------------
def deskew_image(pil_image: Image.Image) -> Image.Image:
    """Определяет угол наклона текста и выравнивает изображение без обрезки."""
    cv_image = cv2.cvtColor(np.array(pil_image.convert("RGB")), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)
    gray = cv2.bitwise_not(gray)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

    coords = np.column_stack(np.where(thresh > 0))
    if len(coords) == 0:
        return pil_image

    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle

    if abs(angle) < 0.3:  # почти ровно — не трогаем
        return pil_image

    (h, w) = cv_image.shape[:2]
    center = (w // 2, h // 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    cos = abs(matrix[0, 0])
    sin = abs(matrix[0, 1])
    new_w = int((h * sin) + (w * cos))
    new_h = int((h * cos) + (w * sin))
    matrix[0, 2] += (new_w / 2) - center[0]
    matrix[1, 2] += (new_h / 2) - center[1]

    rotated = cv2.warpAffine(
        cv_image, matrix, (new_w, new_h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE
    )
    return Image.fromarray(cv2.cvtColor(rotated, cv2.COLOR_BGR2RGB))


def upscale_image(pil_image: Image.Image, scale: float = 2.0) -> Image.Image:
    cv_image = cv2.cvtColor(np.array(pil_image.convert("RGB")), cv2.COLOR_RGB2BGR)
    upscaled = cv2.resize(cv_image, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    return Image.fromarray(cv2.cvtColor(upscaled, cv2.COLOR_BGR2RGB))


def _configure_tesseract() -> None:
    """Настраивает путь к tesseract.exe и к папке tessdata (TESSDATA_PREFIX)."""
    cmd = CONFIG.get("tesseract_cmd") or r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    pytesseract.pytesseract.tesseract_cmd = cmd

    # Автоопределение tessdata: рядом с exe в подпапке tessdata, либо в самом exe-каталоге.
    exe_dir = os.path.dirname(cmd)
    candidates = [
        os.environ.get("TESSDATA_PREFIX", ""),
        os.path.join(exe_dir, "tessdata"),
        exe_dir,
    ]
    for candidate in candidates:
        if candidate and os.path.isfile(os.path.join(candidate, "eng.traineddata")):
            os.environ["TESSDATA_PREFIX"] = candidate
            return


MIN_OCR_CONFIDENCE = 40.0
# Сколько «глифоподобных» деталей на снимке считать признаком пёстрой картинки,
# которую стоит проверить локальным OCR, даже если строк не нашлось.
TEXTURE_GLYPHS_LIMIT = 60


class OcrResult:
    """Результат Tesseract: текст, средняя уверенность (0–100) и число слов."""

    __slots__ = ("text", "confidence", "words")

    def __init__(self, text: str, confidence: float, words: int) -> None:
        self.text = text
        self.confidence = confidence
        self.words = words

    @property
    def looks_like_text(self) -> bool:
        """Похоже ли это на настоящий текст, а не на шум с фотографии."""
        threshold = float(CONFIG.get("min_ocr_confidence") or MIN_OCR_CONFIDENCE)
        return (self.words >= 1 and len(self.text.strip()) >= 2
                and self.confidence >= threshold)

    def __repr__(self) -> str:
        return (f"OcrResult(words={self.words}, confidence={self.confidence:.0f}, "
                f"chars={len(self.text)})")


def ocr_image(image: Image.Image) -> str:
    """Выпрямляет, увеличивает изображение и извлекает текст через Tesseract."""
    return ocr_image_ex(image).text


def ocr_image_ex(image: Image.Image) -> OcrResult:
    """OCR с оценкой уверенности: слова собираются в строки по данным Tesseract."""
    _configure_tesseract()
    deskewed = deskew_image(image)
    upscaled = upscale_image(deskewed, scale=2)
    lang = CONFIG.get("ocr_lang", "eng+rus")
    config = r"--psm 6 --oem 3"
    tried: list[str] = []
    for attempt_lang in (lang, "eng"):
        if not attempt_lang or attempt_lang in tried:
            continue
        tried.append(attempt_lang)
        try:
            data = pytesseract.image_to_data(
                upscaled, lang=attempt_lang, config=config, output_type=pytesseract.Output.DICT
            )
        except pytesseract.TesseractError as exc:
            logger.warning("OCR с языком '%s' не удался (%s).", attempt_lang, exc)
            continue
        return _ocr_result_from_data(data)
    return OcrResult("", 0.0, 0)


def _ocr_result_from_data(data: dict) -> OcrResult:
    """Собирает строки из вывода image_to_data, сохраняя уверенность по словам."""
    total = len(data.get("text") or [])
    lines: dict[tuple[int, int, int], list[str]] = {}
    confidences: list[float] = []
    words = 0
    for index in range(total):
        word = (data["text"][index] or "").strip()
        if not word:
            continue
        try:
            confidence = float(data["conf"][index])
        except (TypeError, ValueError, KeyError):
            confidence = -1.0
        key = (int(data["block_num"][index]), int(data["par_num"][index]), int(data["line_num"][index]))
        lines.setdefault(key, []).append(word)
        confidences.append(max(0.0, confidence))
        words += 1
    text = _clean_ocr_text("\n".join(" ".join(parts) for _, parts in sorted(lines.items())))
    mean_confidence = sum(confidences) / len(confidences) if confidences else 0.0
    return OcrResult(text, mean_confidence, words)


def _clean_ocr_text(text: str) -> str:
    """Убирает типичный мусор OCR: дубли пробелов, одиночные символы, пустые строки."""
    lines = []
    for line in text.splitlines():
        cleaned = re.sub(r"[ \t]+", " ", line).strip()
        if len(cleaned) <= 1:            # одиночные «символы-артефакты» — в мусор
            continue
        lines.append(cleaned)
    return "\n".join(lines).strip()


# ---------------------------------------------------------------------------
# «Есть ли на картинке текст» — локальная проверка до обращения к нейросети.
#
# Идея: фотографии, иконки, обои и пустые области не должны уходить в API.
# Проверка идёт по структуре: текст — это строка из похожих по высоте и
# равномерно расположенных глифов. Если структура видна уверенно, снимок
# обрабатывается сразу; если непонятно — уточняем дешёвым локальным OCR.
# ---------------------------------------------------------------------------
def _component_solidity(labels, index: int, x: int, y: int, w: int, h: int) -> float:
    """Насколько фигура «залита»: 1.0 — выпуклое пятно, меньше — буква с просветами."""
    mask = (labels[y:y + h, x:x + w] == index).astype(np.uint8)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return 1.0
    contour = max(contours, key=cv2.contourArea)
    area = cv2.contourArea(contour)
    hull_area = cv2.contourArea(cv2.convexHull(contour))
    if hull_area <= 0 or area <= 0:
        return 1.0
    return float(area / hull_area)


def analyse_text_presence(image: Image.Image) -> tuple[float, bool, int]:
    """Грубая локальная проверка «есть ли на снимке текст».

    Возвращает (оценка, уверенно_текст, число_деталей). Оценка — сколько
    похожих по высоте глифов складывается в лучшую текстовую строку:

      * 0        — строк не видно (заливка, градиент, шум, часть фотографий);
      * 3 и выше — структура похожа на текст;
      * 1–2      — непонятно.

    Число деталей (мелких «глифоподобных» пятен) отличает пустой кадр от
    пёстрого: если деталей много, снимок стоит проверить локальным OCR.
    """
    try:
        gray = np.array(image.convert("L"))
    except Exception as exc:  # noqa: BLE001
        logger.debug("Не удалось проанализировать изображение: %s", exc)
        return 1.0, False, 0
    if gray.ndim != 2 or gray.size == 0:
        return 1.0, False, 0
    height, width = gray.shape[:2]
    if height < 8 or width < 8:
        return 0.0, False, 0

    # Крупные снимки уменьшаем: зернистость фотографий при усреднении уходит,
    # а скорость анализа перестаёт зависеть от размера экрана.
    largest = max(height, width)
    if largest > 1600:
        factor = 1600.0 / largest
        gray = cv2.resize(gray, None, fx=factor, fy=factor, interpolation=cv2.INTER_AREA)
    # Мелкие выделения наоборот увеличиваем: глифы должны быть высотой ~8–40 px.
    elif largest < 400:
        factor = min(3.0, 800.0 / largest)
        if factor > 1.1:
            gray = cv2.resize(gray, None, fx=factor, fy=factor, interpolation=cv2.INTER_CUBIC)

    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 12
    )
    # Открытие убирает одиночные точки (шум), но не склеивает соседние буквы —
    # иначе слова превращаются в одно пятно и перестают быть похожи на текст.
    binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))

    scaled_height = binary.shape[0]
    count, labels, stats, _centroids = cv2.connectedComponentsWithStats(binary, 8)
    if count <= 2:
        return 0.0, False, 0
    # Отбор кандидатов делаем векторно: на полноэкранных снимках компонентов
    # бывают сотни тысяч, и поштучный обход в Python заметно тормозит.
    boxes = stats[1:].astype(np.int64)
    box_w, box_h, area = boxes[:, 2], boxes[:, 3], boxes[:, 4]
    ratio = box_w / np.maximum(box_h, 1)
    fill = area / np.maximum(box_w * box_h, 1)
    keep = ((box_h >= 7) & (box_h <= scaled_height * 0.5) & (box_w >= 2) & (area >= 15)
            & (ratio <= 30) & (fill >= 0.12) & (fill <= 0.92)
            & ~((ratio > 2.5) & (fill > 0.6)))
    glyphs: list[tuple[int, int, int, int, int, float, int]] = [
        (int(boxes[i, 0]), int(boxes[i, 1]), int(box_w[i]), int(box_h[i]), int(area[i]),
         float(ratio[i]), i + 1)
        for i in np.nonzero(keep)[0]
    ]

    if len(glyphs) < 2:
        return float(len(glyphs)), False, len(glyphs)

    # Строки собираем по перекрытию по вертикали, а не по центру: так надёжнее
    # для текста с выносными элементами (у, р, б, Ц).
    lines: list[list[tuple[int, int, int, int, int, float, int]]] = []
    for glyph in sorted(glyphs, key=lambda b: b[1]):
        top, bottom = glyph[1], glyph[1] + glyph[3]
        placed = False
        for line in lines:
            line_top = min(b[1] for b in line)
            line_bottom = max(b[1] + b[3] for b in line)
            line_h = float(np.median([b[3] for b in line]))
            overlap = min(bottom, line_bottom) - max(top, line_top)
            if overlap > 0.55 * min(glyph[3], line_h) and 0.5 <= glyph[3] / line_h <= 2.0:
                line.append(glyph)
                placed = True
                break
        if not placed:
            lines.append([glyph])

    best = 0.0
    good_lines = 0
    for line in lines:
        if len(line) < 2:
            continue
        same = [b for b in line if 0.6 <= b[3] / float(np.median([g[3] for g in line])) <= 1.6]
        if len(same) < 2:
            continue
        line_h = float(np.median([b[3] for b in same]))
        ordered = sorted(same, key=lambda b: b[0])
        span = ordered[-1][0] + ordered[-1][2] - ordered[0][0]
        gaps = [max(0, ordered[i + 1][0] - (ordered[i][0] + ordered[i][2]))
                for i in range(len(ordered) - 1)]
        gap_median = float(np.median(gaps)) if gaps else 0.0
        ink = sum(b[4] for b in same)
        density = ink / float(max(1, span * line_h))
        # У настоящей строки буквы стоят в ряд: они не налезают друг на друга и
        # занимают по ширине заметно меньше длины строки. У зернистой фотографии
        # пятна собираются «стопками», и плотность строки сразу выдаёт шум.
        width_share = sum(b[2] for b in same) / float(max(1, span))
        if density > 0.9 or width_share > 1.5:
            continue
        score = 0.0
        if len(same) >= 3 and gap_median <= 1.5 * line_h and density >= 0.06 \
                and span >= 1.8 * line_h:
            score = float(len(same))
        elif (len(same) == 2 and gap_median <= 1.5 * line_h and density >= 0.12
              and all(0.25 <= b[5] <= 1.8 for b in same)
              and 0.5 * line_h <= span <= 12 * line_h
              and min(_component_solidity(labels, b[6], b[0], b[1], b[2], b[3])
                      for b in same) <= 0.85):
            # Короткое слово из двух знаков: у настоящих букв почти всегда есть
            # просветы/изломы, у залитых фигур (кружок, кнопка) — нет.
            score = 2.0
        if score >= 2.0:
            good_lines += 1
        best = max(best, score)

    has_text = best >= 3.0 or good_lines >= 2
    return best, has_text, len(glyphs)


def confirm_text_with_ocr(image: Image.Image) -> bool | None:
    """Уточняет «есть ли текст» локальным OCR (Tesseract).

    Именно OCR надёжно отличает текст от фактуры фотографии, поэтому спорные
    снимки отдаём ему. Проверка идёт на уменьшенной копии и в быстром режиме
    (LSTM), чтобы не задерживать обработку.

    Возвращает True/False или None, если проверить не удалось (Tesseract не
    установлен, ошибка распознавания) — тогда решение принимает вызывающий.
    """
    try:
        _configure_tesseract()
        small = image if max(image.size) <= 1100 else _fit_for_preview(image, 1100, 1100)
        data = pytesseract.image_to_data(
            small, lang=CONFIG.get("ocr_lang", "eng+rus"), config="--psm 6 --oem 1",
            output_type=pytesseract.Output.DICT,
        )
    except Exception as exc:  # noqa: BLE001
        logger.info("Проверка через Tesseract недоступна: %s", exc)
        return None
    result = _ocr_result_from_data(data)
    logger.info("Локальная проверка текста: %r", result)
    text = result.text.strip()
    if len(text) < 2:
        return False
    threshold = float(CONFIG.get("min_ocr_confidence") or MIN_OCR_CONFIDENCE)
    # Одиночное короткое «слово» легко получается из фигур (кружок похож на «O»),
    # поэтому для него требуем заметно более высокую уверенность.
    if result.words >= 2:
        return result.confidence >= threshold
    return result.confidence >= max(75.0, threshold + 25.0)


def image_has_text(image: Image.Image) -> bool:
    """Итоговое решение «стоит ли отправлять этот снимок на перевод».

    Сначала дешёвая проверка структуры. Если строк текста не видно и деталей на
    снимке мало (заливка, градиент, простой пейзаж) — сразу «нет», без OCR.
    В остальных случаях последнее слово за локальным OCR: он надёжно отличает
    текст от фактуры фотографии. Если OCR недоступен — снимок всё равно
    обрабатывается: лучше лишний запрос, чем потерянный перевод.
    """
    score, confident, glyphs = analyse_text_presence(image)
    if score <= 0 and glyphs < TEXTURE_GLYPHS_LIMIT:
        logger.info("Текстовых структур на снимке нет (оценка 0, деталей %d).", glyphs)
        return False
    verdict = confirm_text_with_ocr(image)
    if verdict is None:
        logger.info("Оценка %.0f (уверенно=%s, деталей %d), подтвердить нечем — обрабатываю.",
                    score, confident, glyphs)
        return True
    logger.info("Оценка %.0f (уверенно=%s, деталей %d), локальная проверка: текст=%s.",
                score, confident, glyphs, verdict)
    return verdict


# ---------------------------------------------------------------------------
# Работа с буфером обмена (Windows: WinAPI/CF_DIB; Linux: xclip/wl-paste)
# ---------------------------------------------------------------------------
def _linux_clipboard_tool() -> list[str] | None:
    """Возвращает команду для чтения картинки из буфера (Wayland → X11)."""
    if os.environ.get("WAYLAND_DISPLAY") and shutil.which("wl-paste"):
        return ["wl-paste", "--no-newline", "--type", "image/png"]
    if shutil.which("xclip"):
        return ["xclip", "-selection", "clipboard", "-t", "image/png", "-o"]
    if shutil.which("xsel"):
        return ["xsel", "--clipboard", "--output"]
    return None


def _linux_clipboard_text() -> str:
    """Текст из буфера — нужен для «номера изменения» (аналога sequence number)."""
    for cmd in (["wl-paste", "--no-newline"],
                ["xclip", "-selection", "clipboard", "-o"],
                ["xsel", "--clipboard", "--output"]):
        if shutil.which(cmd[0]):
            try:
                out = subprocess.run(cmd, capture_output=True, timeout=3)
                return out.stdout.decode("utf-8", "ignore")
            except Exception:  # noqa: BLE001
                continue
    return ""


def get_image_from_clipboard() -> Image.Image | None:
    """Получает изображение из буфера обмена (кросс-платформенно)."""
    # PIL умеет читать буфер на Windows и macOS.
    try:
        image = ImageGrab.grabclipboard()
        if image and isinstance(image, Image.Image):
            return image
    except Exception as exc:  # noqa: BLE001
        logger.debug("ImageGrab.grabclipboard недоступен: %s", exc)

    if IS_LINUX:
        tool = _linux_clipboard_tool()
        if tool is None:
            return None
        try:
            out = subprocess.run(tool, capture_output=True, timeout=5)
            if out.returncode == 0 and out.stdout:
                return Image.open(io.BytesIO(out.stdout))
        except Exception as exc:  # noqa: BLE001
            logger.debug("Не удалось прочитать картинку через %s: %s", tool[0], exc)
        return None

    if IS_WINDOWS:
        data = None
        try:
            win32clipboard.OpenClipboard()
            try:
                data = win32clipboard.GetClipboardData(win32con.CF_DIB)
            except Exception as exc:  # noqa: BLE001
                logger.debug("CF_DIB недоступен: %s", exc)
            finally:
                win32clipboard.CloseClipboard()
        except Exception as exc:  # noqa: BLE001
            logger.debug("Буфер обмена недоступен: %s", exc)

        if data is None:
            return None
        try:
            header = b"BM" + struct.pack("<I", len(data) + 14) + b"\x00\x00\x00\x00\x36\x00\x00\x00"
            return Image.open(io.BytesIO(header + data))
        except Exception as exc:  # noqa: BLE001
            logger.debug("Не удалось собрать изображение из CF_DIB: %s", exc)
            return None
    return None


def wait_for_clipboard_image(timeout: float = 5.0) -> Image.Image | None:
    """Публичный хелпер: ждёт появления картинки в буфере обмена."""
    start = time.time()
    while time.time() - start < timeout:
        image = get_image_from_clipboard()
        if image and isinstance(image, Image.Image):
            return image
        time.sleep(0.2)
    return None


def _clipboard_sequence() -> int:
    """Номер изменения буфера обмена — растёт при каждом копировании.

    Windows — нативный счётчик; Linux — хеш текущего содержимого.
    """
    if IS_WINDOWS:
        try:
            return int(ctypes.windll.user32.GetClipboardSequenceNumber())
        except Exception:  # noqa: BLE001
            return 0
    if IS_LINUX:
        try:
            return hash(_linux_clipboard_text()) & 0x7FFFFFFF
        except Exception:  # noqa: BLE001
            return 0
    return 0


# ---------------------------------------------------------------------------
# Задачи обработки: отмена и «перес-скриншот».
#
# Каждый захват — отдельная задача (Job) со своим номером и флагом отмены.
# Новый захват отменяет предыдущий: если снимок сделан не тот или хочется
# переснять — достаточно нажать горячую клавишу ещё раз, обработка бросится,
# а новый снимок обработается сразу.
# ---------------------------------------------------------------------------
class Cancelled(Exception):
    """Обработка прервана: пользователь отменил её или сделал новый скриншот."""


class Job:
    """Одна операция «захват → распознавание → перевод»."""

    __slots__ = ("kind", "gen", "cancel", "created")

    def __init__(self, kind: str, gen: int) -> None:
        self.kind = kind                 # "hotkey" (по клавише) | "live" («на лету»)
        self.gen = gen                   # номер задачи, растёт с каждым захватом
        self.cancel = threading.Event()
        self.created = time.time()

    @property
    def interactive(self) -> bool:
        """Запущена ли задача вручную (тогда окну результата можно забрать фокус)."""
        return self.kind == "hotkey"

    def cancelled(self) -> bool:
        return self.cancel.is_set()

    def check(self) -> None:
        """Бросает Cancelled, если обработку прервали."""
        if self.cancel.is_set():
            raise Cancelled()


_job_seq = itertools.count(1)
_job_lock = threading.Lock()
_CURRENT_JOB: "Job | None" = None
# Активные HTTP-ответы по номеру задачи: закрываем их при отмене, чтобы
# не ждать окончания генерации и не платить за ненужный ответ.
_job_http: dict[int, object] = {}


def _new_job(kind: str) -> Job:
    """Начинает новую задачу, прерывая предыдущую (перес-скриншот = отмена)."""
    global _CURRENT_JOB
    with _job_lock:
        previous, _CURRENT_JOB = _CURRENT_JOB, Job(kind, next(_job_seq))
        job = _CURRENT_JOB
    if previous is not None and not previous.cancel.is_set():
        previous.cancel.set()
        _close_job_http(previous.gen)
        logger.info("Обработка #%d прервана новым захватом (#%d).", previous.gen, job.gen)
    return job


def _finish_job(job: Job) -> None:
    """Снимает задачу с учёта, если её не вытеснила более новая."""
    global _CURRENT_JOB
    with _job_lock:
        if _CURRENT_JOB is job:
            _CURRENT_JOB = None
    _job_http.pop(job.gen, None)


def _current_job() -> "Job | None":
    return _CURRENT_JOB


def _is_active_job(gen: int | None) -> bool:
    """True, если результат задачи с таким номером ещё актуален."""
    if gen is None:
        return True
    job = _CURRENT_JOB
    return job is None or job.gen == gen


def cancel_current_job(reason: str = "отменено пользователем") -> bool:
    """Прерывает текущую обработку. True — если было что прерывать."""
    job = _CURRENT_JOB
    if job is None or job.cancel.is_set():
        return False
    job.cancel.set()
    _close_job_http(job.gen)
    logger.info("Обработка #%d прервана (%s).", job.gen, reason)
    return True


def _register_job_http(gen: int, response) -> None:
    _job_http[gen] = response


def _close_job_http(gen: int) -> None:
    """Закрывает соединение задачи: заблокированное чтение прервётся сразу."""
    response = _job_http.pop(gen, None)
    if response is not None:
        try:
            response.close()
        except Exception:  # noqa: BLE001
            pass


# ---------------------------------------------------------------------------
# Режим «на лету»: как только в буфере появится новая картинка — сразу OCR+перевод.
# Удобно использовать с Win+Shift+S или любым другим средством скриншотов.
# ---------------------------------------------------------------------------
LIVE_STATE = {"enabled": False, "thread": None, "last_seq": 0, "last_digest": ""}
_live_lock = threading.Lock()


def _image_digest(image: Image.Image) -> str:
    """Отпечаток картинки: одинаковые снимки не обрабатываем дважды."""
    try:
        small = image.convert("L").resize((32, 32), Image.Resampling.BILINEAR)
        payload = small.tobytes() + f"{image.size}".encode("ascii")
        return hashlib.sha1(payload).hexdigest()
    except Exception:  # noqa: BLE001
        return ""


def _live_watcher() -> None:
    """Фоновый поток: следит за буфером обмена.

    Сама обработка идёт в отдельном потоке, поэтому наблюдатель остаётся
    отзывчивым: новый снимок в буфере мгновенно вытесняет предыдущую обработку.
    """
    logger.info("Режим «на лету» включён — жду картинку в буфере обмена.")
    # Текущее состояние считываем один раз, чтобы не реагировать на уже лежащую картинку.
    LIVE_STATE["last_seq"] = _clipboard_sequence()
    while LIVE_STATE["enabled"]:
        time.sleep(0.25)
        if not LIVE_STATE["enabled"]:
            break
        seq = _clipboard_sequence()
        if seq == LIVE_STATE["last_seq"]:
            continue
        LIVE_STATE["last_seq"] = seq
        image = get_image_from_clipboard()
        if not image or not isinstance(image, Image.Image):
            continue
        digest = _image_digest(image)
        if digest and digest == LIVE_STATE["last_digest"]:
            continue                          # та же самая картинка — не пересчитываем
        LIVE_STATE["last_digest"] = digest
        job = _new_job("live")
        threading.Thread(target=_run_job, args=(job, image), name=f"live-job-{job.gen}",
                         daemon=True).start()
    logger.info("Режим «на лету» выключен.")


def set_live_mode(enabled: bool) -> None:
    """Включает/выключает автоматическое распознавание из буфера обмена."""
    with _live_lock:
        if enabled and not LIVE_STATE["enabled"]:
            LIVE_STATE["enabled"] = True
            LIVE_STATE["thread"] = threading.Thread(target=_live_watcher, name="live-ocr", daemon=True)
            LIVE_STATE["thread"].start()
        elif not enabled and LIVE_STATE["enabled"]:
            LIVE_STATE["enabled"] = False
            # Незачем держать в буфере прошлый кадр — иначе при следующем
            # включении он обработается повторно.
            LIVE_STATE["last_digest"] = ""


def is_live_mode() -> bool:
    return bool(LIVE_STATE["enabled"])


PROGRESS: "Progress | None" = None


def _progress() -> "Progress":
    global PROGRESS
    if PROGRESS is None:
        PROGRESS = Progress()
    return PROGRESS


def process_image(image: Image.Image, job: "Job | None" = None) -> None:
    """Общий пайплайн: распознавание + перевод → показ результата (hotkey и live)."""
    if job is None:
        job = _new_job("hotkey")
    mode = CONFIG.get("ocr_mode", "tesseract")
    logger.info("Задача #%d (%s): снимок %dx%d, режим %s.",
                job.gen, job.kind, *image.size, mode)
    progress = _progress()
    progress.build("Распознавание…", owner=job.gen)
    progress.update(10, "Подготовка изображения…", owner=job.gen)
    on_top = _result_on_top(job)
    stream_state = {"last_push": 0.0, "text": ""}

    def push_partial(raw: str) -> None:
        """Живой показ перевода: окно обновляется по мере генерации ответа."""
        visible = _stream_visible_text(raw)
        if not visible or visible == stream_state["text"]:
            return
        now = time.time()
        if now - stream_state["last_push"] < 0.12:   # не чаще ~8 раз в секунду
            return
        stream_state["last_push"] = now
        stream_state["text"] = visible
        _result_partial(job.gen, visible, image=image, focus=job.interactive, on_top=on_top)

    job.check()
    # В режиме «нейросеть» снимок без текста отсекаем до запроса. В режиме
    # Tesseract решение принимает сам OCR — по уверенности распознавания.
    if mode == "ai" and bool(CONFIG.get("skip_no_text", True)) and not image_has_text(image):
        logger.info("Задача #%d: на снимке нет текста — пропускаю без запроса.", job.gen)
        _toast("На снимке нет текста — пропущено")
        return

    if mode == "ai":
        # Нейросеть сама читает картинку и сразу переводит — без Tesseract.
        progress.update(20, "Читаю и перевожу (нейросеть)…", owner=job.gen)
        translation, original = translate_image_via_ai(image, job=job, on_delta=push_partial)
        if not translation.strip():
            logger.info("Задача #%d: модель не нашла текст на изображении.", job.gen)
            _result_abort(job.gen)
            _toast("На снимке нет текста — пропущено")
            return
        progress.finish("Готово", owner=job.gen)
        _result_done(job.gen, image, translation, original,
                     focus=job.interactive, on_top=on_top)
        logger.info("Задача #%d: перевод (vision) готов (%d симв.).", job.gen, len(translation))
        return

    # Режим Tesseract: сначала OCR, затем перевод нейросетью.
    progress.update(25, "Распознаю текст (Tesseract)…", owner=job.gen)
    ocr = ocr_image_ex(image)
    logger.info("Задача #%d: %r", job.gen, ocr)
    if not ocr.text.strip() or not ocr.looks_like_text:
        logger.info("Задача #%d: текст не распознан (уверенность %.0f) — пропускаю.",
                    job.gen, ocr.confidence)
        _toast("На снимке нет текста — пропущено")
        return
    job.check()
    if bool(CONFIG.get("log_recognized_text")):
        logger.info("Распознанный текст:\n%s", ocr.text)
    else:
        logger.info("Распознанный текст: %d симв. (полный текст не пишется в журнал)",
                    len(ocr.text))
    progress.update(60, "Перевожу (нейросеть)…", owner=job.gen)
    translation = translate_text(ocr.text, job=job, on_delta=push_partial)
    progress.finish("Готово", owner=job.gen)
    _result_done(job.gen, image, translation, ocr.text, focus=job.interactive, on_top=on_top)
    logger.info("Задача #%d: перевод готов (%d симв.).", job.gen, len(translation))


def _result_on_top(job: Job) -> bool:
    """Держать ли окно результата поверх других окон для этой задачи."""
    if CONFIG.get("keep_image_on_top"):
        return True
    # В режиме «на лету» окно должно быть видно сразу, не отбирая фокус.
    return (not job.interactive) and bool(CONFIG.get("live_on_top", True))


def _run_job(job: Job, image: Image.Image) -> None:
    """Выполняет задачу целиком и снимает её с учёта — точка входа для потоков."""
    try:
        process_image(image, job)
    except Cancelled:
        logger.info("Задача #%d отменена.", job.gen)
        _mark_result_cancelled(job.gen)
        _toast("Обработка отменена", kind="warn")
    except RuntimeError as exc:
        logger.error("Задача #%d: %s", job.gen, exc)
        _result_abort(job.gen)
        _toast(str(exc), kind="error", seconds=6.0)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Непредвиденная ошибка в задаче #%d", job.gen)
        _result_abort(job.gen)
        _toast(f"Произошла ошибка: {exc}", kind="error", seconds=6.0)
    finally:
        _finish_job(job)
        _progress().close(owner=job.gen)


# ---------------------------------------------------------------------------
# Перевод: несколько провайдеров через единый OpenAI-совместимый интерфейс
# ---------------------------------------------------------------------------
def _build_messages(text: str) -> list[dict]:
    target = CONFIG.get("target_lang", "русский") or "русский"
    source = CONFIG.get("source_lang", "авто") or "авто"
    source_hint = "" if source.lower() in ("авто", "auto", "") else f" Исходный язык — {source}."
    system = (
        f"Ты профессиональный переводчик.{source_hint} Переведи текст на {target} язык "
        "как можно ближе к оригиналу — точно и дословно, без отсебятины, пояснений и пересказа.\n"
        "Правила:\n"
        "1. Сохраняй смысл, порядок и структуру: строки, абзацы, нумерацию, списки.\n"
        "2. Не добавляй и не убирай ничего от себя.\n"
        "3. Имена собственные, числа, единицы и коды оставляй без изменений.\n"
        "4. Верни ТОЛЬКО перевод — без кавычек, комментариев и повторения исходного текста.\n"
        "5. Если фрагмент нечитаем или не имеет смысла — оставь его как есть."
    )
    return [
        {"role": "system", "content": system},
        {"role": "user", "content": text},
    ]


def _provider_settings() -> tuple[str, str, str]:
    provider_id = CONFIG.get("active_provider", "")
    section = get_provider(provider_id)
    if section is None:
        raise RuntimeError(f"Провайдер '{provider_id}' не найден в настройках.")
    base_url = (section.get("base_url") or "").rstrip("/")
    if not base_url:
        raise RuntimeError(f"Не задан Base URL для провайдера '{section.get('name', provider_id)}'")
    url = base_url if base_url.endswith("/chat/completions") else base_url + "/chat/completions"
    model = (CONFIG.get("active_model") or "").strip()
    if not model:
        raise RuntimeError(f"Не выбрана модель для провайдера '{section.get('name', provider_id)}'")
    api_key = get_api_key(provider_id)
    return url, model, api_key


def fetch_balance() -> str:
    """Возвращает строку с балансом активного провайдера (если он это поддерживает).

    DeepSeek: GET {base_url}/user/balance -> {is_available, balance_infos:[{currency,total_balance}]}.
    Для локальных/неизвестных серверов возвращает пустую строку (баланс не применим).
    """
    provider_id = CONFIG.get("active_provider", "")
    section = get_provider(provider_id)
    if section is None:
        return ""
    base_url = (section.get("base_url") or "").rstrip("/")
    if not base_url:
        return ""
    api_key = get_api_key(provider_id)
    if not api_key:
        return "Нет API-ключа"
    # Баланс есть только у DeepSeek (и совместимых биллинговых API).
    if "deepseek" in base_url.lower():
        url = base_url + "/user/balance"
    else:
        return ""  # локальные серверы баланс не отдают
    try:
        resp = requests.get(url, headers={"Authorization": f"Bearer {api_key}"}, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        infos = data.get("balance_infos") or []
        if not infos:
            return "Баланс недоступен" if data.get("is_available") is False else ""
        parts = []
        for info in infos:
            total = info.get("total_balance")
            currency = info.get("currency", "")
            if total is not None:
                parts.append(f"{total} {currency}".strip())
        return " · ".join(parts) if parts else ""
    except requests.HTTPError as exc:
        logger.warning("Не удалось получить баланс: %s", exc)
        return ""
    except requests.RequestException as exc:
        logger.warning("Сеть недоступна при запросе баланса: %s", exc)
        return ""


def _check_response(response, model: str, *, provider_id: str = "", api_key: str = "") -> None:
    """Понятные ошибки для типовых кодов ответа сервера."""
    if response.status_code == 401:
        if not api_key or api_key == DEFAULT_DEEPSEEK_API_KEY:
            env_name = _provider_env_name(provider_id) or "DEEPSEEK_API_KEY"
            raise RuntimeError(
                "API-ключ не найден. Задайте переменную окружения "
                f"{env_name} или впишите ключ в настройках (трей → Настройки)."
            )
        raise RuntimeError(
            "Ошибка авторизации (401). Проверьте API-ключ в настройках (трей → Настройки)."
        )
    if response.status_code == 402:
        raise RuntimeError("Недостаточно средств на балансе API (402).")
    if response.status_code == 404:
        raise RuntimeError(
            f"Модель '{model}' не найдена (404). Проверьте имя модели и base_url."
        )
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        raise RuntimeError(f"Сервер вернул ошибку: {exc}\n{response.text[:500]}") from exc


def _extract_json_content(response) -> str:
    """Текст ответа из обычного (не потокового) JSON-ответа."""
    try:
        data = response.json()
        content = data["choices"][0]["message"]["content"]
    except (ValueError, KeyError, IndexError, TypeError) as exc:
        logger.error("Неожиданный ответ сервера: %s", response.text[:500])
        raise RuntimeError("Сервер вернул ответ в неожиданном формате.") from exc
    if isinstance(content, list):     # некоторые серверы отдают список частей
        content = "".join(part.get("text", "") for part in content if isinstance(part, dict))
    if not isinstance(content, str):
        logger.error("Неожиданный тип содержимого ответа: %r", type(content))
        raise RuntimeError("Сервер вернул ответ в неожиданном формате.")
    return content


def _stream_completion(response, *, job: "Job | None", on_delta) -> str:
    """Читает SSE-поток ответа и отдаёт текст по мере генерации.

    Поток позволяет показывать перевод сразу, а при отмене — просто закрыть
    соединение, не дожидаясь конца генерации.
    """
    parts: list[str] = []
    try:
        # Маленький chunk_size обязателен: библиотека читает ровно столько байт
        # за раз, и на большом буфере поток «залипает» до конца генерации.
        for raw_line in response.iter_lines(chunk_size=32, decode_unicode=False):
            if job is not None and job.cancelled():
                raise Cancelled()
            if not raw_line:
                continue
            line = raw_line.decode("utf-8", "ignore").strip()
            if not line.startswith("data:"):
                continue
            payload = line[5:].strip()
            if payload == "[DONE]":
                break
            try:
                chunk = json.loads(payload)
            except ValueError:
                continue
            for choice in chunk.get("choices") or []:
                delta = choice.get("delta") or {}
                piece = delta.get("content")
                if not piece:
                    message = choice.get("message") or {}
                    piece = message.get("content")
                if isinstance(piece, str) and piece:
                    parts.append(piece)
            if parts and on_delta is not None:
                on_delta("".join(parts))
    except Cancelled:
        raise
    except Exception as exc:  # noqa: BLE001
        if job is not None and job.cancelled():
            raise Cancelled() from exc
        raise RuntimeError(f"Обрыв соединения при получении перевода:\n{exc}") from exc
    return "".join(parts)


def _chat_completion(messages: list[dict], *, label: str, job: "Job | None" = None,
                     on_delta=None) -> str:
    """Общий вызов чат-модели (перевод — через нейросеть). Возвращает текст ответа."""
    url, model, api_key = _provider_settings()
    headers = {"Content-Type": "application/json"}
    if api_key:
        # HTTP-заголовки кодируются в latin-1: не-ASCII ключ (например, русский
        # плейсхолдер) ломает запрос — отсекаем с понятной ошибкой.
        try:
            api_key.encode("latin-1")
        except UnicodeEncodeError as exc:
            raise RuntimeError(
                "API-ключ содержит не-ASCII символы. Укажите реальный ключ "
                "(только латиница/цифры) в настройках или через переменную окружения."
            ) from exc
        headers["Authorization"] = f"Bearer {api_key}"

    want_stream = bool(CONFIG.get("stream_results", True)) and on_delta is not None
    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0.0,
        "stream": want_stream,
    }

    logger.info("%s: провайдер=%s модель=%s поток=%s",
                label, CONFIG.get("active_provider"), model, want_stream)
    if job is not None:
        job.check()
    timeout = float(CONFIG.get("request_timeout") or 120)
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=timeout,
                                 stream=want_stream)
    except requests.RequestException as exc:
        logger.error("Сеть недоступна: %s", exc)
        raise RuntimeError(f"Не удалось соединиться с сервером перевода:\n{exc}") from exc

    gen = job.gen if job is not None else 0
    if job is not None:
        _register_job_http(gen, response)
    try:
        if job is not None:
            job.check()          # отмена могла прийти, пока шёл запрос
        _check_response(response, model, provider_id=CONFIG.get("active_provider", ""),
                        api_key=api_key)
        if not want_stream:
            return _strip_model_noise(_extract_json_content(response))
        content_type = (response.headers.get("Content-Type") or "").lower()
        if "event-stream" not in content_type:
            # Сервер проигнорировал поток — читаем ответ как обычный JSON.
            return _strip_model_noise(_extract_json_content(response))
        return _strip_model_noise(_stream_completion(response, job=job, on_delta=on_delta))
    finally:
        if job is not None:
            _job_http.pop(gen, None)
        try:
            response.close()
        except Exception:  # noqa: BLE001
            pass


def translate_text(text: str, *, job: "Job | None" = None, on_delta=None) -> str:
    """Переводит распознанный текст через нейросеть. Возвращает строку перевода."""
    text = (text or "").strip()
    if not text:
        return ""
    result = _chat_completion(_build_messages(text), label=f"Перевод текста (длина={len(text)})",
                              job=job, on_delta=on_delta)
    return _split_original_vs_translation(result)[0]


# ---------------------------------------------------------------------------
# Перевод картинки напрямую нейросетью (vision): модель сама читает текст
# на изображении и возвращает готовый перевод — без Tesseract.
# ---------------------------------------------------------------------------
def _image_to_data_url(image: Image.Image, max_side: int = 1600) -> str:
    """Кодирует изображение в data:image/png;base64 для vision-запросов."""
    img = image.convert("RGB")
    w, h = img.size
    if max(w, h) > max_side:
        scale = max_side / max(w, h)
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.Resampling.LANCZOS)
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


# Маркеры блочного ответа модели. Блок «ПЕРЕВОД» идёт первым: так его видно
# в окне сразу, по мере генерации ответа.
TRANSLATION_MARKER = re.compile(r"(?im)^[ \t]*(?:перевод|translation)[ \t]*:[ \t]*")
ORIGINAL_MARKER = re.compile(r"(?im)^[ \t]*(?:оригинал|original)[ \t]*:[ \t]*")
# Ответ «на изображении нет текста».
NO_TEXT_COMPACT = {"notext", "неттекста", "неттекстанаизображении", "notextfound", "текстаненайден"}
# Служебные слова, которые могут появляться в потоке по частям.
STREAM_MARKERS = ("перевод", "translation", "оригинал", "original", "no_text", "notext")


def _build_vision_messages(image: Image.Image) -> list[dict]:
    target = CONFIG.get("target_lang", "русский") or "русский"
    source = CONFIG.get("source_lang", "авто") or "авто"
    source_hint = "" if source.lower() in ("авто", "auto", "") else f" Исходный язык — {source}."
    system = (
        f"Ты профессиональный переводчик.{source_hint} На изображении есть текст. "
        f"Считай его и переведи на {target} язык как можно ближе к оригиналу — точно и дословно, "
        "без отсебятины и пересказа.\n"
        "Ответь строго в таком формате (два блока):\n"
        "ПЕРЕВОД:\n<перевод>\n"
        "ОРИГИНАЛ:\n<текст как на картинке>\n"
        "Правила:\n"
        "1. Сохраняй структуру: строки, абзацы, нумерацию, списки.\n"
        "2. Имена, числа, единицы и коды оставляй без изменений.\n"
        "3. Не добавляй пояснений и комментариев.\n"
        "4. Если читаемого текста на изображении нет (фотография, иконка, пустая область, "
        "интерфейс без надписей) — ответь ровно NO_TEXT и ничего больше."
    )
    user = {
        "role": "user",
        "content": [
            {"type": "text", "text": "Переведи текст с картинки."},
            {"type": "image_url", "image_url": {"url": _image_to_data_url(image)}},
        ],
    }
    return [{"role": "system", "content": system}, user]


def translate_image_via_ai(image: Image.Image, *, job: "Job | None" = None,
                           on_delta=None) -> tuple[str, str]:
    """Переводит картинку напрямую vision-моделью.

    Возвращает (перевод, распознанный_оригинал). Если модель вернула и оригинал,
    и перевод — показываем оба; иначе оригинал пустой. Пустой перевод означает,
    что текста на изображении нет.
    """
    result = _chat_completion(_build_vision_messages(image), label="Перевод изображения (vision)",
                              job=job, on_delta=on_delta)
    if _is_no_text_answer(result):
        return "", ""
    return _split_original_vs_translation(result)


def _is_no_text_answer(text: str) -> bool:
    """Ответ модели «на картинке нет текста» (или пустой ответ)."""
    compact = re.sub(r"[^0-9a-zа-яё]+", "", (text or "").lower())
    if not compact:
        return True
    return compact in NO_TEXT_COMPACT


def _split_original_vs_translation(text: str) -> tuple[str, str]:
    """Выделяет из ответа модели перевод и оригинал (порядок блоков любой)."""
    text = (text or "").strip()
    if not text:
        return "", ""
    translation_match = TRANSLATION_MARKER.search(text)
    original_match = ORIGINAL_MARKER.search(text)

    if translation_match and original_match:
        first, second = sorted((translation_match, original_match), key=lambda m: m.start())
        middle = text[first.end():second.start()].strip()
        tail = text[second.end():].strip()
        if first is translation_match:
            return middle, tail
        return tail, middle
    if translation_match:
        return text[translation_match.end():].strip(), ""
    if original_match:
        # Модель ответила только блоком «ОРИГИНАЛ»: всё до него — перевод.
        return text[:original_match.start()].strip(), text[original_match.end():].strip()
    return text, ""


def _trim_partial_marker(text: str) -> str:
    """Срезает незавершённый маркер в конце потока («ОРИГ», «ПЕРЕ», «NO_T»)."""
    lines = text.split("\n")
    while lines:
        candidate = re.sub(r"[^0-9a-zа-яё_]+", "", lines[-1].strip().lower())
        if candidate and any(word.startswith(candidate) for word in STREAM_MARKERS):
            lines.pop()
            continue
        break
    return "\n".join(lines)


def _stream_visible_text(raw: str) -> str:
    """Что показывать в окне, пока ответ ещё генерируется.

    Скрывает служебные маркеры и блок «ОРИГИНАЛ», чтобы в окне рос именно перевод.
    """
    text = raw or ""
    if not text.strip():
        return ""
    original_match = ORIGINAL_MARKER.search(text)
    if original_match:
        text = text[:original_match.start()]
    translation_match = TRANSLATION_MARKER.search(text)
    if translation_match:
        text = text[translation_match.end():]
    return _trim_partial_marker(text).strip()


def _strip_model_noise(content: str) -> str:
    """Убирает служебные обёртки, которыми иногда отвечают модели."""
    text = (content or "").strip()
    # <｜end▁of▁thinking｜>-блоки рассуждений (deepseek-reasoner и др.)
    text = re.sub(r"(?is)<(?:think|thinking|reasoning)>.*?</(?:think|thinking|reasoning)>", "", text)
    text = text.strip()
    # Обрамление ``` ... ``` без указания языка
    if text.startswith("```") and text.endswith("```"):
        inner = text.strip("`").strip()
        first_newline = inner.find("\n")
        if first_newline != -1 and " " not in inner[:first_newline]:
            inner = inner[first_newline + 1:]
        text = inner.strip()
    text = re.sub(r"(?i)^перевод\s*:\s*", "", text).strip()
    return text


# ---------------------------------------------------------------------------
# Дизайн-система (токены)
#
# Мини-дизайн-система для tkinter: единые цветовые роли, шкала отступов
# (кратна 4/8), типографика, радиусы-эмуляция и повторно используемые
# компоненты (кнопки, поля, заголовки). Все окна строятся только на токенах,
# без "магических" значений — как принято в дизайн-системах.
# ---------------------------------------------------------------------------

# --- Цветовые роли ---
class Color:
    # Бренд / акцент
    PRIMARY = "#F54B64"          # основной акцент (действие)
    PRIMARY_HOVER = "#F78361"    # состояние наведения
    PRIMARY_PRESSED = "#D93B52"  # состояние нажатия
    ACCENT = "#FFD42B"           # вторичный акцент (внимание)

    # Поверхности
    SURFACE = "#4E586E"          # фон окна
    SURFACE_RAISED = "#FFFFFF"   # карточки / текст
    SURFACE_SUNKEN = "#3E4759"   # утопленные области

    # Текст
    TEXT = "#222222"             # основной текст на светлом
    TEXT_MUTED = "#6B7280"       # второстепенный
    TEXT_ON_DARK = "#FFFFFF"     # текст на тёмном фоне
    TEXT_ON_DARK_MUTED = "#C3C9D6"

    # Служебные
    BORDER = "#D8DCE6"
    SUCCESS = "#2E9E5B"
    SUCCESS_BG = "#E6F6EC"
    DANGER = "#D93B52"

# --- Шкала отступов (px), кратна 4 ---
class Space:
    XS = 4
    SM = 8
    MD = 12
    LG = 16
    XL = 24
    XXL = 32

# --- Типографика ---
class Font:
    FAMILY = "Segoe UI"
    CAPTION = (FAMILY, 9)
    BODY = (FAMILY, 11)
    BODY_BOLD = (FAMILY, 11, "bold")
    TITLE = (FAMILY, 13, "bold")
    MONO = ("Consolas", 11)

# Обратная совместимость с прежними именами.
PRIMARY_COLOR = Color.PRIMARY
PRIMARY_COLOR_2 = Color.PRIMARY_HOVER
SECONDARY_COLOR = Color.ACCENT
DARK_GREY = Color.SURFACE
WHITE = Color.SURFACE_RAISED


def _styled_entry(parent: tk.Misc, **kwargs) -> tk.Entry:
    """Поле ввода по токенам дизайн-системы."""
    return tk.Entry(
        parent,
        bg=Color.SURFACE_RAISED,
        fg=Color.TEXT,
        insertbackground=Color.TEXT,
        relief="flat",
        bd=0,
        highlightthickness=1,
        highlightbackground=Color.BORDER,
        highlightcolor=Color.PRIMARY,
        font=Font.BODY,
        **kwargs,
    )


def _button(parent: tk.Misc, text: str, command, *, variant: str = "primary", **kwargs) -> tk.Button:
    """Кнопка по токенам. Варианты: primary | secondary | ghost | danger."""
    palette = {
        "primary": (Color.PRIMARY, Color.TEXT_ON_DARK, Color.PRIMARY_HOVER),
        "secondary": (Color.ACCENT, Color.TEXT, Color.PRIMARY_HOVER),
        "ghost": (Color.SURFACE, Color.TEXT_ON_DARK, Color.PRIMARY),
        "danger": (Color.SURFACE, Color.TEXT_ON_DARK, Color.DANGER),
    }
    bg, fg, hover = palette.get(variant, palette["primary"])
    opts = {
        "text": text, "command": command,
        "bg": bg, "fg": fg, "activebackground": hover, "activeforeground": Color.TEXT_ON_DARK,
        "bd": 0, "relief": "flat", "cursor": "hand2", "font": Font.BODY,
        "padx": Space.MD + 2, "pady": Space.SM - 2, "highlightthickness": 0,
    }
    opts.update(kwargs)
    return tk.Button(parent, **opts)


# ---------------------------------------------------------------------------
# Единый Tk-поток: все окна (результат, настройки, уведомления) живут в одном
# потоке. Tkinter не потокобезопасен, поэтому UI-задачи ставим в очередь и
# выполняем их через after() в выделенном потоке.
# ---------------------------------------------------------------------------
_ui_lock = threading.Lock()
_ui_root: "tk.Tk | None" = None
_ui_thread: "threading.Thread | None" = None
_ui_queue: "queue.Queue" = queue.Queue()


def _start_ui_thread() -> None:
    global _ui_root, _ui_thread
    with _ui_lock:
        if _ui_root is not None:
            return
        started = threading.Event()

        def runner() -> None:
            global _ui_root
            root = tk.Tk()
            root.withdraw()
            try:
                ttk.Style(root).theme_use("clam")
            except Exception:  # noqa: BLE001
                pass
            _ui_root = root
            started.set()

            def pump() -> None:
                try:
                    while True:
                        func = _ui_queue.get_nowait()
                        try:
                            func()
                        except Exception:  # noqa: BLE001
                            logger.exception("Ошибка в UI-задаче")
                except queue.Empty:
                    pass
                root.after(40, pump)

            root.after(40, pump)
            root.mainloop()

        _ui_thread = threading.Thread(target=runner, name="tk-ui", daemon=True)
        _ui_thread.start()
        started.wait(timeout=10)


def run_on_ui(func) -> None:
    """Выполняет func() в едином Tk-потоке (создаёт его при необходимости)."""
    _start_ui_thread()
    _ui_queue.put(func)


def _set_topmost(win, value: bool) -> None:
    try:
        win.attributes("-topmost", bool(value))
    except Exception:  # noqa: BLE001
        pass


def _cursor_monitor_rect(win) -> tuple[int, int, int, int]:
    """Рабочая область монитора под курсором: (x, y, ширина, высота).

    На Windows учитывается панель задач и активный монитор (важно и для
    многомониторных конфигураций, и для небольших экранов); на других ОС —
    размеры основного экрана.
    """
    if IS_WINDOWS:
        try:
            import ctypes
            from ctypes import wintypes

            class POINT(ctypes.Structure):
                _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

            class RECT(ctypes.Structure):
                _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                            ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

            class MONITORINFO(ctypes.Structure):
                _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", RECT),
                            ("rcWork", RECT), ("dwFlags", wintypes.DWORD)]

            point = POINT()
            if ctypes.windll.user32.GetCursorPos(ctypes.byref(point)):
                monitor = ctypes.windll.user32.MonitorFromPoint(point, 2)  # NEAREST
                info = MONITORINFO()
                info.cbSize = ctypes.sizeof(MONITORINFO)
                if monitor and ctypes.windll.user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
                    work = info.rcWork
                    width = work.right - work.left
                    height = work.bottom - work.top
                    if width > 0 and height > 0:
                        return work.left, work.top, width, height
        except Exception as exc:  # noqa: BLE001
            logger.debug("Не удалось определить монитор под курсором: %s", exc)
    return 0, 0, win.winfo_screenwidth(), win.winfo_screenheight()


# ---------------------------------------------------------------------------
# Мини-окно прогресса («песочные часы» + проценты)
# Не блокирует обработку: живёт в UI-потоке, обновляется через PROGRESS.update().
# ---------------------------------------------------------------------------
class Progress:
    """Маленькое всегда-поверх окно с этапом, процентом и оценкой остатка времени."""

    def __init__(self) -> None:
        self.win: "tk.Toplevel | None" = None
        self._bar: "tk.Frame | None" = None
        self._label: "tk.Label | None" = None
        self._pct: "tk.Label | None" = None
        self._eta: "tk.Label | None" = None
        self._canvas: "tk.Canvas | None" = None
        self._cancel_btn: "tk.Button | None" = None
        self._spinner_angle = 0
        self._spinner_job: "str | None" = None
        self._tick_job: "str | None" = None
        # Номер задачи, которой принадлежит плашка: обновления чужой задачи
        # (например, уже отменённой) не должны её трогать.
        self._owner: "int | None" = None
        # Состояние прогресса
        self._start = time.time()
        self._percent = 0.0          # целевой процент
        self._shown = 0.0            # отображаемый (плавно догоняет целевой)
        self._stage = "Обработка…"

    def _belongs_to(self, owner: "int | None") -> bool:
        """True, если плашкой управляет именно эта задача."""
        return owner is None or self._owner == owner

    def show(self, title: str = "Обработка…", owner: "int | None" = None) -> None:
        self._start = time.time()
        self._percent = 0.0
        self._shown = 0.0
        self._stage = title
        self._owner = owner
        run_on_ui(lambda: self._build(title))

    def build(self, title: str = "Обработка…", owner: "int | None" = None) -> None:
        """Создаёт окно, только если его ещё нет; иначе просто меняет этап."""
        self._start = time.time()
        self._percent = 0.0
        self._shown = 0.0
        self._stage = title
        self._owner = owner
        run_on_ui(lambda: self._ensure_built(title, owner))

    def _ensure_built(self, title: str, owner: "int | None") -> None:
        if not self._belongs_to(owner):
            return
        if self.win is not None and self.win.winfo_exists():
            if self._label is not None:
                self._label.configure(text=title)
            return
        self._build(title)

    def _build(self, title: str) -> None:
        # Защита от дублирования: если окно уже есть — закрываем старое.
        if self.win is not None:
            try:
                self.win.destroy()
            except tk.TclError:
                pass
            self.win = None

        win = tk.Toplevel(_ui_root)
        win.title("pyAutoImgTranslate")
        win.configure(bg=Color.SURFACE)
        win.attributes("-topmost", True)
        win.resizable(False, False)
        win.overrideredirect(True)  # без системной рамки — компактная плашка
        self.win = win

        frame = tk.Frame(win, bg=Color.SURFACE, highlightthickness=1,
                         highlightbackground=Color.PRIMARY)
        frame.pack(fill="both", expand=True)

        header = tk.Frame(frame, bg=Color.SURFACE)
        header.pack(fill="x", padx=Space.MD, pady=(Space.MD, Space.SM))

        # Кнопка отмены: обработку можно прервать в любой момент.
        self._cancel_btn = _button(header, "Отмена", lambda: _cancel_from_progress(self),
                                   variant="danger", font=Font.CAPTION,
                                   padx=Space.SM, pady=Space.XS - 1)
        self._cancel_btn.pack(side="right", padx=(Space.SM, 0))

        # «Песочные часы» — рисуем в Canvas
        self._canvas = tk.Canvas(header, width=22, height=22, bg=Color.SURFACE,
                                 highlightthickness=0)
        self._canvas.pack(side="left", padx=(0, Space.SM))

        text_col = tk.Frame(header, bg=Color.SURFACE)
        text_col.pack(side="left", fill="x", expand=True)
        self._label = tk.Label(text_col, text=title, bg=Color.SURFACE, fg=Color.TEXT_ON_DARK,
                               font=Font.BODY_BOLD, anchor="w")
        self._label.pack(anchor="w")
        self._pct = tk.Label(text_col, text="0%", bg=Color.SURFACE, fg=Color.TEXT_ON_DARK_MUTED,
                             font=Font.CAPTION, anchor="w")
        self._pct.pack(anchor="w")
        self._eta = tk.Label(text_col, text="", bg=Color.SURFACE, fg=Color.TEXT_ON_DARK_MUTED,
                             font=Font.CAPTION, anchor="w")
        self._eta.pack(anchor="w")

        # progress bar
        track = tk.Frame(frame, bg=Color.SURFACE_SUNKEN, height=6)
        track.pack(fill="x", padx=Space.MD, pady=(0, Space.MD))
        track.pack_propagate(False)
        self._bar = tk.Frame(track, bg=Color.PRIMARY, height=6)
        self._bar.place(x=0, y=0, relwidth=0.0, relheight=1.0)

        self._place_bottom_center(win)
        self._animate()

    @staticmethod
    def _place_bottom_center(win: tk.Toplevel, margin: int = 48) -> None:
        """Ставит плашку по центру внизу — не перекрывает область выделения."""
        win.update_idletasks()
        w = max(win.winfo_width(), win.winfo_reqwidth())
        h = win.winfo_reqheight()
        area_x, area_y, area_w, area_h = _cursor_monitor_rect(win)
        x = area_x + max(0, (area_w - w) // 2)
        y = area_y + max(0, area_h - h - margin)
        win.geometry(f"+{x}+{y}")

    def _animate(self) -> None:
        """Спиннер + плавное доведение процента + пересчёт ETA, ~10 раз в секунду."""
        if not self.win or not self.win.winfo_exists():
            return
        try:
            assert self._canvas is not None
            self._canvas.delete("all")
            self._canvas.create_arc(2, 2, 20, 20, start=self._spinner_angle,
                                    extent=280, style="arc", outline=Color.PRIMARY, width=3)
            self._spinner_angle = (self._spinner_angle - 20) % 360

            # Плавно доводим отображаемый процент к целевому (не перескакивает).
            if self._shown < self._percent:
                self._shown = min(self._percent, self._shown + max(0.5, (self._percent - self._shown) * 0.2))

            elapsed = time.time() - self._start
            if self._bar is not None:
                self._bar.place_configure(relwidth=self._shown / 100.0)
            if self._pct is not None:
                self._pct.configure(text=f"{int(self._shown)}%")
            if self._eta is not None:
                self._eta.configure(text=self._eta_text(elapsed, self._shown))

            self._spinner_job = self.win.after(100, self._animate)
        except tk.TclError:
            pass

    @staticmethod
    def _eta_text(elapsed: float, percent: float) -> str:
        """Строка с прошедшим и оставшимся временем (оценка по текущей скорости)."""
        if percent <= 1.0:
            return f"прошло {int(elapsed)} с"
        if percent >= 99.0:
            return f"завершено за {int(elapsed)} с"
        remaining = elapsed * (100.0 - percent) / percent
        if remaining < 1:
            return f"{int(elapsed)} с · осталось <1 с"
        if remaining < 90:
            return f"{int(elapsed)} с · осталось ~{int(remaining)} с"
        return f"{int(elapsed)} с · осталось ~{int(remaining / 60)} мин"

    def update(self, percent: float, stage: str | None = None,
               owner: "int | None" = None) -> None:
        if not self._belongs_to(owner):
            return
        run_on_ui(lambda: self._update(percent, stage, owner))

    def finish(self, stage: str = "Готово", owner: "int | None" = None) -> None:
        """Доводит шкалу до 100% и показывает финальный этап."""
        if not self._belongs_to(owner):
            return
        run_on_ui(lambda: self._update(100, stage, owner))

    def _update(self, percent: float, stage: str | None, owner: "int | None" = None) -> None:
        if not self._belongs_to(owner):
            return
        # Целевой процент растёт только вперёд (не даём шкале прыгать назад).
        self._percent = max(self._percent, max(0.0, min(100.0, float(percent))))
        if stage:
            self._stage = stage
        if not self.win or not self.win.winfo_exists():
            return
        try:
            if stage and self._label is not None:
                self._label.configure(text=stage)
        except tk.TclError:
            pass

    def close(self, owner: "int | None" = None) -> None:
        """Закрывает плашку; чужая задача её не трогает."""
        if not self._belongs_to(owner):
            return
        self._owner = None
        run_on_ui(self._close)

    def _close(self) -> None:
        for job in (self._spinner_job, self._tick_job):
            if job and self.win:
                try:
                    self.win.after_cancel(job)
                except Exception:  # noqa: BLE001
                    pass
        self._spinner_job = None
        self._tick_job = None
        if self.win:
            try:
                self.win.destroy()
            except tk.TclError:
                pass
        self.win = None
        self._bar = None
        self._label = None
        self._pct = None
        self._eta = None
        self._canvas = None
        self._cancel_btn = None


def _cancel_from_progress(progress: "Progress") -> None:
    """Нажатие «Отмена» в плашке прогресса."""
    if cancel_current_job("кнопка «Отмена»") and progress._label is not None:
        try:
            progress._label.configure(text="Отмена…")
        except tk.TclError:
            pass


def _segmented(parent: tk.Misc, options: list[tuple[str, str]], var: tk.StringVar,
               command=None) -> tk.Frame:
    """Сегментированный переключатель (как radio-group в дизайн-системе)."""
    frame = tk.Frame(parent, bg=Color.SURFACE_SUNKEN)
    buttons = {}

    def select(value: str) -> None:
        var.set(value)
        for val, btn in buttons.items():
            active = val == value
            btn.configure(bg=Color.PRIMARY if active else Color.SURFACE_SUNKEN,
                          fg=Color.TEXT_ON_DARK if active else Color.TEXT_ON_DARK_MUTED)
        if command:
            command()

    for value, label in options:
        btn = tk.Button(frame, text=label, command=lambda v=value: select(v),
                        bd=0, relief="flat", font=Font.CAPTION, cursor="hand2",
                        padx=Space.MD, pady=Space.XS, highlightthickness=0,
                        bg=Color.SURFACE_SUNKEN, fg=Color.TEXT_ON_DARK_MUTED,
                        activebackground=Color.PRIMARY_HOVER, activeforeground=Color.TEXT_ON_DARK)
        btn.pack(side="left")
        buttons[value] = btn
    select(var.get())
    return frame


# ---------------------------------------------------------------------------
# Окно результата: картинка + перевод.
#
# Окно одно на всё приложение. Перевод печатается по мере генерации ответа,
# поэтому окно появляется сразу, а текст в нём растёт. Размеры считаются от
# рабочей области монитора под курсором — на небольших и на многомониторных
# конфигурациях окно остаётся целиком видимым.
# ---------------------------------------------------------------------------
_RESULT: dict = {
    "win": None,          # tk.Toplevel
    "text": None,         # tk.Text с переводом
    "title": None,        # tk.Label заголовка
    "cancel": None,       # кнопка «Отменить» (показывается во время обработки)
    "var": None,          # StringVar переключателя «Перевод / Оригинал»
    "gen": None,          # номер задачи, которой принадлежит окно
    "mode": "translation",
    "translation": "",
    "source": "",
    "streaming": False,
}


def _result_title() -> str:
    mode_label = "нейросеть" if CONFIG.get("ocr_mode", "ai") == "ai" else "Tesseract"
    return f"Перевод · {mode_label}"


def _result_alive() -> bool:
    win = _RESULT.get("win")
    try:
        return bool(win is not None and win.winfo_exists())
    except tk.TclError:
        return False


def _text_columns(area_w: int, image_w: int, stacked: bool) -> int:
    """Сколько символов в строке помещается в оставшемся месте."""
    char_width = 8                        # ≈ ширина символа в Segoe UI 11
    available = area_w - 170 - (0 if stacked else image_w + Space.LG)
    return max(20, min(74, available // char_width))


def _text_rows(area_h: int, image_h: int, stacked: bool) -> int:
    """Сколько строк текста помещается по высоте."""
    line_height = 18
    available = area_h - (230 + image_h if stacked else 210)
    return max(5, min(26, available // line_height))


def _open_result_window(image: Image.Image, gen: int, *, focus: bool, on_top: bool,
                        streaming: bool) -> None:
    """Создаёт окно результата (вызывается только в UI-потоке)."""
    global _result_window
    if _result_alive():
        try:
            _RESULT["win"].destroy()
        except tk.TclError:
            pass
    _result_window = None
    _RESULT.update({"win": None, "text": None, "title": None, "cancel": None, "var": None,
                    "mode": "translation", "translation": "", "source": "",
                    "streaming": streaming, "gen": gen})

    win = tk.Toplevel(_ui_root)
    _result_window = win
    _RESULT["win"] = win
    win.title("Перевод с экрана")
    win.configure(bg=Color.SURFACE)
    win.attributes("-topmost", bool(on_top))
    win.bind("<Escape>", lambda _e: win.destroy())

    area_x, area_y, area_w, area_h = _cursor_monitor_rect(win)
    win.maxsize(area_w, area_h)
    stacked = area_w < 900                # узкий экран — текст под картинкой
    if stacked:
        img_max_w, img_max_h = max(120, int(area_w * 0.6)), max(90, int(area_h * 0.34))
    else:
        img_max_w = max(120, min(460, int(area_w * 0.34)))
        img_max_h = max(120, min(640, area_h - 280))
    preview = _fit_for_preview(image, max_w=img_max_w, max_h=img_max_h)

    outer = tk.Frame(win, bg=Color.SURFACE)
    outer.pack(fill="both", expand=True, padx=Space.LG, pady=(Space.LG, Space.SM))

    img_label = tk.Label(outer, bg=Color.SURFACE, bd=0, highlightthickness=0)
    img_tk = ImageTk.PhotoImage(preview)
    img_label.configure(image=img_tk)
    img_label.image = img_tk  # type: ignore[attr-defined]  # защита от сборщика мусора

    right = tk.Frame(outer, bg=Color.SURFACE)
    if stacked:
        img_label.pack(anchor="w")
        right.pack(fill="both", expand=True, pady=(Space.MD, 0))
    else:
        img_label.pack(side="left", anchor="n")
        right.pack(side="left", fill="both", expand=True, padx=(Space.LG, 0))
    right.rowconfigure(2, weight=1)
    right.columnconfigure(0, weight=1)

    title = tk.Label(right, text=_result_title(), bg=Color.SURFACE,
                     fg=Color.TEXT_ON_DARK, font=Font.TITLE, anchor="w")
    title.grid(row=0, column=0, sticky="ew")
    _RESULT["title"] = title

    view_var = tk.StringVar(value="translation")
    _RESULT["var"] = view_var

    def on_view() -> None:
        _RESULT["mode"] = view_var.get()
        _render_result(gen)

    seg = _segmented(right, [("translation", "Перевод"), ("source", "Оригинал")],
                     view_var, command=on_view)
    seg.grid(row=1, column=0, sticky="w", pady=(Space.SM, Space.SM))

    # Карточка текста.
    card = tk.Frame(right, bg=Color.SURFACE_RAISED, highlightthickness=1,
                    highlightbackground=Color.BORDER)
    card.grid(row=2, column=0, sticky="nsew")
    card.rowconfigure(0, weight=1)
    card.columnconfigure(0, weight=1)

    inner = tk.Frame(card, bg=Color.SURFACE_RAISED)
    inner.grid(row=0, column=0, sticky="nsew", padx=Space.MD, pady=Space.MD)
    inner.rowconfigure(0, weight=1)
    inner.columnconfigure(0, weight=1)

    scrollbar = tk.Scrollbar(inner, orient=tk.VERTICAL)
    scrollbar.grid(row=0, column=1, sticky="ns")

    text_widget = tk.Text(
        inner, wrap="word", yscrollcommand=scrollbar.set,
        bg=Color.SURFACE_RAISED, fg=Color.TEXT, font=Font.BODY,
        bd=0, highlightthickness=0, padx=Space.SM, pady=Space.SM,
        width=_text_columns(area_w, preview.width, stacked),
        height=_text_rows(area_h, preview.height, stacked),
        cursor="arrow",
        selectbackground=Color.PRIMARY, selectforeground=Color.TEXT_ON_DARK,
    )
    text_widget.grid(row=0, column=0, sticky="nsew")
    scrollbar.config(command=text_widget.yview)
    _RESULT["text"] = text_widget

    def current_text() -> str:
        if _RESULT["mode"] == "source":
            return _RESULT["source"] or ""
        return _RESULT["translation"] or ""

    def copy_all() -> None:
        win.clipboard_clear()
        win.clipboard_append(current_text())
        logger.info("Скопировано в буфер обмена (%s).", _RESULT["mode"])

    def copy_selected() -> None:
        try:
            selected = text_widget.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            return
        win.clipboard_clear()
        win.clipboard_append(selected)

    def open_settings() -> None:
        _set_topmost(win, False)
        open_settings_window(win)

    def on_cancel() -> None:
        if cancel_current_job("кнопка в окне перевода"):
            title.configure(text=f"{_result_title()} · отмена…")
            logger.info("Отмена из окна перевода.")

    # --- Панель действий ---
    actions = tk.Frame(win, bg=Color.SURFACE)
    actions.pack(fill="x", padx=Space.LG, pady=(0, Space.LG))
    _button(actions, "Копировать", copy_all, variant="primary").pack(side="left", padx=(0, Space.SM))
    _button(actions, "Настройки", open_settings, variant="secondary").pack(side="left", padx=(0, Space.SM))
    cancel_btn = _button(actions, "Отменить", on_cancel, variant="danger")
    _RESULT["cancel"] = cancel_btn
    if streaming:
        cancel_btn.pack(side="left", padx=(0, Space.SM))
    _button(actions, "Закрыть", win.destroy, variant="ghost").pack(side="right")

    # Контекстное меню текста
    context_menu = tk.Menu(text_widget, tearoff=0)
    context_menu.add_command(label="Копировать", command=copy_selected)
    context_menu.add_command(label="Копировать всё", command=copy_all)
    text_widget.bind("<Button-3>", lambda e: context_menu.tk_popup(e.x_root, e.y_root))

    # Горячие клавиши окна.
    win.bind("<Control-c>", lambda _e: copy_all())
    win.bind("<Control-o>", lambda _e: (
        view_var.set("source" if view_var.get() == "translation" else "translation"), on_view()))

    _render_result(gen, autoscroll=False)
    _center_window(win, focus=focus)


def _render_result(gen: int, *, autoscroll: bool = False) -> None:
    """Перерисовывает текст окна (только UI-поток)."""
    if _RESULT.get("gen") != gen or not _result_alive():
        return
    widget = _RESULT.get("text")
    if widget is None:
        return
    text = (_RESULT["source"] if _RESULT["mode"] == "source" else _RESULT["translation"]) or ""
    text = text.strip()
    if not text:
        text = "…" if _RESULT["streaming"] else "— пусто —"
    try:
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert(tk.END, text)
        widget.configure(state="disabled")
        if autoscroll and _RESULT["mode"] == "translation":
            widget.see(tk.END)
    except tk.TclError:
        pass


def _hide_result_cancel(gen: int) -> None:
    if _RESULT.get("gen") != gen:
        return
    button = _RESULT.get("cancel")
    if button is None:
        return
    try:
        button.pack_forget()
    except tk.TclError:
        pass


def _result_partial(gen: int, text: str, *, image: Image.Image, focus: bool,
                    on_top: bool) -> None:
    """Промежуточный перевод: окно появляется на первом же куске ответа."""
    if not _is_active_job(gen):
        return

    def build() -> None:
        if _RESULT.get("gen") != gen or not _result_alive():
            _open_result_window(image, gen, focus=focus, on_top=on_top, streaming=True)
        if _RESULT.get("gen") != gen:
            return
        _RESULT["translation"] = text
        _render_result(gen, autoscroll=True)

    run_on_ui(build)


def _result_done(gen: int, image: Image.Image, translation: str, source_text: str = "",
                 *, focus: bool = True, on_top: bool = False) -> None:
    """Финальный перевод: окно создаётся, если поток так и не начался."""
    if not _is_active_job(gen):
        return

    def build() -> None:
        if _RESULT.get("gen") != gen or not _result_alive():
            _open_result_window(image, gen, focus=focus, on_top=on_top, streaming=False)
        if _RESULT.get("gen") != gen:
            return
        _RESULT["translation"] = translation
        _RESULT["source"] = source_text or ""
        _RESULT["streaming"] = False
        _RESULT["mode"] = "translation"
        var = _RESULT.get("var")
        if var is not None:
            var.set("translation")
        title = _RESULT.get("title")
        if title is not None:
            title.configure(text=_result_title())
        _render_result(gen)
        _hide_result_cancel(gen)

    run_on_ui(build)


def _mark_result_cancelled(gen: int) -> None:
    """Помечает окно как «отменено»: частичный перевод остаётся на экране."""
    def build() -> None:
        if _RESULT.get("gen") != gen or not _result_alive():
            return
        _RESULT["streaming"] = False
        title = _RESULT.get("title")
        if title is not None:
            try:
                title.configure(text=f"{_result_title()} · отменено")
            except tk.TclError:
                pass
        _render_result(gen)
        _hide_result_cancel(gen)

    run_on_ui(build)


def _result_abort(gen: int) -> None:
    """Закрывает окно, если показывать нечего (текста не нашлось или ошибка)."""
    def build() -> None:
        if _RESULT.get("gen") != gen:
            return
        if (_RESULT.get("translation") or "").strip() and _result_alive():
            # Частичный перевод полезнее, чем ничего — оставляем окно.
            _RESULT["streaming"] = False
            _hide_result_cancel(gen)
            return
        if _result_alive():
            try:
                _RESULT["win"].destroy()
            except tk.TclError:
                pass
        _RESULT["win"] = None

    run_on_ui(build)


def _fit_for_preview(image: Image.Image, max_w: int, max_h: int) -> Image.Image:
    w, h = image.size
    if w <= max_w and h <= max_h:
        return image
    scale = min(max_w / w, max_h / h)
    return image.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.Resampling.LANCZOS)


def _center_window(win, *, focus: bool = True) -> None:
    """Ставит окно по центру монитора под курсором, не давая вылезти за экран."""
    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    area_x, area_y, area_w, area_h = _cursor_monitor_rect(win)
    x = area_x + max(0, (area_w - w) // 2)
    y = area_y + max(0, (area_h - h) // 3)
    win.geometry(f"+{x}+{y}")
    if focus:
        try:
            win.focus_force()
        except tk.TclError:
            pass


# ---------------------------------------------------------------------------
# Окно настроек
# ---------------------------------------------------------------------------
# Ссылка на единственное открытое окно настроек (защита от дублирования).
_settings_window: "tk.Toplevel | None" = None

# Ссылка на единственное открытое окно результата.
_result_window: "tk.Toplevel | None" = None


def open_settings_window(parent: tk.Misc | None = None) -> None:
    """Окно настроек: провайдеры, модели, ключи — всё редактируется и применяется на лету.

    Окно всегда одно: повторный вызов из трея или из окна результата
    не создаёт второе, а поднимает и фокусирует уже открытое.
    parent=None — самостоятельное окно (из трея); иначе — дочернее окно результата.
    """

    def build() -> None:
        global _settings_window
        # Если окно уже открыто — поднимаем его и выходим.
        if _settings_window is not None:
            try:
                if _settings_window.winfo_exists():
                    _set_topmost(_settings_window, True)
                    _settings_window.deiconify()
                    _settings_window.lift()
                    _settings_window.focus_force()
                    return
            except tk.TclError:
                pass
            _settings_window = None

        root = _ui_root if parent is None else parent
        win = tk.Toplevel(root)
        _settings_window = win
        win.title(f"Настройки перевода · {APP_VERSION}")
        win.configure(bg=Color.SURFACE)
        win.attributes("-topmost", True)
        win.minsize(640, 500)

        def on_close() -> None:
            global _settings_window
            _settings_window = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)

        pad = {"padx": Space.SM, "pady": Space.XS + 1}
        label_opts = {"bg": Color.SURFACE, "fg": Color.TEXT_ON_DARK, "anchor": "w", "font": Font.BODY}
        check_opts = {
            "bg": Color.SURFACE, "fg": Color.TEXT_ON_DARK, "selectcolor": Color.SURFACE_SUNKEN,
            "activebackground": Color.SURFACE, "activeforeground": Color.TEXT_ON_DARK,
            "anchor": "w", "font": Font.BODY,
        }

        # Рабочая копия конфига: правки идут сюда, на диск — по «Сохранить».
        draft = json.loads(json.dumps(CONFIG))

        # Режим распознавания (нужен раньше предупреждения о vision).
        OCR_MODE_LABELS = {"tesseract": "Tesseract (локально)", "ai": "Нейросеть (vision, точнее)"}
        reverse_ocr = {v: k for k, v in OCR_MODE_LABELS.items()}
        ocr_mode_var = tk.StringVar(
            value=OCR_MODE_LABELS.get(CONFIG.get("ocr_mode", "tesseract"), OCR_MODE_LABELS["tesseract"]))

        # ------------------------------------------------------------------
        # Верхний блок: выбор активного провайдера и модели (с переподключением)
        # ------------------------------------------------------------------
        top = tk.Frame(win, bg=Color.SURFACE)
        top.pack(fill="x", **pad)
        tk.Label(top, text="Активный провайдер:", **label_opts).pack(side="left")
        active_provider_var = tk.StringVar()
        active_provider_box = ttk.Combobox(top, state="readonly", width=26, textvariable=active_provider_var)
        active_provider_box.pack(side="left", padx=Space.XS + 2)
        tk.Label(top, text="Модель:", **label_opts).pack(side="left", padx=(Space.MD, 0))
        active_model_var = tk.StringVar()
        active_model_box = ttk.Combobox(top, width=26, textvariable=active_model_var)
        active_model_box.pack(side="left", padx=Space.XS + 2)

        # Предупреждение о vision (свободный выбор модели)
        vision_warn = tk.Label(win, text="", bg=Color.SURFACE, fg=Color.ACCENT,
                               anchor="w", justify="left", wraplength=600, font=Font.CAPTION)
        vision_warn.pack(fill="x", padx=Space.MD + 2)

        # Баланс нейросети (для облачных провайдеров)
        balance_var = tk.StringVar(value="")
        balance_row = tk.Frame(win, bg=Color.SURFACE)
        balance_row.pack(fill="x", padx=Space.MD + 2)
        balance_label = tk.Label(balance_row, textvariable=balance_var, bg=Color.SURFACE,
                                 fg=Color.TEXT_ON_DARK_MUTED, anchor="w", font=Font.CAPTION)
        balance_label.pack(side="left")

        def refresh_balance() -> None:
            balance_var.set("Баланс: …")

            def worker() -> None:
                value = fetch_balance()
                text = f"Баланс API: {value}" if value else "Баланс: не применимо"
                try:
                    win.after(0, lambda: balance_var.set(text))
                except tk.TclError:
                    pass

            threading.Thread(target=worker, daemon=True).start()

        _button(balance_row, "Обновить", refresh_balance, variant="ghost",
                font=Font.CAPTION, padx=Space.SM, pady=1).pack(side="left", padx=(Space.SM, 0))

        def update_vision_warning(*_a) -> None:
            if reverse_ocr.get(ocr_mode_var.get(), "ai") != "ai":
                vision_warn.configure(text="")
                return
            model = active_model_var.get().strip()
            if model and not model_supports_vision(draft.get("active_provider", ""), model):
                vision_warn.configure(
                    text=f"⚠ Модель «{model}» может не поддерживать изображения. "
                         "Для распознавания картинок выберите vision-модель (например, deepseek-flash) "
                         "или переключитесь на Tesseract.")
            else:
                vision_warn.configure(text="")

        # ------------------------------------------------------------------
        # Список провайдеров + редактор выбранного
        # ------------------------------------------------------------------
        body = tk.Frame(win, bg=Color.SURFACE)
        body.pack(fill="both", expand=True, **pad)

        left = tk.Frame(body, bg=Color.SURFACE)
        left.pack(side="left", fill="y")
        tk.Label(left, text="Все провайдеры", **label_opts).pack(anchor="w")
        providers_list = tk.Listbox(left, height=8, width=26, bg=Color.SURFACE_RAISED, fg=Color.TEXT,
                                    selectbackground=Color.PRIMARY, selectforeground=Color.TEXT_ON_DARK,
                                    bd=0, highlightthickness=0, activestyle="none", font=Font.BODY)
        providers_list.pack(fill="y", expand=True, pady=Space.XS)
        prov_btns = tk.Frame(left, bg=Color.SURFACE)
        prov_btns.pack(fill="x")

        right = tk.Frame(body, bg=Color.SURFACE)
        right.pack(side="left", fill="both", expand=True, padx=(Space.MD, 0))
        right.columnconfigure(1, weight=1)

        def rlabel(text: str, row: int) -> None:
            tk.Label(right, text=text, **label_opts).grid(row=row, column=0, sticky="w", **pad)

        rlabel("Название", 0)
        name_entry = _styled_entry(right, width=36)
        name_entry.grid(row=0, column=1, sticky="ew", **pad)

        rlabel("ID (латиница)", 1)
        id_entry = _styled_entry(right, width=36)
        id_entry.grid(row=1, column=1, sticky="ew", **pad)

        rlabel("Base URL", 2)
        url_entry = _styled_entry(right, width=36)
        url_entry.grid(row=2, column=1, sticky="ew", **pad)

        rlabel("API-ключ", 3)
        key_entry = _styled_entry(right, width=36, show="•")
        key_entry.grid(row=3, column=1, sticky="ew", **pad)

        rlabel("Env-переменная ключа", 4)
        env_entry = _styled_entry(right, width=36)
        env_entry.grid(row=4, column=1, sticky="ew", **pad)

        rlabel("Модели", 5)
        models_frame = tk.Frame(right, bg=Color.SURFACE)
        models_frame.grid(row=5, column=1, sticky="ew", **pad)
        models_frame.columnconfigure(0, weight=1)
        models_list = tk.Listbox(models_frame, height=5, bg=Color.SURFACE_RAISED, fg=Color.TEXT,
                                 selectbackground=Color.PRIMARY, selectforeground=Color.TEXT_ON_DARK,
                                 bd=0, highlightthickness=0, activestyle="none", exportselection=False,
                                 font=Font.BODY)
        models_list.grid(row=0, column=0, sticky="ew")
        models_btns = tk.Frame(models_frame, bg=Color.SURFACE)
        models_btns.grid(row=0, column=1, sticky="ns", padx=(Space.XS, 0))

        # Общие настройки (языки, автозапуск) — внизу, чтобы не мешать.
        common = tk.Frame(win, bg=Color.SURFACE)
        common.pack(fill="x", **pad)

        tk.Label(common, text="Распознавание:", **label_opts).grid(row=0, column=0, sticky="w")
        ocr_mode_box = ttk.Combobox(common, state="readonly", width=24, textvariable=ocr_mode_var,
                                    values=list(OCR_MODE_LABELS.values()))
        ocr_mode_box.grid(row=0, column=1, sticky="w", padx=Space.XS)
        tk.Label(common, text="Языки OCR:", **label_opts).grid(row=0, column=2, sticky="w", padx=(Space.MD, 0))
        ocr_entry = _styled_entry(common, width=14)
        ocr_entry.insert(0, CONFIG.get("ocr_lang", "eng+rus"))
        ocr_entry.grid(row=0, column=3, sticky="w", padx=Space.XS)
        tk.Label(common, text="Целевой язык:", **label_opts).grid(row=1, column=0, sticky="w")
        target_entry = _styled_entry(common, width=14)
        target_entry.insert(0, CONFIG.get("target_lang", "русский"))
        target_entry.grid(row=1, column=1, sticky="w", padx=Space.XS)

        autostart_var = tk.BooleanVar(value=is_autostart_enabled())
        tk.Checkbutton(common, text="Автозапуск при входе в систему", variable=autostart_var,
                       **check_opts).grid(row=2, column=0, columnspan=2, sticky="w", pady=(Space.XS, 0))
        on_top_var = tk.BooleanVar(value=bool(CONFIG.get("keep_image_on_top")))
        tk.Checkbutton(common, text="Окно перевода поверх других", variable=on_top_var,
                       **check_opts).grid(row=2, column=2, columnspan=2, sticky="w", pady=(Space.XS, 0))
        live_on_top_var = tk.BooleanVar(value=bool(CONFIG.get("live_on_top", True)))
        tk.Checkbutton(common, text="Окно «на лету» поверх других", variable=live_on_top_var,
                       **check_opts).grid(row=3, column=2, columnspan=2, sticky="w", pady=(Space.XS, 0))
        skip_no_text_var = tk.BooleanVar(value=bool(CONFIG.get("skip_no_text", True)))
        tk.Checkbutton(common, text="Пропускать снимки без текста", variable=skip_no_text_var,
                       **check_opts).grid(row=3, column=0, columnspan=2, sticky="w", pady=(Space.XS, 0))
        stream_var = tk.BooleanVar(value=bool(CONFIG.get("stream_results", True)))
        tk.Checkbutton(common, text="Перевод по мере генерации", variable=stream_var,
                       **check_opts).grid(row=4, column=0, columnspan=2, sticky="w", pady=(Space.XS, 0))

        # Режим выделения области + подсказка о действующей клавише.
        capture_var = tk.StringVar(value=capture_mode_of(draft))
        capture_row = tk.Frame(common, bg=Color.SURFACE)
        capture_row.grid(row=5, column=0, columnspan=4, sticky="w", pady=(Space.SM, 0))
        tk.Label(capture_row, text="Выделение области:", **label_opts).pack(side="left")
        capture_hint = tk.Label(capture_row, text="", bg=Color.SURFACE, fg=Color.TEXT_ON_DARK_MUTED,
                                font=Font.CAPTION)

        def update_capture_hint(*_args) -> None:
            hint = hotkey_hint(capture_var.get(), draft.get("hotkey", ""))
            if capture_var.get() != capture_mode():
                hint += " · применится после перезапуска"
            capture_hint.configure(text=hint)

        capture_seg = _segmented(
            capture_row, [("builtin", "Своё окно (как Win+Shift+S)"), ("system", "Системное")],
            capture_var, command=update_capture_hint)
        capture_seg.pack(side="left", padx=Space.SM)
        capture_hint.pack(side="left", padx=(Space.SM, 0))

        status = tk.Label(win, text="", bg=Color.SURFACE, fg=Color.ACCENT, anchor="w",
                          justify="left", wraplength=600, font=Font.CAPTION)
        status.pack(fill="x", padx=Space.MD)

        # ------------------------------------------------------------------
        # Работа с draft-конфигом
        # ------------------------------------------------------------------
        selected_id = {"value": draft.get("active_provider") or ""}

        def provider_ids() -> list[str]:
            return [p["id"] for p in draft["providers"]]

        def refresh_active_boxes() -> None:
            names = [p.get("name", p["id"]) for p in draft["providers"]]
            active_provider_box.configure(values=names)
            current = get_provider(draft.get("active_provider", ""))
            if current:
                active_provider_var.set(current.get("name", current["id"]))
                models = current.get("models") or []
                active_model_box.configure(values=models)
                if draft.get("active_model") not in models:
                    draft["active_model"] = models[0] if models else ""
                active_model_var.set(draft.get("active_model") or "")
            update_vision_warning()

        def refresh_providers_list() -> None:
            providers_list.delete(0, tk.END)
            for p in draft["providers"]:
                providers_list.insert(tk.END, p.get("name", p["id"]))
            ids = provider_ids()
            if selected_id["value"] in ids:
                idx = ids.index(selected_id["value"])
                providers_list.selection_clear(0, tk.END)
                providers_list.selection_set(idx)
                providers_list.activate(idx)
            refresh_active_boxes()

        def refresh_models_list() -> None:
            models_list.delete(0, tk.END)
            section = get_provider(selected_id["value"], draft)
            for m in (section or {}).get("models", []):
                models_list.insert(tk.END, m)

        def load_selected_into_fields() -> None:
            section = get_provider(selected_id["value"], draft)
            for entry, value in ((name_entry, section.get("name", "") if section else ""),
                                 (id_entry, section.get("id", "") if section else ""),
                                 (url_entry, section.get("base_url", "") if section else ""),
                                 (key_entry, section.get("api_key", "") if section else ""),
                                 (env_entry, section.get("api_key_env", "") if section else "")):
                entry.delete(0, tk.END)
                entry.insert(0, value or "")
            refresh_models_list()

        def on_provider_select(_event=None) -> None:
            sel = providers_list.curselection()
            if not sel:
                return
            selected_id["value"] = provider_ids()[sel[0]]
            load_selected_into_fields()

        def commit_fields_to_draft() -> None:
            """Переносит правки полей в текущего провайдера draft."""
            section = get_provider(selected_id["value"], draft)
            if section is None:
                return
            new_id = id_entry.get().strip() or section["id"]
            # не допускаем дубликатов id
            if new_id != section["id"] and new_id in provider_ids():
                new_id = section["id"]
            section["id"] = new_id
            section["name"] = name_entry.get().strip() or new_id
            section["base_url"] = url_entry.get().strip()
            section["api_key"] = key_entry.get().strip()
            section["api_key_env"] = env_entry.get().strip()
            selected_id["value"] = new_id
            if draft.get("active_provider") == "" or draft.get("active_provider") not in provider_ids():
                draft["active_provider"] = new_id
            refresh_active_boxes()

        providers_list.bind("<<ListboxSelect>>", on_provider_select)

        # --- Кнопки провайдеров ---
        def add_provider() -> None:
            commit_fields_to_draft()
            base = "provider"
            n = 1
            while f"{base}{n}" in provider_ids():
                n += 1
            new = {"id": f"{base}{n}", "name": f"Новый провайдер {n}", "base_url": "",
                   "api_key": "", "api_key_env": "", "models": []}
            draft["providers"].append(new)
            selected_id["value"] = new["id"]
            refresh_providers_list()
            load_selected_into_fields()

        def remove_provider() -> None:
            section = get_provider(selected_id["value"], draft)
            if section is None:
                return
            if len(draft["providers"]) <= 1:
                messagebox.showwarning("Настройки", "Должен остаться хотя бы один провайдер.", parent=win)
                return
            if not messagebox.askyesno("Настройки", f"Удалить провайдера «{section.get('name')}»?", parent=win):
                return
            draft["providers"] = [p for p in draft["providers"] if p["id"] != section["id"]]
            if draft.get("active_provider") == section["id"]:
                draft["active_provider"] = provider_ids()[0]
            selected_id["value"] = provider_ids()[0]
            refresh_providers_list()
            load_selected_into_fields()

        def make_active() -> None:
            commit_fields_to_draft()
            draft["active_provider"] = selected_id["value"]
            section = get_provider(selected_id["value"], draft)
            models = (section or {}).get("models") or []
            draft["active_model"] = models[0] if models else ""
            refresh_active_boxes()

        def _mini_btn(parent_, text, cmd, variant="primary"):
            return _button(parent_, text, cmd, variant=variant, font=Font.CAPTION,
                           padx=Space.SM, pady=Space.XS - 1)

        _mini_btn(prov_btns, "Добавить", add_provider).pack(side="left", padx=Space.XS // 2, pady=2)
        _mini_btn(prov_btns, "Удалить", remove_provider, "ghost").pack(side="left", padx=Space.XS // 2, pady=2)
        _mini_btn(prov_btns, "Сделать активным", make_active, "secondary").pack(side="left", padx=Space.XS // 2, pady=2)

        # --- Кнопки моделей ---
        def add_model() -> None:
            section = get_provider(selected_id["value"], draft)
            if section is None:
                return
            value = simpledialog.askstring("Модель", "Имя модели:", parent=win)
            if not value:
                return
            value = value.strip()
            if value and value not in section["models"]:
                section["models"].append(value)
            refresh_models_list()
            refresh_active_boxes()

        def remove_model() -> None:
            section = get_provider(selected_id["value"], draft)
            sel = models_list.curselection()
            if section is None or not sel:
                return
            section["models"].pop(sel[0])
            refresh_models_list()
            refresh_active_boxes()

        _mini_btn(models_btns, "+", add_model).pack(pady=2)
        _mini_btn(models_btns, "−", remove_model, "ghost").pack(pady=2)

        # --- Смена активного провайдера/модели сверху ---
        def active_provider_changed(_event=None) -> None:
            name = active_provider_var.get()
            section = next((p for p in draft["providers"] if p.get("name", p["id"]) == name), None)
            if section:
                draft["active_provider"] = section["id"]
                draft["active_model"] = (section.get("models") or [""])[0]
                active_model_box.configure(values=section.get("models") or [])
                active_model_var.set(draft["active_model"])
            refresh_balance()

        active_provider_box.bind("<<ComboboxSelected>>", active_provider_changed)

        def active_model_changed(_event=None) -> None:
            draft["active_model"] = active_model_var.get().strip()
            update_vision_warning()

        active_model_box.bind("<<ComboboxSelected>>", active_model_changed)
        ocr_mode_box.bind("<<ComboboxSelected>>", update_vision_warning)

        # --- Сбор общих настроек в draft ---
        def collect_common() -> None:
            draft["ocr_mode"] = reverse_ocr.get(ocr_mode_var.get(), "tesseract")
            draft["ocr_lang"] = ocr_entry.get().strip() or "eng"
            draft["target_lang"] = target_entry.get().strip() or "русский"
            draft["keep_image_on_top"] = bool(on_top_var.get())
            draft["live_on_top"] = bool(live_on_top_var.get())
            draft["skip_no_text"] = bool(skip_no_text_var.get())
            draft["stream_results"] = bool(stream_var.get())
            draft["capture_mode"] = capture_var.get() or capture_mode_of(draft)
            draft["active_model"] = active_model_var.get().strip() or draft.get("active_model", "")

        # --- Тест соединения на текущих полях (на лету) ---
        def test_connection() -> None:
            commit_fields_to_draft()
            collect_common()
            status.configure(text="Проверка соединения…", fg=SECONDARY_COLOR)

            def worker() -> None:
                try:
                    result = _test_connection_for(draft)
                    msg = f"OK. Ответ: {result[:60]}" if result else "OK, но пустой ответ."
                    color = "#7ED957"
                except Exception as exc:  # noqa: BLE001
                    msg = f"Ошибка: {exc}"
                    color = PRIMARY_COLOR_2
                win.after(0, lambda: status.configure(text=msg, fg=color))

            threading.Thread(target=worker, daemon=True).start()

        # --- Открыть config.json ---
        def open_config() -> None:
            apply_config(show_message=False)
            try:
                os.startfile(CONFIG_PATH)  # type: ignore[attr-defined]
            except Exception as exc:  # noqa: BLE001
                logger.warning("Не удалось открыть config.json: %s", exc)
                messagebox.showinfo("Настройки", f"Файл настроек:\n{CONFIG_PATH}", parent=win)

        # --- Перезагрузить конфиг с диска (на лету) ---
        def reload_from_disk() -> None:
            fresh = load_config()
            draft.clear()
            draft.update(json.loads(json.dumps(fresh)))
            selected_id["value"] = draft.get("active_provider") or provider_ids()[0]
            ocr_mode_var.set(OCR_MODE_LABELS.get(draft.get("ocr_mode", "tesseract"), OCR_MODE_LABELS["tesseract"]))
            ocr_entry.delete(0, tk.END)
            ocr_entry.insert(0, draft.get("ocr_lang", "eng"))
            target_entry.delete(0, tk.END)
            target_entry.insert(0, draft.get("target_lang", "русский"))
            autostart_var.set(is_autostart_enabled())
            on_top_var.set(bool(draft.get("keep_image_on_top")))
            live_on_top_var.set(bool(draft.get("live_on_top", True)))
            skip_no_text_var.set(bool(draft.get("skip_no_text", True)))
            stream_var.set(bool(draft.get("stream_results", True)))
            capture_var.set(capture_mode_of(draft))
            update_capture_hint()
            refresh_providers_list()
            load_selected_into_fields()
            _apply_runtime(draft)
            status.configure(text="Конфиг перезагружен с диска и применён.", fg="#7ED957")

        # --- Применение: обновить глобальный CONFIG и переподключиться ---
        def apply_config(show_message: bool = True) -> None:
            commit_fields_to_draft()
            collect_common()
            draft["add_to_autostart"] = bool(autostart_var.get())
            save_config(draft)
            _apply_runtime(draft)
            logger.info("Настройки применены на лету: провайдер=%s модель=%s",
                        CONFIG["active_provider"], CONFIG["active_model"])
            if show_message:
                messagebox.showinfo("Настройки", "Сохранено и применено без перезапуска.", parent=win)

        btns = tk.Frame(win, bg=Color.SURFACE)
        btns.pack(fill="x", **pad)
        _mini_btn(btns, "Проверить", test_connection, "secondary").pack(side="left", padx=Space.XS - 1)
        _mini_btn(btns, "Перезагрузить из файла", reload_from_disk, "ghost").pack(side="left", padx=Space.XS - 1)
        _mini_btn(btns, "Открыть config.json", open_config, "ghost").pack(side="left", padx=Space.XS - 1)
        _mini_btn(btns, "Отмена", on_close, "ghost").pack(side="right", padx=Space.XS - 1)
        _mini_btn(btns, "Сохранить", apply_config).pack(side="right", padx=Space.XS - 1)

        # Первичная отрисовка
        refresh_providers_list()
        load_selected_into_fields()
        update_capture_hint()
        refresh_balance()
        _center_window(win)

    if parent is None:
        run_on_ui(build)
    else:
        build()  # parent живёт в UI-потоке — вызывающий уже там


def _apply_runtime(config: dict) -> None:
    """Применяет конфиг к работающему приложению (переподключение на лету).

    Глобальный CONFIG заменяется новым содержимым, поэтому следующий же
    скриншот пойдёт через новый провайдер/модель без перезапуска.
    """
    CONFIG.clear()
    CONFIG.update(json.loads(json.dumps(config)))
    refresh_provider_labels()
    # Синхронизируем автозапуск: включаем/выключаем и чиним устаревший путь.
    want = bool(CONFIG.get("add_to_autostart"))
    if want != is_autostart_enabled() or (want and autostart_is_stale()):
        set_autostart(want)
    _apply_topmost_to_result()


def _test_connection_for(config: dict) -> str:
    """Пробный перевод с использованием временного конфига (не трогает глобальный)."""
    provider_id = config.get("active_provider", "")
    section = get_provider(provider_id, config)
    if section is None:
        raise RuntimeError("Провайдер не выбран")
    base_url = (section.get("base_url") or "").rstrip("/")
    if not base_url:
        raise RuntimeError("Не задан Base URL")
    url = base_url if base_url.endswith("/chat/completions") else base_url + "/chat/completions"
    model = (config.get("active_model") or "").strip()
    if not model:
        raise RuntimeError("Не выбрана модель")
    key = (section.get("api_key") or "").strip()
    if not key and section.get("api_key_env"):
        key = env_var(section["api_key_env"])
    if not key and provider_id == "deepseek":
        key = env_var("DEEPSEEK_API_KEY") or DEFAULT_DEEPSEEK_API_KEY

    headers = {"Content-Type": "application/json"}
    if key:
        headers["Authorization"] = f"Bearer {key}"
    payload = {"model": model, "messages": _build_messages("ping"),
               "temperature": 0.0, "stream": False}
    timeout = float(config.get("request_timeout") or 120)
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=timeout)
    except requests.RequestException as exc:
        raise RuntimeError(f"Нет соединения: {exc}") from exc
    if response.status_code == 401:
        if not key or key == DEFAULT_DEEPSEEK_API_KEY:
            env_name = (section.get("api_key_env") or "DEEPSEEK_API_KEY").strip()
            raise RuntimeError(f"API-ключ не найден. Задайте переменную окружения "
                               f"{env_name} или впишите ключ в настройках.")
        raise RuntimeError("401 — неверный API-ключ")
    if response.status_code == 404:
        raise RuntimeError(f"404 — модель '{model}' не найдена")
    response.raise_for_status()
    data = response.json()
    return data["choices"][0]["message"]["content"].strip()


# ---------------------------------------------------------------------------
# Горячая клавиша и логика обработки скриншота
# ---------------------------------------------------------------------------
def _linux_capture_screenshot() -> bool:
    """Запускает инструмент выделения области и кладёт результат в буфер обмена.

    Пробует по очереди grim (Wayland), gnome-screenshot, spectacle, scrot, maim,
    flameshot — что найдётся в системе.
    """
    candidates: list[list[str]] = []
    if os.environ.get("WAYLAND_DISPLAY") and shutil.which("grim"):
        # grim без -g делает полный снимок; область — через slurp, если есть.
        if shutil.which("slurp"):
            candidates.append(["bash", "-c", "grim -g \"$(slurp)\" - | wl-copy"])
        candidates.append(["bash", "-c", "grim - | wl-copy"])
    if shutil.which("gnome-screenshot"):
        candidates.append(["gnome-screenshot", "-a", "-c"])
    if shutil.which("spectacle"):
        candidates.append(["spectacle", "-r", "-n", "-b", "-c"])
    if shutil.which("flameshot"):
        candidates.append(["flameshot", "gui", "--clipboard"])
    if shutil.which("scrot"):
        candidates.append(["bash", "-c", "scrot -s /tmp/pait_sel.png && "
                           "(xclip -selection clipboard -t image/png -i /tmp/pait_sel.png || true)"])
    if shutil.which("maim"):
        candidates.append(["bash", "-c", "maim -s /tmp/pait_sel.png && "
                           "(xclip -selection clipboard -t image/png -i /tmp/pait_sel.png || true)"])

    if not candidates:
        return False

    for cmd in candidates:
        try:
            logger.info("Захват области через: %s", cmd[0])
            subprocess.run(cmd, timeout=120)
            return True
        except Exception as exc:  # noqa: BLE001
            logger.warning("Не удалось запустить %s: %s", cmd[0], exc)
            continue
    return False


def _windows_process_names() -> set[str]:
    """Имена запущенных процессов (Windows). Пустое множество — если не вышло."""
    names: set[str] = set()
    if not IS_WINDOWS:
        return names
    try:
        import ctypes
        from ctypes import wintypes

        query_limited = 0x1000                      # PROCESS_QUERY_LIMITED_INFORMATION
        psapi = ctypes.WinDLL("psapi")
        kernel32 = ctypes.WinDLL("kernel32")
        count = 1024
        while True:
            array = (wintypes.DWORD * count)()
            needed = wintypes.DWORD()
            if not psapi.EnumProcesses(array, ctypes.sizeof(array), ctypes.byref(needed)):
                return names
            used = needed.value // ctypes.sizeof(wintypes.DWORD)
            if used < count:
                break
            count *= 2
            if count > 65536:
                return names
        for pid in list(array)[:used]:
            if not pid:
                continue
            handle = kernel32.OpenProcess(query_limited, False, pid)
            if not handle:
                continue
            try:
                buffer = ctypes.create_unicode_buffer(1024)
                size = wintypes.DWORD(len(buffer))
                if kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
                    names.add(os.path.basename(buffer.value).lower())
            finally:
                kernel32.CloseHandle(handle)
    except Exception as exc:  # noqa: BLE001
        logger.debug("Не удалось перечислить процессы: %s", exc)
    return names


# Процессы штатного средства скриншотов Windows. Если оверлей выделения был
# открыт и закрылся, а новой картинки в буфере так и не появилось — значит,
# выделение отменили, и ждать впустую не нужно.
_WINDOWS_SNIP_HELPERS = {"snippingtool.exe", "screensketch.exe", "snipandsketch.exe",
                         "screenclip.exe"}


# ---------------------------------------------------------------------------
# Своё выделение области: полноэкранный оверлей, как Win+Shift+S.
#
# Экран «замораживается» одним снимком, поверх него показывается затемнение и
# рамка выделения. Это убирает зависимость от буфера обмена и системных
# «Ножниц»: выделение видно сразу, отмена мгновенная, а картинка берётся из
# уже снятого кадра — без гонок с чужими программами.
# ---------------------------------------------------------------------------
# Кэш геометрии рабочего стола: перечисление мониторов не должно задерживать захват.
_VIRTUAL_BOUNDS: dict = {}


def _virtual_screen_bounds() -> tuple[int, int, int, int]:
    """Объединение всех мониторов: (x, y, ширина, высота) в пикселях.

    Геометрия меняется редко, поэтому результат кэшируется: перечисление
    мониторов не должно задерживать появление оверлея при каждом захвате.
    """
    cached = _VIRTUAL_BOUNDS.get("value")
    if cached is not None:
        return cached
    bounds = _query_virtual_screen_bounds()
    _VIRTUAL_BOUNDS["value"] = bounds
    return bounds


def _query_virtual_screen_bounds() -> tuple[int, int, int, int]:
    if IS_WINDOWS:
        try:
            import ctypes
            from ctypes import wintypes

            class RECT(ctypes.Structure):
                _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                            ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

            rects: list[tuple[int, int, int, int]] = []

            def callback(_monitor, _hdc, rect_pointer, _data):
                rect = ctypes.cast(rect_pointer, ctypes.POINTER(RECT)).contents
                rects.append((rect.left, rect.top, rect.right, rect.bottom))
                return True

            monitor_proc = ctypes.WINFUNCTYPE(
                ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p,
                ctypes.POINTER(RECT), ctypes.c_double)
            ctypes.windll.user32.EnumDisplayMonitors(None, None, monitor_proc(callback), 0)
            if rects:
                left = min(r[0] for r in rects)
                top = min(r[1] for r in rects)
                right = max(r[2] for r in rects)
                bottom = max(r[3] for r in rects)
                if right > left and bottom > top:
                    return left, top, right - left, bottom - top
        except Exception as exc:  # noqa: BLE001
            logger.debug("Не удалось перечислить мониторы: %s", exc)
    if _ui_root is not None:
        return 0, 0, _ui_root.winfo_screenwidth(), _ui_root.winfo_screenheight()
    return 0, 0, 1920, 1080


def grab_virtual_screen() -> "Image.Image | None":
    """Снимок всего рабочего стола (все мониторы) или None, если не вышло."""
    try:
        image = ImageGrab.grab(all_screens=True)
        if image and isinstance(image, Image.Image) and image.width > 0:
            return image.convert("RGB")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Не удалось снять экран целиком: %s", exc)
    return None


class RegionOverlay:
    """Полноэкранное окно выделения области поверх замороженного снимка.

    Отрисовка сделана максимально дешёвой: фон — один раз затемнённая копия
    экрана, а при движении мыши двигаются только рамка, подпись размера и
    (для небольших выделений) яркий фрагмент внутри рамки. Перерисовка
    «полуэкранных» прямоугольников на каждое движение мыши заметно тормозила
    выделение, поэтому её здесь нет.
    """

    def __init__(self, shot: Image.Image, bounds: tuple[int, int, int, int]) -> None:
        self.shot = shot
        self.bounds = bounds
        self.result: "Image.Image | None" = None
        self.cancelled = False
        self.done = threading.Event()
        self.win: "tk.Toplevel | None" = None
        self._canvas: "tk.Canvas | None" = None
        self._photo: "ImageTk.PhotoImage | None" = None
        self._crop_photo: "ImageTk.PhotoImage | None" = None
        self._crop_item: "int | None" = None
        self._rect_item: "int | None" = None
        self._hint_item: "int | None" = None
        self._size_item: "int | None" = None
        self._start: tuple[int, int] | None = None
        self._current: tuple[int, int] = (0, 0)
        self._canvas_w = 0
        self._canvas_h = 0
        self._scale = 1.0
        self._last_crop = 0.0
        self._poll_job: "str | None" = None
        self._job: "Job | None" = None

    # --- построение и закрытие -------------------------------------------
    def build(self, job: "Job | None" = None) -> None:
        """Создаёт оверлей (вызывать в UI-потоке)."""
        self._job = job
        x, y, width, height = self.bounds
        win = tk.Toplevel(_ui_root)
        self.win = win
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.geometry(f"{width}x{height}+{x}+{y}")
        win.configure(bg="black", cursor="crosshair")
        win.bind("<Escape>", lambda _e: self.cancel("Escape"))
        win.bind("<Button-3>", lambda _e: self.cancel("правая кнопка"))
        win.bind("<Button-1>", self._on_press)
        win.bind("<B1-Motion>", self._on_motion)
        win.bind("<ButtonRelease-1>", self._on_release)

        canvas = tk.Canvas(win, width=width, height=height, highlightthickness=0, bd=0,
                           cursor="crosshair", bg="black")
        canvas.pack(fill="both", expand=True)
        self._canvas = canvas
        self._canvas_w = width
        self._canvas_h = height
        self._scale = self.shot.width / max(1, width)

        # Затемнение делаем прозрачностью самого окна: это мгновенно и не требует
        # полноэкранной картинки (11 Мпикс в Tk — это ~180 мс и лишние 40 МБ).
        # Выделенный фрагмент рисуется поверх: он совпадает с тем, что видно за
        # окном, поэтому выглядит ярким — как в системных «Ножницах».
        dim = self._dim_level()
        self._veil_by_alpha = False
        if dim > 0.01:
            try:
                win.attributes("-alpha", 1.0 - dim)
                self._veil_by_alpha = True
            except tk.TclError as exc:
                logger.debug("Прозрачность окна недоступна (%s) — затемняю картинку.", exc)
        if not self._veil_by_alpha and dim > 0.01:
            # Запасной путь для систем без прозрачности окна.
            self._photo = ImageTk.PhotoImage(self._dimmed())

        # Яркий фрагмент внутри рамки (для небольших выделений).
        self._crop_item = canvas.create_image(0, 0, anchor="nw")
        self._rect_item = canvas.create_rectangle(0, 0, 0, 0, outline=Color.PRIMARY, width=2)
        self._size_item = canvas.create_text(0, 0, text="", anchor="nw", fill=Color.TEXT_ON_DARK,
                                             font=Font.BODY_BOLD)
        # Подсказка: сначала плашка под ней, потом сам текст.
        canvas.create_rectangle(width // 2 - 260, Space.LG, width // 2 + 260, Space.XL + 30,
                                fill="#10131A", outline="")
        self._hint_item = canvas.create_text(
            width // 2, Space.XL + 15, text="Выделите область мышью · Esc или правая кнопка — отмена",
            fill=Color.TEXT_ON_DARK, font=Font.BODY_BOLD)

        try:
            win.focus_force()
        except tk.TclError:
            pass
        self._poll_job = win.after(120, self._poll_job_state)

    @staticmethod
    def _dim_level() -> float:
        """Насколько затемнять экран при выделении (0 — не затемнять)."""
        try:
            dim = float(CONFIG.get("capture_dim", 0.45) or 0.0)
        except (TypeError, ValueError):
            dim = 0.45
        return max(0.0, min(0.85, dim))

    def _dimmed(self) -> Image.Image:
        """Затемнённая копия снимка (яркость = 1 − capture_dim)."""
        try:
            dim = float(CONFIG.get("capture_dim", 0.45) or 0.0)
        except (TypeError, ValueError):
            dim = 0.45
        dim = max(0.0, min(0.85, dim))
        if dim <= 0.01:
            return self.shot
        return ImageEnhance.Brightness(self.shot).enhance(1.0 - dim)

    def _poll_job_state(self) -> None:
        """Закрывается, если задачу отменили снаружи (новый захват, трей)."""
        if self.win is None or not self.win.winfo_exists():
            return
        if self._job is not None and self._job.cancelled():
            self.cancel("задача отменена")
            return
        self._poll_job = self.win.after(120, self._poll_job_state)

    def close_window(self) -> None:
        if self._poll_job and self.win is not None:
            try:
                self.win.after_cancel(self._poll_job)
            except Exception:  # noqa: BLE001
                pass
        self._poll_job = None
        if self.win is not None:
            try:
                self.win.destroy()
            except tk.TclError:
                pass
        self.win = None
        self._canvas = None
        self._photo = None
        self._crop_photo = None
        self._crop_item = None
        self._rect_item = None
        self._hint_item = None
        self._size_item = None
        self.done.set()

    def cancel(self, reason: str = "отменено") -> None:
        if self.done.is_set():
            return
        logger.info("Выделение области отменено (%s).", reason)
        self.cancelled = True
        self.result = None
        self.close_window()

    # --- выделение --------------------------------------------------------
    def _to_image_point(self, x: int, y: int) -> tuple[int, int]:
        """Переводит координаты окна в координаты снимка (на случай масштаба)."""
        if self._scale == 1.0:
            return x, y
        return int(x * self._scale), int(y * self._scale)

    def begin(self, x: int, y: int) -> None:
        self._start = (x, y)
        self._current = (x, y)
        self._last_crop = 0.0

    def update(self, x: int, y: int) -> None:
        """Двигает рамку и подпись; тяжёлых перерисовок здесь нет."""
        if self._start is None or self._canvas is None:
            return
        self._current = (x, y)
        x0, y0 = self._start
        left, right = sorted((x0, x))
        top, bottom = sorted((y0, y))
        if self._rect_item is not None:
            self._canvas.coords(self._rect_item, left, top, right, bottom)
        if self._size_item is not None:
            label = f"{right - left} × {bottom - top}"
            # Подпись держим внутри экрана, чтобы не уезжала за край.
            label_y = top - 22 if top >= 28 else bottom + 6
            self._canvas.coords(self._size_item, min(left, self._canvas_w - 90), label_y)
            self._canvas.itemconfigure(self._size_item, text=label, state="normal")
        if self._hint_item is not None:
            self._canvas.itemconfigure(self._hint_item, state="hidden")
        self._show_crop(left, top, right, bottom)

    def _show_crop(self, left: int, top: int, right: int, bottom: int) -> None:
        """Показывает незатемнённый фрагмент внутри рамки.

        Создание картинки стоит времени, пропорционального области, поэтому
        обновляем не чаще ~10 раз в секунду и пропускаем очень большие рамки:
        при перетаскивании через весь экран важнее плавность.
        """
        if self._crop_item is None or self._canvas is None:
            return
        area = (right - left) * (bottom - top)
        if area < 64 or area > 600_000 or self._scale != 1.0:
            return
        now = time.time()
        if now - self._last_crop < 0.12:
            return
        self._last_crop = now
        try:
            crop = self.shot.crop((left, top, right, bottom))
            self._crop_photo = ImageTk.PhotoImage(crop)
            self._canvas.itemconfigure(self._crop_item, image=self._crop_photo)
            self._canvas.coords(self._crop_item, left, top)
        except Exception as exc:  # noqa: BLE001
            logger.debug("Не удалось показать выделенный фрагмент: %s", exc)

    def finish(self, x: int, y: int) -> None:
        if self._start is None:
            self.cancel("выделение не начато")
            return
        self.update(x, y)
        x0, y0 = self._start
        left, right = sorted((x0, x))
        top, bottom = sorted((y0, y))
        if right - left < 8 or bottom - top < 8:
            self.cancel("область слишком мала")
            return
        ix0, iy0 = self._to_image_point(left, top)
        ix1, iy1 = self._to_image_point(right, bottom)
        if ix1 - ix0 < 2 or iy1 - iy0 < 2:
            self.cancel("область слишком мала")
            return
        try:
            self.result = self.shot.crop((ix0, iy0, ix1, iy1))
        except Exception as exc:  # noqa: BLE001
            logger.error("Не удалось вырезать область: %s", exc)
            self.result = None
        logger.info("Выделена область %dx%d (экран %dx%d).",
                    ix1 - ix0, iy1 - iy0, self.shot.width, self.shot.height)
        self.close_window()

    # --- обработчики Tk ---------------------------------------------------
    def _on_press(self, event) -> None:
        self.begin(event.x, event.y)

    def _on_motion(self, event) -> None:
        self.update(event.x, event.y)

    def _on_release(self, event) -> None:
        self.finish(event.x, event.y)


def select_region_interactive(job: "Job | None" = None) -> tuple["Image.Image | None", str]:
    """Показывает оверлей выделения и ждёт результат (можно вызывать из потока).

    Возвращает (снимок, причина). Причина: "ok" — область выделена,
    "cancelled" — выделение отменили, "timeout" — не успели,
    "unsupported" — снять экран не удалось (нужен системный захват).
    """
    shot = grab_virtual_screen()
    if shot is None:
        return None, "unsupported"
    bounds = _virtual_screen_bounds()
    # Снимок и границы мониторов должны совпадать; если система отдала другую
    # геометрию — подстраиваем границы под картинку.
    if (bounds[2], bounds[3]) != shot.size and bounds[0] == 0 and bounds[1] == 0:
        logger.debug("Размер снимка %s не совпал с экраном %s — беру размер снимка.",
                     shot.size, bounds[2:])
        bounds = (bounds[0], bounds[1], shot.width, shot.height)

    overlay = RegionOverlay(shot, bounds)
    run_on_ui(lambda: overlay.build(job))
    timeout = max(10.0, float(CONFIG.get("capture_timeout") or 60))
    if not overlay.done.wait(timeout):
        logger.info("Выделение области не завершено за %.0f с — отменяю.", timeout)
        run_on_ui(lambda: overlay.cancel("истекло время"))
        overlay.done.wait(3.0)
        return None, "timeout"
    if overlay.result is None:
        return None, "cancelled"
    return overlay.result, "ok"


def on_hotkey() -> None:
    """Горячая клавиша: ждём новую картинку в буфере обмена и обрабатываем её.

    Повторное нажатие во время ожидания или обработки отменяет предыдущую
    задачу — можно спокойно переснять область. Обработку также можно прервать
    кнопкой «Отмена» в плашке прогресса, в окне перевода или из меню трея.
    """
    job = _new_job("hotkey")
    progress = _progress()
    try:
        image = _capture_region(job, progress)
        if image is None:
            return
        logger.info("Задача #%d: снимок получен (%dx%d).", job.gen, *image.size)
        process_image(image, job)
    except Cancelled:
        logger.info("Задача #%d отменена до начала перевода.", job.gen)
        _toast("Обработка отменена", kind="warn")
    except RuntimeError as exc:
        logger.error("Задача #%d: %s", job.gen, exc)
        _toast(str(exc), kind="error", seconds=6.0)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Непредвиденная ошибка в задаче #%d", job.gen)
        _toast(f"Произошла ошибка: {exc}", kind="error", seconds=6.0)
    finally:
        _finish_job(job)
        progress.close(owner=job.gen)


def _capture_region(job: Job, progress: "Progress") -> "Image.Image | None":
    """Получает снимок выделенной области: своим оверлеем или системным средством."""
    if capture_mode() == "builtin":
        job.check()
        logger.info("Задача #%d: своё выделение области.", job.gen)
        image, reason = select_region_interactive(job)
        if image is not None:
            return image
        job.check()
        if reason != "unsupported":
            return None                     # пользователь отменил выделение
        logger.info("Своё выделение недоступно — перехожу к системному захвату.")
        _toast("Своё выделение недоступно — открываю системное", kind="warn")
        return _capture_via_clipboard(job, progress, open_system_selector=True)
    return _capture_via_clipboard(job, progress)


def _capture_via_clipboard(job: Job, progress: "Progress", *,
                           open_system_selector: bool = False) -> "Image.Image | None":
    """Системный захват: ждём результат выделения в буфере обмена.

    На Windows оверлей открывает сама система (Win+Shift+S), на Linux —
    вызывается один из установленных инструментов выделения.
    open_system_selector — нажать системное сочетание за пользователя
    (нужно, если своё выделение не сработало).
    """
    # На Linux нет встроенного Win+Shift+S — вызываем выбранный инструмент.
    if IS_LINUX:
        progress.build("Захват области…", owner=job.gen)
        progress.update(0, "Выделите область", owner=job.gen)
        if not _linux_capture_screenshot():
            _toast("Не найден инструмент скриншотов. Установите один из: grim, "
                   "gnome-screenshot, spectacle, scrot, maim, flameshot.",
                   kind="error", seconds=7.0)
            return None
    else:
        progress.build("Ожидание выделения…", owner=job.gen)
        if open_system_selector and keyboard is not None:
            try:
                keyboard.send("windows+shift+s")
                logger.info("Открываю системное выделение области (Win+Shift+S).")
            except Exception as exc:  # noqa: BLE001
                logger.warning("Не удалось нажать Win+Shift+S: %s", exc)

    seq_before = _clipboard_sequence()
    logger.info("Задача #%d: жду результат выделения в буфере обмена…", job.gen)
    progress.update(5, "Выделите область скриншота", owner=job.gen)

    timeout = max(5.0, float(CONFIG.get("capture_timeout") or 60))
    start = time.time()
    helper_seen = False
    helper_gone_since: float | None = None
    last_process_check = 0.0
    while time.time() - start < timeout:
        job.check()
        # Новая картинка в буфере (или буфер изменился).
        if _clipboard_sequence() != seq_before:
            candidate = get_image_from_clipboard()
            if candidate and isinstance(candidate, Image.Image):
                logger.info("Задача #%d: снимок получен из буфера обмена (%dx%d).",
                            job.gen, *candidate.size)
                return candidate
        elapsed = time.time() - start
        progress.update(min(90.0, elapsed / timeout * 90.0), "Ожидание выделения…",
                        owner=job.gen)
        now = time.time()
        if IS_WINDOWS and now - last_process_check > 0.8:
            last_process_check = now
            if _windows_process_names() & _WINDOWS_SNIP_HELPERS:
                helper_seen, helper_gone_since = True, None
            elif helper_seen:
                # Оверлей закрылся: даём пару секунд на запись в буфер.
                helper_gone_since = helper_gone_since or now
                if now - helper_gone_since > 2.5:
                    logger.info("Задача #%d: окно захвата закрыто без снимка — выхожу.",
                                job.gen)
                    _toast("Захват отменён")
                    return None
        time.sleep(0.2)

    logger.info("Задача #%d: снимок не получен (выделение отменено или истёк таймаут).",
                job.gen)
    return None


# ---------------------------------------------------------------------------
# Всплывающие подсказки: не крадут фокус, не блокируют работу и сами исчезают
# ---------------------------------------------------------------------------
_TOAST_STATE: dict = {"win": None, "after": None}


def _close_toast() -> None:
    """Закрывает текущую подсказку (вызывается в UI-потоке)."""
    win = _TOAST_STATE.get("win")
    _TOAST_STATE["win"] = None
    job_id = _TOAST_STATE.get("after")
    _TOAST_STATE["after"] = None
    if job_id is not None and _ui_root is not None:
        try:
            _ui_root.after_cancel(job_id)
        except Exception:  # noqa: BLE001
            pass
    if win is not None:
        try:
            win.destroy()
        except tk.TclError:
            pass


def _toast(message: str, *, kind: str = "info", seconds: float = 2.4) -> None:
    """Показывает ненавязчивое сообщение внизу экрана (и дублирует его в лог)."""
    logger.info("Сообщение: %s", (message or "").replace("\n", " "))
    border = {"info": Color.PRIMARY, "warn": Color.ACCENT,
              "error": Color.DANGER, "success": Color.SUCCESS}.get(kind, Color.PRIMARY)

    def build() -> None:
        _close_toast()
        if _ui_root is None:
            return
        win = tk.Toplevel(_ui_root)
        _TOAST_STATE["win"] = win
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(bg=Color.SURFACE)

        frame = tk.Frame(win, bg=Color.SURFACE, highlightthickness=1,
                         highlightbackground=border)
        frame.pack(fill="both", expand=True)
        label = tk.Label(frame, text=message, bg=Color.SURFACE, fg=Color.TEXT_ON_DARK,
                         font=Font.BODY, justify="left", wraplength=420)
        label.pack(padx=Space.MD, pady=Space.SM)
        for widget in (win, frame, label):
            widget.bind("<Button-1>", lambda _e: _close_toast())

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        area_x, area_y, area_w, area_h = _cursor_monitor_rect(win)
        x = area_x + max(Space.LG, area_w - w - Space.XL)
        y = area_y + max(Space.LG, area_h - h - 72)
        win.geometry(f"+{x}+{y}")
        _TOAST_STATE["after"] = _ui_root.after(int(max(0.8, seconds) * 1000), _close_toast)

    run_on_ui(build)


# ---------------------------------------------------------------------------
# Системный трей
# ---------------------------------------------------------------------------
def create_image_for_tray() -> Image.Image:
    try:
        return Image.open(resource_path("translator.png")).convert("RGBA")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Не удалось загрузить translator.png: %s", exc)
        image = Image.new("RGB", (64, 64), color=(50, 100, 150))
        draw = ImageDraw.Draw(image)
        draw.rectangle((16, 16, 48, 48), fill=(200, 200, 0))
        return image


def on_exit(icon, _item) -> None:
    logger.info("Завершение работы.")
    try:
        icon.stop()
    finally:
        if keyboard is not None:
            keyboard.unhook_all()
        os._exit(0)


def on_settings(icon, _item) -> None:
    run_on_ui(lambda: open_settings_window(None))


def _refresh_menu(icon) -> None:
    """Перерисовывает меню трея (галочки отражают актуальное состояние)."""
    try:
        icon.update_menu()
    except Exception:  # noqa: BLE001
        try:
            icon.menu = _build_menu()
        except Exception as exc:  # noqa: BLE001
            logger.debug("Не удалось обновить меню трея: %s", exc)


def on_toggle_autostart(icon, _item) -> None:
    enabled = not is_autostart_enabled()
    if set_autostart(enabled):
        _toast("Автозапуск включён" if enabled else "Автозапуск выключен", kind="success")
    else:
        _toast("Не удалось изменить автозапуск (подробности в журнале)", kind="error")
    _refresh_menu(icon)


def on_toggle_live(icon, _item) -> None:
    enabled = not is_live_mode()
    set_live_mode(enabled)
    _toast("Режим «на лету»: " + ("ВКЛ" if enabled else "ВЫКЛ"), kind="info")
    _refresh_menu(icon)


def on_toggle_top(icon, _item) -> None:
    CONFIG["keep_image_on_top"] = not bool(CONFIG.get("keep_image_on_top"))
    save_config(CONFIG)
    _apply_topmost_to_result()
    _refresh_menu(icon)


def on_cancel_job(icon, _item) -> None:
    """Пункт меню «Отменить обработку»."""
    if cancel_current_job("меню трея"):
        _toast("Обработка отменяется…", kind="warn")
    else:
        _toast("Обрабатывать нечего")


def _apply_topmost_to_result() -> None:
    """Применяет настройку «поверх других окон» к уже открытому окну перевода."""
    def build() -> None:
        if _result_alive():
            _set_topmost(_RESULT["win"], bool(CONFIG.get("keep_image_on_top")))

    run_on_ui(build)


def _build_menu():
    from pystray import Menu, MenuItem

    return Menu(
        MenuItem(f"pyAutoImgTranslate {APP_VERSION}", on_settings),
        MenuItem(f"Перевод области — {effective_hotkey()}", on_settings),
        Menu.SEPARATOR,
        MenuItem("Режим «на лету»", on_toggle_live, checked=lambda _i: is_live_mode()),
        MenuItem("Отменить обработку", on_cancel_job,
                 enabled=lambda _i: _current_job() is not None),
        Menu.SEPARATOR,
        MenuItem("Автозапуск при входе", on_toggle_autostart,
                 checked=lambda _i: is_autostart_enabled()),
        MenuItem("Окно перевода поверх других", on_toggle_top,
                 checked=lambda _i: bool(CONFIG.get("keep_image_on_top"))),
        Menu.SEPARATOR,
        MenuItem("Настройки…", on_settings),
        Menu.SEPARATOR,
        MenuItem("Выход", on_exit),
    )


def start_tray_icon() -> None:
    from pystray import Icon

    # title/name должны быть ASCII-frendly: X11 кодирует имя окна в latin-1.
    icon = Icon("pyAutoImgTranslate", create_image_for_tray(),
                f"pyAutoImgTranslate {APP_VERSION}", _build_menu())
    icon.run()


# ---------------------------------------------------------------------------
# Глобальные горячие клавиши
#   Windows — модуль keyboard
#   Linux   — pynput (X11); при отсутствии — работа в режиме «на лету»/трея
# ---------------------------------------------------------------------------
def _parse_hotkey(hotkey: str) -> set[str]:
    """'ctrl+alt+s' -> {'ctrl','alt','s'} с нормализацией имён."""
    parts = {p.strip().lower() for p in hotkey.split("+") if p.strip()}
    alias = {"win": "cmd", "super": "cmd", "windows": "cmd",
             "control": "ctrl", "escape": "esc"}
    return {alias.get(p, p) for p in parts}


class _LinuxHotkeys:
    """Слушатель глобальных горячих клавиш через pynput (X11)."""

    def __init__(self) -> None:
        self._listener = None
        self._bindings: list[tuple[frozenset[str], object]] = []

    def add(self, hotkey: str, callback) -> bool:
        try:
            from pynput import keyboard as pk  # type: ignore
        except Exception as exc:  # noqa: BLE001
            logger.warning("pynput недоступен — горячие клавиши на Linux отключены: %s", exc)
            return False
        self._bindings.append((frozenset(_parse_hotkey(hotkey)), callback))
        if self._listener is None:
            self._pk = pk
            self._pressed: set[str] = set()
            self._listener = pk.Listener(on_press=self._on_press, on_release=self._on_release)
            self._listener.daemon = True
            self._listener.start()
        return True

    def _norm(self, key) -> str | None:
        pk = self._pk
        if isinstance(key, pk.Key):
            name = key.name.lower()
            return {"ctrl_l": "ctrl", "ctrl_r": "ctrl", "alt_l": "alt", "alt_r": "alt",
                    "shift_l": "shift", "shift_r": "shift", "cmd_l": "cmd", "cmd_r": "cmd"}.get(name, name)
        if isinstance(key, pk.KeyCode) and key.char:
            return key.char.lower()
        return None

    def _on_press(self, key) -> None:
        name = self._norm(key)
        if not name:
            return
        self._pressed.add(name)
        for combo, callback in self._bindings:
            if combo.issubset(self._pressed):
                threading.Thread(target=callback, daemon=True).start()

    def _on_release(self, key) -> None:
        name = self._norm(key)
        if name:
            self._pressed.discard(name)

    def wait(self) -> None:
        while True:
            time.sleep(3600)


# ---------------------------------------------------------------------------
# Точка входа
# ---------------------------------------------------------------------------
# Уникальное имя/id для SingleInstance
_MUTEX_NAME = "Global\\pyAutoImgTranslate_SingleInstance"
_LOCK_FILE = Path.home() / ".cache" / "pyAutoImgTranslate.lock"


def _already_running() -> bool:
    """Возвращает True, если копия приложения уже запущена."""
    if IS_WINDOWS:
        try:
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.CreateMutexW(None, False, _MUTEX_NAME)
            if not handle:
                return False
            if kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
                return True
            setattr(_already_running, "_handle", handle)  # удерживаем дескриптор
            return False
        except Exception:  # noqa: BLE001
            return False
    # Linux/macOS — файловая блокировка через flock.
    try:
        import fcntl  # type: ignore

        _LOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
        fh = open(_LOCK_FILE, "w")
        try:
            fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            return True
        setattr(_already_running, "_fh", fh)  # удерживаем дескриптор
        return False
    except Exception:  # noqa: BLE001
        return False


def main() -> None:
    setup_logging()
    _enable_dpi_awareness()
    logger.info("Запуск pyAutoImgTranslate %s. ОС=%s, провайдер=%s, модель=%s",
                APP_VERSION, sys.platform, CONFIG.get("active_provider"),
                CONFIG.get("active_model"))
    provider_id = str(CONFIG.get("active_provider") or "")
    logger.info("Ключ провайдера %s: %s.", provider_id or "—", key_source_hint(provider_id))

    if _already_running():
        logger.warning("Приложение уже запущено — второй экземпляр не стартует.")
        sys.exit(0)

    ensure_autostart_consistency()

    tray_thread = threading.Thread(target=start_tray_icon, daemon=True)
    tray_thread.start()

    hotkey = effective_hotkey()
    live_hotkey = CONFIG.get("live_hotkey", _default_live_hotkey())
    logger.info("Режим выделения области: %s, клавиша перевода: %s.",
                capture_mode(), hotkey)
    if hotkey != (CONFIG.get("hotkey") or _default_hotkey()):
        logger.info("Клавиша из настроек (%s) занята системным скриншотом — "
                    "использую %s.", CONFIG.get("hotkey"), hotkey)

    if IS_WINDOWS:
        try:
            keyboard.add_hotkey(hotkey, lambda: threading.Thread(target=on_hotkey, daemon=True).start())
            logger.info("Горячая клавиша перевода %s зарегистрирована.", hotkey)
            try:
                keyboard.add_hotkey(live_hotkey, _toggle_live_from_hotkey)
                logger.info("Горячая клавиша режима «на лету» %s зарегистрирована.", live_hotkey)
            except Exception as exc:  # noqa: BLE001
                logger.warning("Не удалось зарегистрировать %s: %s", live_hotkey, exc)
            logger.info("Приложение в трее.")
            keyboard.wait()
        except Exception as exc:  # noqa: BLE001
            logger.error("Не удалось зарегистрировать горячую клавишу '%s': %s", hotkey, exc)
            raise SystemExit(1)
    else:
        # Linux: pynput, а при его отсутствии — живём без хоткеев (трей + live).
        hotkeys = _LinuxHotkeys()
        ok = hotkeys.add(hotkey, lambda: threading.Thread(target=on_hotkey, daemon=True).start())
        if ok:
            logger.info("Горячая клавиша перевода %s зарегистрирована (pynput).", hotkey)
            hotkeys.add(live_hotkey, _toggle_live_from_hotkey)
        logger.info("Приложение в трее.")
        hotkeys.wait()


def _toggle_live_from_hotkey() -> None:
    enabled = not is_live_mode()
    set_live_mode(enabled)
    logger.info("Режим «на лету» %s (горячая клавиша).", "включён" if enabled else "выключен")
    _toast("Режим «на лету»: " + ("ВКЛ" if enabled else "ВЫКЛ"),
           kind="success" if enabled else "info")


if __name__ == "__main__":
    main()
