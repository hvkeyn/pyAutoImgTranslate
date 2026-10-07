"""
pyAutoImgTranslate — перевод текста с экрана.

Возможности:
  * Win+Shift+S — выделение области экрана, OCR (Tesseract), перевод ИИ.
  * Поддержка нескольких провайдеров перевода: локальный LM Studio (OpenAI-совместимый
    сервер) и DeepSeek (облачный API). Можно добавлять свои модели.
  * Автозапуск вместе с Windows, работа в системном трее, логирование в файл.
  * Ненавязчивое окно результата с картинкой и аккуратным блоком перевода.
"""

from __future__ import annotations

import ctypes
import base64
import io
import json
import logging
import logging.handlers
import os
import queue
import re
import struct
import sys
import threading
import time
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

import cv2
import keyboard
import numpy as np
import pytesseract
import requests
import win32clipboard
import win32con

from PIL import Image, ImageDraw, ImageGrab, ImageTk

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
DEFAULT_DEEPSEEK_API_KEY = "sk-ЗАМЕНИТЕ_НА_СВОЙ_КЛЮЧ"

# Схема конфига полностью управляется данными: провайдеры лежат списком в
# CONFIG["providers"], у каждого — свой base_url, ключ и список моделей.
DEFAULT_CONFIG = {
    "active_provider": "deepseek",     # id активного провайдера
    "active_model": "deepseek-flash",    # выбранная модель
    "ocr_lang": "eng+rus",
    "ocr_mode": "ai",    # "ai" (нейросеть читает картинку — точнее) | "tesseract" (локальный OCR)
    "tesseract_cmd": r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    "source_lang": "авто",
    "target_lang": "русский",
    "request_timeout": 120,
    "add_to_autostart": False,
    "keep_image_on_top": False,
    "hotkey": "windows+shift+s",
    "live_hotkey": "windows+shift+a",
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


def get_api_key(provider_id: str) -> str:
    section = get_provider(provider_id) or {}
    key = (section.get("api_key") or "").strip()
    if key:
        return key
    env_name = (section.get("api_key_env") or "").strip()
    if env_name:
        env = os.environ.get(env_name, "").strip()
        if env:
            return env
    if provider_id == "deepseek":
        env = os.environ.get("DEEPSEEK_API_KEY", "").strip()
        if env:
            return env
        return DEFAULT_DEEPSEEK_API_KEY
    return ""


# ---------------------------------------------------------------------------
# Автозапуск вместе с Windows (реестр HKCU, не требует прав администратора)
# ---------------------------------------------------------------------------
AUTOSTART_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_REG_NAME = "pyAutoImgTranslate"


def _autostart_command() -> str:
    if getattr(sys, "frozen", False):
        return f'"{Path(sys.executable).resolve()}"'
    script = Path(__file__).resolve()
    python = Path(sys.executable).resolve()
    pythonw = python.with_name("pythonw.exe")
    launcher = pythonw if pythonw.is_file() else python
    return f'"{launcher}" "{script}"'


def is_autostart_enabled() -> bool:
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, AUTOSTART_REG_NAME)
            return bool(value)
    except FileNotFoundError:
        return False
    except Exception as exc:  # noqa: BLE001
        logger.warning("Не удалось проверить автозапуск: %s", exc)
        return False


def set_autostart(enabled: bool) -> bool:
    try:
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
        logger.info("Автозапуск %s.", "включён" if enabled else "выключен")
        return True
    except Exception as exc:  # noqa: BLE001
        logger.error("Не удалось изменить автозапуск: %s", exc)
        return False


def ensure_autostart_consistency() -> None:
    """Синхронизирует реестр с настройкой в config.json при старте."""
    want = bool(CONFIG.get("add_to_autostart"))
    if want and not is_autostart_enabled():
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


def ocr_image(image: Image.Image) -> str:
    """Выпрямляет, увеличивает изображение и извлекает текст через Tesseract."""
    _configure_tesseract()
    deskewed = deskew_image(image)
    upscaled = upscale_image(deskewed, scale=2)
    lang = CONFIG.get("ocr_lang", "eng+rus")
    config = r"--psm 6 --oem 3"
    try:
        text = pytesseract.image_to_string(upscaled, lang=lang, config=config)
    except pytesseract.TesseractError as exc:
        logger.warning("OCR с языком '%s' не удался (%s), пробую 'eng'.", lang, exc)
        text = pytesseract.image_to_string(upscaled, lang="eng", config=config)
    return _clean_ocr_text(text)


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
# Работа с буфером обмена
# ---------------------------------------------------------------------------
def get_image_from_clipboard() -> Image.Image | None:
    """Получает изображение из буфера обмена (через ImageGrab или CF_DIB)."""
    image = ImageGrab.grabclipboard()
    if image and isinstance(image, Image.Image):
        return image

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
        image = Image.open(io.BytesIO(header + data))
        return image
    except Exception as exc:  # noqa: BLE001
        logger.debug("Не удалось собрать изображение из CF_DIB: %s", exc)
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
    """Номер изменения буфера обмена — растёт при каждом копировании."""
    try:
        return int(ctypes.windll.user32.GetClipboardSequenceNumber())
    except Exception:  # noqa: BLE001
        return 0


# ---------------------------------------------------------------------------
# Режим «на лету»: как только в буфере появится новая картинка — сразу OCR+перевод.
# Удобно использовать с Win+Shift+S или любым другим средством скриншотов.
# ---------------------------------------------------------------------------
LIVE_STATE = {"enabled": False, "thread": None, "last_seq": 0}
_live_lock = threading.Lock()


def _live_watcher() -> None:
    """Фоновый поток: следит за буфером обмена и обрабатывает новые картинки."""
    logger.info("Режим «на лету» включён — жду картинку в буфере обмена.")
    # Текущее состояние считываем один раз, чтобы не реагировать на уже лежащую картинку.
    LIVE_STATE["last_seq"] = _clipboard_sequence()
    while LIVE_STATE["enabled"]:
        time.sleep(0.3)
        if not LIVE_STATE["enabled"]:
            break
        seq = _clipboard_sequence()
        if seq == LIVE_STATE["last_seq"]:
            continue
        LIVE_STATE["last_seq"] = seq
        image = get_image_from_clipboard()
        if not image or not isinstance(image, Image.Image):
            continue
        try:
            process_image(image)
        except Exception as exc:  # noqa: BLE001
            logger.exception("Ошибка в режиме «на лету»: %s", exc)
            _notify(f"Ошибка: {exc}")
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


def is_live_mode() -> bool:
    return bool(LIVE_STATE["enabled"])


PROGRESS: "Progress | None" = None


def _progress() -> "Progress":
    global PROGRESS
    if PROGRESS is None:
        PROGRESS = Progress()
    return PROGRESS


def process_image(image: Image.Image) -> None:
    """Общий пайплайн: распознавание+перевод → показ результата. Используется в hotkey и live."""
    mode = CONFIG.get("ocr_mode", "tesseract")
    logger.info("Скриншот получен (%dx%d). Режим: %s", *image.size, mode)
    progress = _progress()
    progress.build("Распознавание…")
    progress.update(10, "Подготовка изображения…")

    try:
        if mode == "ai":
            # Нейросеть сама читает картинку и сразу переводит — без Tesseract.
            progress.update(20, "Читаю и перевожу (нейросеть)…")
            translation, original = translate_image_via_ai(image)
            progress.finish("Готово")
            if not translation.strip():
                _notify("Не удалось распознать текст на изображении.")
                return
            logger.info("Перевод (vision) получен (%d симв.).", len(translation))
            show_result(image, translation, source_text=original)
            return

        # Режим Tesseract: сначала OCR, затем перевод нейросетью.
        progress.update(25, "Распознаю текст (Tesseract)…")
        text = ocr_image(image)
        if not text.strip():
            _notify("На изображении не найден текст.")
            return
        logger.info("Распознанный текст:\n%s", text)
        progress.update(60, "Перевожу (нейросеть)…")
        translation = translate_text(text)
        progress.finish("Готово")
        logger.info("Перевод получен (%d симв.).", len(translation))
        show_result(image, translation, source_text=text)
    except RuntimeError as exc:
        logger.error("Ошибка обработки: %s", exc)
        _notify(str(exc))
    finally:
        progress.close()


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


def _chat_completion(messages: list[dict], *, label: str) -> str:
    """Общий вызов чат-модели (перевод — через нейросеть). Возвращает текст ответа."""
    url, model, api_key = _provider_settings()
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0.0,
        "stream": False,
    }

    logger.info("%s: провайдер=%s модель=%s", label, CONFIG.get("active_provider"), model)
    timeout = float(CONFIG.get("request_timeout") or 120)
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=timeout)
    except requests.RequestException as exc:
        logger.error("Сеть недоступна: %s", exc)
        raise RuntimeError(f"Не удалось соединиться с сервером перевода:\n{exc}") from exc

    if response.status_code == 401:
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

    try:
        data = response.json()
        content = data["choices"][0]["message"]["content"]
    except (ValueError, KeyError, IndexError, TypeError) as exc:
        logger.error("Неожиданный ответ сервера: %s", response.text[:500])
        raise RuntimeError("Сервер вернул ответ в неожиданном формате.") from exc

    return _strip_model_noise(content)


def translate_text(text: str) -> str:
    """Переводит распознанный текст через нейросеть. Возвращает строку перевода."""
    text = (text or "").strip()
    if not text:
        return ""
    return _chat_completion(_build_messages(text), label=f"Перевод текста (длина={len(text)})")


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


def _build_vision_messages(image: Image.Image) -> list[dict]:
    target = CONFIG.get("target_lang", "русский") or "русский"
    source = CONFIG.get("source_lang", "авто") or "авто"
    source_hint = "" if source.lower() in ("авто", "auto", "") else f" Исходный язык — {source}."
    system = (
        f"Ты профессиональный переводчик.{source_hint} На изображении есть текст. "
        f"Считай его и переведи на {target} язык как можно ближе к оригиналу — точно и дословно, "
        "без отсебятины и пересказа.\n"
        "Ответь строго в таком формате (два блока):\n"
        "ОРИГИНАЛ:\n<текст как на картинке>\n"
        "ПЕРЕВОД:\n<перевод>\n"
        "Правила:\n"
        "1. Сохраняй структуру: строки, абзацы, нумерацию, списки.\n"
        "2. Имена, числа, единицы и коды оставляй без изменений.\n"
        "3. Не добавляй пояснений и комментариев."
    )
    user = {
        "role": "user",
        "content": [
            {"type": "text", "text": "Переведи текст с картинки."},
            {"type": "image_url", "image_url": {"url": _image_to_data_url(image)}},
        ],
    }
    return [{"role": "system", "content": system}, user]


def translate_image_via_ai(image: Image.Image) -> tuple[str, str]:
    """Переводит картинку напрямую vision-моделью.

    Возвращает (перевод, распознанный_оригинал). Если модель вернула и оригинал,
    и перевод — показываем оба; иначе оригинал пустой.
    """
    result = _chat_completion(_build_vision_messages(image), label="Перевод изображения (vision)")
    return _split_original_vs_translation(result)


def _split_original_vs_translation(text: str) -> tuple[str, str]:
    """Выделяет из ответа модель оригинал и перевод, если он оформлен блоками."""
    text = (text or "").strip()
    # Ожидаемые маркеры: "ОРИГИНАЛ:" и "ПЕРЕВОД:"
    match = re.search(r"(?im)^\s*(?:оригинал|original)\s*:\s*(.+?)\s*^\s*(?:перевод|translation)\s*:\s*(.+)\Z", text, re.S)
    if match:
        original = match.group(1).strip()
        translation = match.group(2).strip()
        return translation, original
    return text, ""


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
        self._spinner_angle = 0
        self._spinner_job: "str | None" = None
        self._tick_job: "str | None" = None
        # Состояние прогресса
        self._start = time.time()
        self._percent = 0.0          # целевой процент
        self._shown = 0.0            # отображаемый (плавно догоняет целевой)
        self._stage = "Обработка…"

    def show(self, title: str = "Обработка…") -> None:
        self._start = time.time()
        self._percent = 0.0
        self._shown = 0.0
        self._stage = title
        run_on_ui(lambda: self._build(title))

    def build(self, title: str = "Обработка…") -> None:
        """Создаёт окно, только если его ещё нет; иначе просто меняет этап."""
        self._start = time.time()
        self._percent = 0.0
        self._shown = 0.0
        self._stage = title
        run_on_ui(lambda: self._ensure_built(title))

    def _ensure_built(self, title: str) -> None:
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
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, sh - h - margin)
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

    def update(self, percent: float, stage: str | None = None) -> None:
        run_on_ui(lambda: self._update(percent, stage))

    def finish(self, stage: str = "Готово") -> None:
        """Доводит шкалу до 100% и показывает финальный этап."""
        run_on_ui(lambda: self._update(100, stage))

    def _update(self, percent: float, stage: str | None) -> None:
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

    def close(self) -> None:
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


def show_result(image: Image.Image, translation: str, source_text: str = "") -> None:
    """Показывает окно результата: картинка слева, карточка перевода справа.

    Одновременно открыто только одно окно результата: предыдущее закрывается.
    """
    def build() -> None:
        global _result_window
        # Закрываем предыдущее окно результата, чтобы они не копились.
        if _result_window is not None:
            try:
                if _result_window.winfo_exists():
                    _result_window.destroy()
            except tk.TclError:
                pass
            _result_window = None

        win = tk.Toplevel(_ui_root)
        _result_window = win
        win.title("Перевод с экрана")
        win.configure(bg=Color.SURFACE)
        win.attributes("-topmost", bool(CONFIG.get("keep_image_on_top")))
        win.bind("<Escape>", lambda _e: win.destroy())

        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)

        frame = tk.Frame(win, bg=Color.SURFACE)
        frame.grid(row=0, column=0, sticky="nsew", padx=Space.LG, pady=Space.LG)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=0)
        frame.columnconfigure(1, weight=1)

        # --- Изображение слева ---
        img_label = tk.Label(frame, bg=Color.SURFACE, bd=0, highlightthickness=0)
        img_label.grid(row=0, column=0, padx=(0, Space.LG), pady=0, sticky="n")
        preview = _fit_for_preview(image, max_w=460, max_h=640)
        img_tk = ImageTk.PhotoImage(preview)
        img_label.configure(image=img_tk)
        img_label.image = img_tk  # type: ignore[attr-defined]  # защита от сборщика мусора

        # --- Правая колонка ---
        right = tk.Frame(frame, bg=Color.SURFACE)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)

        # Заголовок с источником: «Перевод · нейросеть» / «Перевод · Tesseract»
        mode_label = "нейросеть" if CONFIG.get("ocr_mode", "ai") == "ai" else "Tesseract"
        title = tk.Label(right, text=f"Перевод · {mode_label}", bg=Color.SURFACE,
                         fg=Color.TEXT_ON_DARK, font=Font.TITLE, anchor="w")
        title.grid(row=0, column=0, sticky="ew")

        has_source = bool((source_text or "").strip())
        view = {"mode": "translation"}

        # Карточка текста создаётся до переключателя (нужна для on_view).
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
            width=46, height=18, cursor="arrow",
            selectbackground=Color.PRIMARY, selectforeground=Color.TEXT_ON_DARK,
        )
        text_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar.config(command=text_widget.yview)

        def current_text() -> str:
            return (source_text if view["mode"] == "source" else translation) or ""

        # Сегментированный переключатель «Перевод / Оригинал» (только если есть оригинал).
        if has_source:
            view_var = tk.StringVar(value="translation")

            def on_view() -> None:
                view["mode"] = view_var.get()
                _render_translation(text_widget, current_text())

            seg = _segmented(right, [("translation", "Перевод"), ("source", "Оригинал")],
                             view_var, command=on_view)
            seg.grid(row=1, column=0, sticky="w", pady=(Space.SM, Space.SM))
        else:
            tk.Frame(right, bg=Color.SURFACE, height=Space.SM).grid(row=1, column=0)
            view_var = tk.StringVar(value="translation")

        def copy_all() -> None:
            win.clipboard_clear()
            win.clipboard_append(current_text())
            logger.info("Скопировано в буфер обмена (%s).", view["mode"])

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

        _render_translation(text_widget, current_text())

        # --- Панель действий ---
        actions = tk.Frame(win, bg=Color.SURFACE)
        actions.grid(row=1, column=0, sticky="ew", padx=Space.LG, pady=(0, Space.LG))
        _button(actions, "Копировать", copy_all, variant="primary").pack(side="left", padx=(0, Space.SM))
        _button(actions, "Настройки", open_settings, variant="secondary").pack(side="left", padx=(0, Space.SM))
        _button(actions, "Закрыть", win.destroy, variant="ghost").pack(side="right")

        # Контекстное меню текста
        context_menu = tk.Menu(text_widget, tearoff=0)
        context_menu.add_command(label="Копировать", command=copy_selected)
        context_menu.add_command(label="Копировать всё", command=copy_all)
        text_widget.bind("<Button-3>", lambda e: context_menu.tk_popup(e.x_root, e.y_root))

        # Горячие клавиши окна.
        win.bind("<Control-c>", lambda _e: copy_all())
        if has_source:
            win.bind("<Control-o>", lambda _e: view_var.set(
                "source" if view_var.get() == "translation" else "translation") or on_view())

        _center_window(win)

    run_on_ui(build)


def _fit_for_preview(image: Image.Image, max_w: int, max_h: int) -> Image.Image:
    w, h = image.size
    if w <= max_w and h <= max_h:
        return image
    scale = min(max_w / w, max_h / h)
    return image.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.Resampling.LANCZOS)


def _render_translation(text_widget: tk.Text, translation: str) -> None:
    """Аккуратно выводит перевод: заголовок-источник, разделитель, перевод."""
    text_widget.configure(state="normal")
    text_widget.delete("1.0", tk.END)
    text_widget.insert(tk.END, (translation or "").strip() or "— пусто —")
    text_widget.configure(state="disabled")


def _center_window(win) -> None:
    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    x = max(0, (screen_w - w) // 2)
    y = max(0, (screen_h - h) // 3)
    win.geometry(f"+{x}+{y}")
    win.focus_force()


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
        win.title("Настройки перевода")
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
        tk.Checkbutton(common, text="Автозапуск с Windows", variable=autostart_var,
                       **check_opts).grid(row=2, column=0, columnspan=2, sticky="w", pady=(Space.XS, 0))
        on_top_var = tk.BooleanVar(value=bool(CONFIG.get("keep_image_on_top")))
        tk.Checkbutton(common, text="Окно перевода поверх других", variable=on_top_var,
                       **check_opts).grid(row=2, column=2, columnspan=2, sticky="w", pady=(Space.XS, 0))

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
    # Синхронизируем автозапуск, если он поменялся.
    want = bool(CONFIG.get("add_to_autostart"))
    if want != is_autostart_enabled():
        set_autostart(want)


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
        key = os.environ.get(section["api_key_env"], "").strip()
    if not key and provider_id == "deepseek":
        key = os.environ.get("DEEPSEEK_API_KEY", "").strip() or DEFAULT_DEEPSEEK_API_KEY

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
        raise RuntimeError("401 — неверный API-ключ")
    if response.status_code == 404:
        raise RuntimeError(f"404 — модель '{model}' не найдена")
    response.raise_for_status()
    data = response.json()
    return data["choices"][0]["message"]["content"].strip()


# ---------------------------------------------------------------------------
# Горячая клавиша и логика обработки скриншота
# ---------------------------------------------------------------------------
process_pending = False


def on_hotkey() -> None:
    """Обработка Win+Shift+S.

    Windows сам открывает оверлей выделения области. Поэтому мы не ждём кликов
    мышью (как раньше — из-за этого ничего не показывалось), а запоминаем текущий
    номер буфера обмена и ждём в нём НОВУЮ картинку — результат выделения.
    """
    global process_pending
    if process_pending:
        return
    process_pending = True
    try:
        seq_before = _clipboard_sequence()
        logger.info("Win+Shift+S — жду результат выделения в буфере обмена…")
        progress = _progress()
        progress.build("Ожидание выделения…")
        progress.update(0, "Выделите область в оверлее Windows")

        image = None
        start = time.time()
        deadline = start + 60.0  # даём время выбрать область
        while time.time() < deadline:
            # Новая картинка в буфере (или буфер изменился).
            if _clipboard_sequence() != seq_before:
                candidate = get_image_from_clipboard()
                if candidate and isinstance(candidate, Image.Image):
                    image = candidate
                    break
            elapsed = time.time() - start
            progress.update(min(90.0, elapsed / 60.0 * 90.0), "Ожидание выделения…")
            time.sleep(0.2)

        if image is None:
            logger.info("Скриншот не получен (выделение отменено или истёк таймаут).")
            return

        logger.info("Скриншот получен из буфера обмена (%dx%d).", *image.size)
        process_image(image)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Непредвиденная ошибка: %s", exc)
        _notify(f"Произошла ошибка:\n{exc}")
    finally:
        _progress().close()
        process_pending = False


def _notify(message: str) -> None:
    """Показывает короткое информационное окно в едином Tk-потоке."""
    def build() -> None:
        win = tk.Toplevel(_ui_root)
        win.withdraw()
        win.attributes("-topmost", True)
        messagebox.showinfo("Перевод с экрана", message, parent=win)
        win.destroy()

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
        keyboard.unhook_all()
        os._exit(0)


def on_settings(icon, _item) -> None:
    run_on_ui(lambda: open_settings_window(None))


def on_toggle_autostart(icon, _item) -> None:
    set_autostart(not is_autostart_enabled())
    icon.menu = _build_menu()


def on_toggle_live(icon, _item) -> None:
    set_live_mode(not is_live_mode())
    icon.menu = _build_menu()


def on_toggle_top(icon, _item) -> None:
    CONFIG["keep_image_on_top"] = not bool(CONFIG.get("keep_image_on_top"))
    save_config(CONFIG)
    icon.menu = _build_menu()


def _build_menu():
    from pystray import Menu, MenuItem

    return Menu(
        MenuItem("Режим «на лету»", on_toggle_live, checked=lambda _i: is_live_mode()),
        MenuItem("Автозапуск", on_toggle_autostart, checked=lambda _i: is_autostart_enabled()),
        MenuItem("Окно поверх других", on_toggle_top,
                 checked=lambda _i: bool(CONFIG.get("keep_image_on_top"))),
        Menu.SEPARATOR,
        MenuItem("Настройки…", on_settings),
        Menu.SEPARATOR,
        MenuItem("Выход", on_exit),
    )


def start_tray_icon() -> None:
    from pystray import Icon

    icon = Icon("pyAutoImgTranslate", create_image_for_tray(), "Перевод с экрана (Win+Shift+S)", _build_menu())
    icon.run()


# ---------------------------------------------------------------------------
# Точка входа
# ---------------------------------------------------------------------------
# Уникальное имя для SingleInstance
_MUTEX_NAME = "Global\\pyAutoImgTranslate_SingleInstance"


def _already_running() -> bool:
    """Возвращает True, если копия приложения уже запущена."""
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


def main() -> None:
    setup_logging()
    logger.info("Запуск pyAutoImgTranslate. Провайдер=%s, модель=%s",
                CONFIG.get("active_provider"), CONFIG.get("active_model"))

    if _already_running():
        logger.warning("Приложение уже запущено — второй экземпляр не стартует.")
        sys.exit(0)

    ensure_autostart_consistency()

    tray_thread = threading.Thread(target=start_tray_icon, daemon=True)
    tray_thread.start()

    hotkey = CONFIG.get("hotkey", "windows+shift+s")
    live_hotkey = CONFIG.get("live_hotkey", "windows+shift+a")
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


def _toggle_live_from_hotkey() -> None:
    enabled = not is_live_mode()
    set_live_mode(enabled)
    logger.info("Режим «на лету» %s (горячая клавиша).", "включён" if enabled else "выключен")
    _notify("Режим «на лету»: " + ("ВКЛ" if enabled else "ВЫКЛ"))


if __name__ == "__main__":
    main()
