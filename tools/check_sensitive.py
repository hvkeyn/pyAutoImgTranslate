"""Проверка репозитория на чувствительные данные перед публикацией.

Запуск:
    python tools/check_sensitive.py            # индекс, рабочее дерево и история
    python tools/check_sensitive.py --staged   # только то, что уже добавлено в индекс
    python tools/check_sensitive.py --no-history
    python tools/check_sensitive.py --strict   # личные пути считаются ошибкой

Код возврата 1 — публиковать нельзя; 0 — всё чисто (предупреждения не блокируют,
если не указан --strict).

Проверяется три вещи:
  * пути — локальные настройки, журналы, отладочные снимки, сборка, ключи;
  * содержимое текстовых файлов — ключи API, токены, приватные ключи, пароли;
  * история git — то же самое во всех коммитах, включая удалённые файлы.

Найденные секреты печатаются замаскированными: полное значение в вывод не попадает.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# --- Что не должно попадать в репозиторий вообще ---------------------------
FORBIDDEN_PATHS = [
    (re.compile(r"(^|/)config\.json$"), "локальные настройки (могут содержать API-ключ)"),
    (re.compile(r"\.log(\.\d+)?$"), "журнал работы (может содержать текст с экрана)"),
    (re.compile(r"(^|/)debug_[^/]*\.(png|jpe?g|webp|bmp)$", re.I), "отладочный снимок экрана"),
    (re.compile(r"(^|/)build/"), "артефакты сборки"),
    (re.compile(r"(^|/)dist(-linux)?/"), "артефакты сборки"),
    (re.compile(r"(^|/)\.venv/"), "виртуальное окружение"),
    (re.compile(r"(^|/)__pycache__/"), "кэш Python"),
    (re.compile(r"\.pyc$"), "байткод Python"),
    (re.compile(r"(^|/)\.credentials"), "файл с учётными данными"),
    (re.compile(r"(^|/)\.env(\.[^/]*)?$"), "файл с секретами окружения"),
    (re.compile(r"\.(pem|key|p12|pfx|keystore)$", re.I), "ключ или сертификат"),
    (re.compile(r"(^|/)id_(rsa|ed25519|ecdsa)"), "приватный SSH-ключ"),
]

# Картинки, которые в репозитории допустимы. Всё остальное — повод проверить,
# не скриншот ли это с личными данными.
ALLOWED_IMAGES = {"translator.png"}
IMAGE_RE = re.compile(r"\.(png|jpe?g|webp|bmp|gif)$", re.I)

# --- Что не должно встречаться в тексте -----------------------------------
SECRET_PATTERNS = [
    (re.compile(r"\bsk-[A-Za-z0-9_-]{20,}"), "похоже на API-ключ (sk-…)"),
    (re.compile(r"\bghp_[A-Za-z0-9]{30,}"), "токен GitHub"),
    (re.compile(r"\bgithub_pat_[A-Za-z0-9_]{30,}"), "токен GitHub"),
    (re.compile(r"\bxox[abprs]-[A-Za-z0-9-]{10,}"), "токен Slack"),
    (re.compile(r"\bAKIA[0-9A-Z]{16}"), "ключ AWS"),
    (re.compile(r"\bAIza[0-9A-Za-z_-]{35}"), "ключ Google"),
    (re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----"), "приватный ключ"),
    (re.compile(r"\bBearer\s+[A-Za-z0-9._~+/-]{20,}"), "токен в заголовке Authorization"),
    (re.compile(r"(?i)api[_-]?key\"?\s*[:=]\s*\"([^\"\n]{8,})\""), "ключ в настройках"),
    (re.compile(r"(?i)password\"?\s*[:=]\s*\"([^\"\n]{4,})\""), "пароль"),
]
# Значения-заглушки: это не секреты.
PLACEHOLDER_MARKERS = ("replace", "your-key", "your_key", "example", "xxxx", "placeholder",
                       "changeme", "lm-studio", "test-key", "none", "null",
                       "замени", "вставь", "свой", "ваш")

PERSONAL_PATH_RE = re.compile(r"(?i)\b[A-Z]:\\Users\\([A-Za-z0-9._-]+)|/home/([A-Za-z0-9._-]+)")
ALLOWED_USER_NAMES = {"public", "default", "user", "username", "yourname", "<user>", "hex"}

TEXT_SUFFIXES = {".py", ".md", ".txt", ".json", ".yml", ".yaml", ".toml", ".ini", ".cfg",
                 ".bat", ".cmd", ".sh", ".ps1", ".spec", ".html", ".css", ".js", ".ts",
                 ".iml", ".xml", ".gitignore", ".gitattributes"}

errors: list[str] = []
warnings: list[str] = []
scanned_files = 0
scanned_blobs = 0


def git(*args: str) -> str:
    result = subprocess.run(["git", *args], cwd=ROOT, capture_output=True)
    return result.stdout.decode("utf-8", "replace")


def mask(value: str) -> str:
    value = value.strip()
    if len(value) <= 8:
        return value[:2] + "***"
    return f"{value[:4]}***{value[-2:]}"


def is_placeholder(value: str) -> bool:
    lowered = (value or "").strip().lower()
    if not lowered:
        return True
    # Настоящие ключи всегда ASCII: кириллица означает текст-заглушку или пояснение.
    if any(ord(char) > 127 for char in value):
        return True
    return any(marker in lowered for marker in PLACEHOLDER_MARKERS)


def check_path(path: str, source: str) -> None:
    """Проверяет путь на запрещённые категории."""
    normalized = path.replace("\\", "/")
    for pattern, reason in FORBIDDEN_PATHS:
        if pattern.search(normalized):
            errors.append(f"{source}: {path} — {reason}")
            return
    name = Path(normalized).name
    if IMAGE_RE.search(name) and name not in ALLOWED_IMAGES and not normalized.startswith(".idea/"):
        warnings.append(f"{source}: {path} — картинка вне allowlist "
                        f"(если это не скриншот с личными данными, добавьте в ALLOWED_IMAGES)")


def scan_text(text: str, source: str) -> None:
    """Ищет секреты и личные пути в тексте."""
    for pattern, reason in SECRET_PATTERNS:
        for match in pattern.finditer(text):
            value = match.group(1) if match.groups() else match.group(0)
            if is_placeholder(value):
                continue
            errors.append(f"{source}: {reason}: {mask(value)}")
    for match in PERSONAL_PATH_RE.finditer(text):
        user = match.group(1) or match.group(2) or ""
        if user.lower() in ALLOWED_USER_NAMES:
            continue
        warnings.append(f"{source}: личный путь в тексте ({mask(user)})")


def scan_blob(blob: bytes, path: str, source: str) -> None:
    """Проверяет содержимое файла (из индекса, дерева или истории)."""
    global scanned_files
    if not blob or len(blob) > 1_000_000 or IMAGE_RE.search(path):
        return
    suffix = Path(path).suffix.lower()
    if suffix not in TEXT_SUFFIXES and Path(path).name not in (".gitignore", ".gitattributes"):
        return
    scanned_files += 1
    scan_text(blob.decode("utf-8", "replace"), f"{source}: {path}" if source != "репозиторий"
              else path)


def check_tree(staged: bool) -> None:
    """Проверяет файлы индекса или всего рабочего дерева."""
    if staged:
        raw = git("diff", "--cached", "--name-only", "-z")
        source = "индекс"
    else:
        raw = git("ls-files", "-z")
        source = "репозиторий"
    for path in filter(None, raw.split("\0")):
        check_path(path, source)
        try:
            if staged:
                blob = subprocess.run(["git", "show", f":{path}"], cwd=ROOT,
                                      capture_output=True).stdout
            else:
                blob = (ROOT / path).read_bytes()
        except OSError:
            continue
        scan_blob(blob, path, source)


def check_history() -> None:
    """Проверяет все коммиты: запрещённые пути и секреты в текстовых блобах."""
    global scanned_blobs
    objects = git("rev-list", "--objects", "--all")
    seen_blobs: set[str] = set()
    for line in objects.splitlines():
        parts = line.split(" ", 1)
        if len(parts) != 2:
            continue
        sha, path = parts
        check_path(path, "история")
        if sha in seen_blobs:
            continue
        seen_blobs.add(sha)
        blob = subprocess.run(["git", "cat-file", "blob", sha], cwd=ROOT,
                              capture_output=True).stdout
        if not blob or len(blob) > 1_000_000 or IMAGE_RE.search(path):
            continue
        if Path(path).suffix.lower() not in TEXT_SUFFIXES and Path(path).name not in (
                ".gitignore", ".gitattributes"):
            continue
        scanned_blobs += 1
        scan_text(blob.decode("utf-8", "replace"), f"история:{path}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Проверка репозитория на чувствительные данные")
    parser.add_argument("--staged", action="store_true", help="только индекс")
    parser.add_argument("--no-history", action="store_true", help="не проверять историю")
    parser.add_argument("--strict", action="store_true", help="личные пути — тоже ошибка")
    args = parser.parse_args()

    if git("rev-parse", "--is-inside-work-tree").strip() != "true":
        print("Это не git-репозиторий.", file=sys.stderr)
        return 2

    check_tree(args.staged)
    if not args.no_history and not args.staged:
        check_history()

    if args.strict:
        errors.extend(warnings)
        warnings.clear()

    for warning in warnings:
        print(f"  предупреждение: {warning}")
    for error in errors:
        print(f"  ОШИБКА: {error}")

    print(f"\nПроверено: файлов {scanned_files}, коммитов-блобов {scanned_blobs}, "
          f"ошибок {len(errors)}, предупреждений {len(warnings)}")
    if errors:
        print("Публиковать нельзя: уберите перечисленное (или перепишите историю).")
        return 1
    print("Чувствительных данных не найдено — можно коммитить и пушить.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
