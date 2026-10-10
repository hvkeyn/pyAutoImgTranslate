"""Проверка перевода через локальный OpenAI-совместимый сервер-заглушку.

Запуск:
    .venv\\Scripts\\python.exe tests\\test_translation_flow.py

Проверяются: потоковая выдача (SSE) и живой показ текста, обычный JSON-ответ,
сервер без поддержки потока, отмена запроса на середине генерации и понятные
ошибки (401). Сеть наружу не используется — сервер поднимается локально.
"""

from __future__ import annotations

import json
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PIL import Image  # noqa: E402

import main as app  # noqa: E402

FAILURES: list[str] = []
STATE: dict = {"pieces": ["Привет"], "delay": 0.0, "mode": "stream", "requests": 0}


def check(name: str, condition: bool, detail: str = "") -> None:
    print(f"  {'ok  ' if condition else 'FAIL'} {name} {'' if condition else detail}".rstrip())
    if not condition:
        FAILURES.append(name)


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.0"

    def log_message(self, *args) -> None:      # тишина в консоли
        pass

    def do_POST(self) -> None:  # noqa: N802
        length = int(self.headers.get("Content-Length") or 0)
        payload = json.loads(self.rfile.read(length) or b"{}")
        STATE["requests"] += 1
        want_stream = bool(payload.get("stream"))

        if STATE["mode"] == "unauthorized":
            body = b'{"error": {"message": "invalid key"}}'
            self.send_response(401)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if want_stream and STATE["mode"] != "json-only":
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self.end_headers()
            for piece in STATE["pieces"]:
                chunk = json.dumps({"choices": [{"delta": {"content": piece}}]})
                try:
                    self.wfile.write(f"data: {chunk}\n\n".encode("utf-8"))
                    self.wfile.flush()
                except OSError:
                    return
                if STATE["delay"]:
                    time.sleep(STATE["delay"])
            try:
                self.wfile.write(b"data: [DONE]\n\n")
                self.wfile.flush()
            except OSError:
                pass
            return

        body = json.dumps({"choices": [{"message": {"content": "".join(STATE["pieces"])}}]},
                          ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def start_server() -> tuple[ThreadingHTTPServer, str]:
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    host, port = server.server_address[:2]
    return server, f"http://{host}:{port}/v1"


def use_stub_provider(base_url: str) -> tuple[dict, dict]:
    """Переключает приложение на локальный сервер, возвращая прежние настройки."""
    original = {"active_provider": app.CONFIG.get("active_provider"),
                "active_model": app.CONFIG.get("active_model"),
                "stream_results": app.CONFIG.get("stream_results")}
    stub = {"id": "stub", "name": "Stub", "base_url": base_url, "api_key": "test-key",
            "api_key_env": "", "models": ["stub-model"]}
    app.CONFIG["providers"] = [p for p in app.CONFIG["providers"] if p.get("id") != "stub"] + [stub]
    app.CONFIG["active_provider"] = "stub"
    app.CONFIG["active_model"] = "stub-model"
    return original, stub


def restore(original: dict) -> None:
    app.CONFIG.update(original)
    app.CONFIG["providers"] = [p for p in app.CONFIG["providers"] if p.get("id") != "stub"]


def test_stream_and_json() -> None:
    server, base_url = start_server()
    original, _ = use_stub_provider(base_url)
    try:
        # 1. Поток: перевод приходит кусками, окно получает промежуточные версии.
        STATE.update({"pieces": ["ПЕРЕВОД:\nПри", "вет, ", "мир"], "delay": 0.05,
                      "mode": "stream"})
        deltas: list[str] = []
        job = app._new_job("hotkey")
        result = app.translate_text("hello", job=job, on_delta=deltas.append)
        app._finish_job(job)
        check("поток: перевод собран", result == "Привет, мир", repr(result))
        check("поток: промежуточные версии приходили", len(deltas) >= 2, repr(deltas))
        check("поток: в окно идёт текст без маркеров",
              [app._stream_visible_text(d) for d in deltas][-1] == "Привет, мир",
              repr([app._stream_visible_text(d) for d in deltas]))
        check("поток: сервер получил запрос с stream=true", STATE["requests"] >= 1)

        # 2. Сервер игнорирует поток и отвечает обычным JSON.
        STATE.update({"pieces": ["Обычный ответ"], "mode": "json-only"})
        deltas.clear()
        result = app.translate_text("hello", job=app._current_job(), on_delta=deltas.append)
        check("JSON вместо потока: текст получен", result == "Обычный ответ", repr(result))

        # 3. Поток выключен настройкой — обычный запрос.
        app.CONFIG["stream_results"] = False
        STATE.update({"pieces": ["Без потока"], "mode": "stream"})
        result = app.translate_text("hello")
        check("поток выключен: текст получен", result == "Без потока", repr(result))
        app.CONFIG["stream_results"] = True

        # 4. Картинка: два блока — перевод и оригинал.
        STATE.update({"pieces": ["ПЕРЕВОД:\nПривет\nОРИГИНАЛ:\nhello"], "mode": "stream"})
        image = Image.new("RGB", (120, 60), "white")
        translation, source = app.translate_image_via_ai(image, job=app._current_job(),
                                                         on_delta=lambda _t: None)
        check("vision: перевод выделен", translation == "Привет", repr(translation))
        check("vision: оригинал выделен", source == "hello", repr(source))

        # 5. Модель сообщает, что текста нет.
        STATE.update({"pieces": ["NO_TEXT"], "mode": "stream"})
        translation, source = app.translate_image_via_ai(image, job=app._current_job(),
                                                         on_delta=lambda _t: None)
        check("vision: NO_TEXT распознан", translation == "" and source == "",
              f"{translation!r} {source!r}")

        # 6. Ошибка авторизации объясняется понятно.
        STATE.update({"mode": "unauthorized"})
        try:
            app.translate_text("hello")
            check("ошибка 401: понятное сообщение", False, "исключения не было")
        except RuntimeError as exc:
            check("ошибка 401: понятное сообщение", "401" in str(exc), str(exc))
    finally:
        restore(original)
        server.shutdown()


def test_cancel_closes_request() -> None:
    server, base_url = start_server()
    original, _ = use_stub_provider(base_url)
    try:
        # Сервер отдаёт куски медленно: отмена должна прервать ожидание сразу.
        STATE.update({"pieces": ["раз", "два", "три", "четыре", "пять"],
                      "delay": 0.6, "mode": "stream"})
        job = app._new_job("hotkey")
        started = time.time()

        def canceller() -> None:
            time.sleep(0.4)
            app.cancel_current_job("тест отмены")

        threading.Thread(target=canceller, daemon=True).start()
        try:
            app.translate_text("hello", job=job, on_delta=lambda _t: None)
            check("отмена: прерывает запрос", False, "исключения не было")
        except app.Cancelled:
            elapsed = time.time() - started
            check("отмена: прерывает запрос", True)
            check("отмена: не ждёт конца генерации", elapsed < 2.0, f"{elapsed:.2f} с")
        finally:
            app._finish_job(job)
    finally:
        restore(original)
        server.shutdown()


def main() -> int:
    print("Проверка перевода (локальный сервер-заглушка)")
    test_stream_and_json()
    test_cancel_closes_request()
    print(f"\nПровалов: {len(FAILURES)}")
    return 1 if FAILURES else 0


if __name__ == "__main__":
    raise SystemExit(main())
