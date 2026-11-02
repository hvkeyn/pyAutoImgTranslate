"""Проверка интерфейсных путей pyAutoImgTranslate в реальном Tk (без сети).

Запуск:
    .venv\\Scripts\\python.exe tests\\test_gui_smoke.py

На короткое время появятся окна приложения (окно перевода, плашка прогресса,
подсказка) — они закрываются автоматически. Проверяется, что окно перевода
создаётся, обновляется по мере генерации, финализируется и закрывается, а
окно прогресса с кнопкой «Отмена» не падает.
"""

from __future__ import annotations

import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PIL import Image, ImageDraw  # noqa: E402

import main as app  # noqa: E402

FAILURES: list[str] = []


def check(name: str, condition: bool, detail: str = "") -> None:
    status = "ok  " if condition else "FAIL"
    print(f"  {status} {name} {detail if not condition else ''}".rstrip())
    if not condition:
        FAILURES.append(name)


def sample_image() -> Image.Image:
    image = Image.new("RGB", (900, 400), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((40, 40, 860, 120), outline="black", width=2)
    draw.text((60, 60), "HELLO WORLD", fill="black")
    return image


def wait_for(predicate, timeout: float = 5.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if predicate():
            return True
        time.sleep(0.05)
    return predicate()


def test_hotkey_wait_loop() -> None:
    """Горячая клавиша: ожидание картинки в буфере, перес-скриншот и отмена."""
    fake = {"seq": 100, "processed": [], "image": sample_image()}
    original = {
        "sequence": app._clipboard_sequence,
        "clipboard": app.get_image_from_clipboard,
        "process": app.process_image,
        "mode": app.CONFIG.get("capture_mode"),
    }
    app.CONFIG["capture_mode"] = "system"        # проверяем системный путь
    app._clipboard_sequence = lambda: fake["seq"]
    app.get_image_from_clipboard = lambda: fake["image"]
    app.process_image = lambda image, job=None: fake["processed"].append(job.kind)
    try:
        # 1. Картинка появляется в буфере — задача уходит в обработку.
        worker = threading.Thread(target=app.on_hotkey, daemon=True)
        worker.start()
        check("горячая клавиша: задача зарегистрирована",
              wait_for(lambda: app._current_job() is not None, timeout=3.0))
        fake["seq"] += 1
        check("горячая клавиша: снимок ушёл в обработку",
              wait_for(lambda: fake["processed"] == ["hotkey"], timeout=5.0), repr(fake["processed"]))
        check("горячая клавиша: задача завершена и снята",
              wait_for(lambda: app._current_job() is None, timeout=5.0))

        # 2. Перес-скриншот отменяет предыдущую задачу.
        fake["processed"].clear()
        fake["seq"] += 1
        first = threading.Thread(target=app.on_hotkey, daemon=True)
        first.start()
        check("перес-скриншот: первая задача ждёт",
              wait_for(lambda: app._current_job() is not None, timeout=3.0))
        first_job = app._current_job()
        second = threading.Thread(target=app.on_hotkey, daemon=True)
        second.start()
        check("перес-скриншот: первая задача отменена",
              wait_for(lambda: first_job.cancelled(), timeout=3.0))
        # Отменённая задача не должна ничего переводить.
        time.sleep(0.5)
        check("перес-скриншот: отменённая задача молчит", fake["processed"] == [], repr(fake["processed"]))

        # 3. Явная отмена: задача завершается без перевода.
        app.cancel_current_job("тест")
        check("отмена: текущих задач нет",
              wait_for(lambda: app._current_job() is None, timeout=5.0))
        check("отмена: перевода не было", fake["processed"] == [], repr(fake["processed"]))
    finally:
        app._clipboard_sequence = original["sequence"]
        app.get_image_from_clipboard = original["clipboard"]
        app.process_image = original["process"]
        app.CONFIG["capture_mode"] = original["mode"]
        app.cancel_current_job("завершение теста")


def test_capture_region_builtin() -> None:
    """Своё выделение: снимок из оверлея уходит в перевод, отмена ничего не делает."""
    processed: list[tuple[int, int]] = []
    original = {
        "select": app.select_region_interactive,
        "process": app.process_image,
        "mode": app.CONFIG.get("capture_mode"),
    }
    app.CONFIG["capture_mode"] = "builtin"
    app.process_image = lambda image, job=None: processed.append(image.size)
    try:
        app.select_region_interactive = lambda job=None: (sample_image(), "ok")
        app.on_hotkey()
        check("своё выделение: снимок ушёл в перевод", processed == [(900, 400)], repr(processed))

        processed.clear()
        app.select_region_interactive = lambda job=None: (None, "cancelled")
        app.on_hotkey()
        check("своё выделение: отмена — перевода нет", processed == [], repr(processed))
        check("своё выделение: после отмены задач нет", app._current_job() is None)
    finally:
        app.select_region_interactive = original["select"]
        app.process_image = original["process"]
        app.CONFIG["capture_mode"] = original["mode"]


def test_region_overlay() -> None:
    """Своё выделение области: рамка мышью, масштаб окна и отмена."""
    shot = Image.new("RGB", (800, 600), "white")
    draw = ImageDraw.Draw(shot)
    draw.rectangle((150, 150, 650, 450), fill=(30, 90, 160))

    overlay = app.RegionOverlay(shot, (0, 0, 800, 600))
    app.run_on_ui(lambda: overlay.build(None))
    check("выделение: оверлей открылся", wait_for(lambda: overlay.win is not None))
    app.run_on_ui(lambda: (overlay.begin(100, 120), overlay.update(500, 420),
                           overlay.finish(500, 420)))
    check("выделение: результат получен", overlay.done.wait(3.0))
    check("выделение: размер области",
          overlay.result is not None and overlay.result.size == (400, 300),
          repr(overlay.result.size if overlay.result else None))
    if overlay.result is not None:
        pixel = overlay.result.getpixel((200, 200))
        check("выделение: пиксели из нужного места", pixel == (30, 90, 160), repr(pixel))
    check("выделение: окно закрылось", overlay.win is None)

    # Масштаб: окно меньше снимка — координаты пересчитываются.
    overlay2 = app.RegionOverlay(shot, (0, 0, 400, 300))
    app.run_on_ui(lambda: overlay2.build(None))
    wait_for(lambda: overlay2.win is not None)
    app.run_on_ui(lambda: (overlay2.begin(50, 60), overlay2.finish(200, 220)))
    check("выделение: масштаб учтён", overlay2.done.wait(3.0))
    check("выделение: область при масштабе",
          overlay2.result is not None and overlay2.result.size == (300, 320),
          repr(overlay2.result.size if overlay2.result else None))

    # Отмена по Esc и слишком маленькая область.
    overlay3 = app.RegionOverlay(shot, (0, 0, 800, 600))
    app.run_on_ui(lambda: overlay3.build(None))
    wait_for(lambda: overlay3.win is not None)
    app.run_on_ui(lambda: overlay3.cancel("тест"))
    check("выделение: отмена", wait_for(overlay3.done.is_set, timeout=3.0)
          and overlay3.cancelled and overlay3.result is None)

    overlay4 = app.RegionOverlay(shot, (0, 0, 800, 600))
    app.run_on_ui(lambda: overlay4.build(None))
    wait_for(lambda: overlay4.win is not None)
    app.run_on_ui(lambda: (overlay4.begin(10, 10), overlay4.finish(14, 14)))
    check("выделение: крошечная область отменяется",
          overlay4.done.wait(3.0) and overlay4.result is None)

    # Если снять экран нельзя — сообщаем, что нужен системный захват.
    original_grab = app.grab_virtual_screen
    app.grab_virtual_screen = lambda: None
    try:
        image, reason = app.select_region_interactive(None)
        check("выделение: нет доступа к экрану → системный захват",
              image is None and reason == "unsupported", reason)
    finally:
        app.grab_virtual_screen = original_grab


def test_overlay_speed() -> None:
    """Оверлей открывается быстро, а движение мыши не перерисовывает экран."""
    shot = Image.new("RGB", (3840, 1440), "white")      # даже больше типичного экрана
    timings: dict = {}
    original_build = app.RegionOverlay.build

    def timed_build(self, job=None):
        started = time.time()
        original_build(self, job)
        timings["build"] = time.time() - started

    app.RegionOverlay.build = timed_build
    overlay = app.RegionOverlay(shot, (0, 0, 3840, 1440))
    try:
        app.run_on_ui(lambda: overlay.build(None))
        check("скорость: окно выделения открылось", wait_for(lambda: overlay.win is not None))
        build_time = timings.get("build", 99.0)
        print(f"    сборка окна выделения: {build_time * 1000:.0f} мс")
        check("скорость: окно строится быстрее 0.25 с", build_time < 0.25, f"{build_time:.3f} с")

        def move_many() -> None:
            overlay.begin(10, 10)
            started = time.time()
            for step in range(50):
                overlay.update(10 + step * 60, 10 + step * 25)
            timings["per_move"] = (time.time() - started) / 50
            overlay.cancel("замер")

        app.run_on_ui(move_many)
        check("скорость: движения мыши обрабатываются", wait_for(overlay.done.is_set, 5.0))
        per_move = timings.get("per_move", 9.0)
        print(f"    обработка одного движения мыши: {per_move * 1000:.2f} мс")
        check("скорость: движение мыши дешевле 20 мс", per_move < 0.02, f"{per_move:.4f} с")
    finally:
        app.RegionOverlay.build = original_build
        if overlay.win is not None:
            app.run_on_ui(overlay.close_window)


def test_overlay_window() -> None:
    """Окно перевода поверх экрана: прозрачность, сквозные клики, отрисовка."""
    window = app.OverlayWindow((0, 0, 600, 400))
    try:
        app.run_on_ui(window.build)
        check("оверлей: окно создано", wait_for(lambda: window.win is not None))
        transparent = str(window.win.attributes("-transparentcolor")).upper()
        check("оверлей: прозрачность по ключевому цвету",
              app.OVERLAY_KEY_COLOR.upper() in transparent, transparent)
        check("оверлей: окно исключено из снимков экрана", window._excluded)
        import ctypes

        style = ctypes.windll.user32.GetWindowLongW(window.hwnd, -20)
        check("оверлей: клики проходят сквозь окно", bool(style & 0x20), hex(style))
        check("оверлей: окно не в панели задач", bool(style & 0x80), hex(style))

        layouts = [((10, 10, 300, 80), ["Перевод", "строки"], 16)]
        app.run_on_ui(lambda: window.draw(layouts))
        check("оверлей: плашки нарисованы",
              wait_for(lambda: len(window.canvas.find_all()) >= 3))
        check("оверлей: область подсвечена рамкой",
              bool(window.canvas.find_withtag("area")),
              str(len(window.canvas.find_withtag("area"))))
        plates = len(window.canvas.find_all()) - len(window.canvas.find_withtag("area"))
        check("оверлей: рамка не мешает плашкам", plates >= 3, str(plates))

        app.CONFIG["overlay_highlight"] = False
        app.run_on_ui(lambda: window.draw(layouts))
        check("оверлей: подсветку можно отключить",
              wait_for(lambda: not window.canvas.find_withtag("area")))
        app.CONFIG["overlay_highlight"] = True

        app.run_on_ui(window.clear)
        check("оверлей: очистка", wait_for(lambda: len(window.canvas.find_all()) == 0))
    finally:
        app.run_on_ui(window.destroy)
    check("оверлей: окно закрыто", wait_for(lambda: window.win is None))


def test_overlay_controller() -> None:
    """Режим включается/выключается, рисует плашки и следит за сменой кадра."""
    image = Image.new("RGB", (900, 500), (24, 28, 36))
    draw = ImageDraw.Draw(image)
    draw.text((40, 40), "MISSION COMPLETE", fill=(240, 240, 240), font=app._overlay_font(30))
    draw.text((40, 120), "Press SPACE to continue", fill=(220, 220, 220),
              font=app._overlay_font(22))
    second = image.copy()
    draw2 = ImageDraw.Draw(second)
    draw2.rectangle((40, 120, 860, 180), fill=(24, 28, 36))
    draw2.text((40, 120), "Press ESC to quit", fill=(220, 220, 220), font=app._overlay_font(22))

    original = {
        "grab": app.ImageGrab.grab,
        "translate": app.translate_blocks,
        "area": app.CONFIG.get("overlay_area"),
        "interval": app.CONFIG.get("overlay_interval"),
        "highlight": app.CONFIG.get("overlay_highlight"),
    }
    frame = {"image": image}
    calls = {"count": 0}

    def fake_translate(texts, job=None):
        calls["count"] += 1
        return [f"перевод {i}" for i, _t in enumerate(texts, 1)]

    app.ImageGrab.grab = lambda **kwargs: frame["image"].copy()
    app.translate_blocks = fake_translate
    app.overlay_cache_clear()
    app.CONFIG["overlay_area"] = [0, 0, image.width, image.height]
    app.CONFIG["overlay_interval"] = 0.4
    app.CONFIG["overlay_stable"] = 0.0
    app.CONFIG["overlay_highlight"] = True

    def item_count() -> int:
        window = app.OVERLAY._window
        return len(window.canvas.find_all()) if window and window.canvas else 0

    try:
        check("оверлей: режим включился", app.OVERLAY.start())
        check("оверлей: режим отмечен включённым", app.is_overlay_enabled())
        check("оверлей: плашки появились на экране", wait_for(lambda: item_count() >= 3, 8.0))
        check("оверлей: кадр обработан", wait_for(lambda: calls["count"] >= 1, 5.0),
              str(calls["count"]))
        # Смена текста на экране: плашки обновляются без перезапуска режима.
        frame["image"] = second
        check("оверлей: смена кадра подхватилась",
              wait_for(lambda: "Press ESC to quit" in " ".join(
                  block.text for block in app.OVERLAY._blocks), 8.0),
              repr([block.text for block in app.OVERLAY._blocks]))
        check("оверлей: перевод для новой строки запрошен",
              wait_for(lambda: calls["count"] >= 2, 8.0), str(calls["count"]))
        check("оверлей: рамка области на месте", wait_for(
            lambda: bool((app.OVERLAY._window.canvas.find_withtag("area")
                          if app.OVERLAY._window and app.OVERLAY._window.canvas else []))))
    finally:
        app.OVERLAY.stop("тест")
        app.ImageGrab.grab = original["grab"]
        app.translate_blocks = original["translate"]
        app.CONFIG["overlay_area"] = original["area"]
        app.CONFIG["overlay_interval"] = original["interval"]
        app.CONFIG["overlay_highlight"] = original["highlight"]
        app.CONFIG["overlay_stable"] = 0.25
    check("оверлей: режим выключился", not app.is_overlay_enabled())
    check("оверлей: окно убрано", wait_for(lambda: app.OVERLAY._window is None))


def widget_texts(widget) -> list[str]:
    """Собирает тексты всех потомков окна (для проверки состава настроек)."""
    texts: list[str] = []
    try:
        for child in widget.winfo_children():
            try:
                texts.append(str(child.cget("text")))
            except Exception:  # noqa: BLE001
                pass
            texts.extend(widget_texts(child))
    except Exception:  # noqa: BLE001
        pass
    return texts


def test_settings_window() -> None:
    """Окно настроек: собирается и содержит управление выделением области."""
    original_balance = app.fetch_balance
    app.fetch_balance = lambda: ""          # не ходим в сеть за балансом
    try:
        app.run_on_ui(lambda: app.open_settings_window(None))
        check("настройки: окно открылось", wait_for(lambda: app._settings_window is not None, 6.0))
        window = app._settings_window
        if window is None:
            return
        texts = widget_texts(window)
        check("настройки: есть выбор выделения области",
              any("Выделение области" in t for t in texts), repr(texts[:8]))
        check("настройки: есть варианты выделения",
              any("Своё окно" in t for t in texts) and any("Системное" in t for t in texts))
        check("настройки: подсказка о клавише",
              any(t.startswith("клавиша:") for t in texts),
              repr([t for t in texts if "клавиш" in t]))
        app.run_on_ui(lambda: window.destroy())
        app._settings_window = None
        check("настройки: окно закрыто", wait_for(lambda: app._settings_window is None))
    finally:
        app.fetch_balance = original_balance


def test_analysis_window() -> None:
    """Окно анализа: заголовок, вкладки «Ответ / Источники», кликабельные ссылки."""
    job = app._new_job("hotkey", "analyze")
    gen = job.gen
    image = sample_image()
    sources = "1. Sony — Википедия\n   https://ru.wikipedia.org/wiki/Sony"
    try:
        app._result_partial(gen, "Скриншот карточки товара", image=image, focus=False,
                            on_top=False, kind="analysis")
        check("анализ: окно открылось", wait_for(app._result_alive))
        check("анализ: окно помечено как анализ", app._RESULT.get("kind") == "analysis")
        labels = widget_texts(app._RESULT["win"])
        check("анализ: заголовок окна", any("Анализ изображения" in t for t in labels),
              repr(labels[:12]))
        check("анализ: вкладки «Ответ» и «Источники»",
              "Ответ" in labels and "Источники" in labels, repr(labels[:12]))

        app._result_done(gen, image,
                         "Скриншот карточки товара\n\nПодробнее:\nЭто наушники Sony.",
                         sources, focus=False, on_top=False, kind="analysis")
        check("анализ: финальный текст",
              wait_for(lambda: "Подробнее" in (app._RESULT["translation"] or "")),
              repr(app._RESULT.get("translation")))

        def show_sources() -> None:
            app._RESULT["mode"] = "source"
            app._render_result(gen)

        app.run_on_ui(show_sources)
        widget = app._RESULT["text"]
        check("анализ: источники во вкладке",
              wait_for(lambda: "wikipedia.org" in widget.get("1.0", "end")))
        check("анализ: ссылка кликабельная", len(widget.tag_ranges("link")) >= 2,
              repr(widget.tag_ranges("link")))
    finally:
        app.run_on_ui(lambda: app._result_alive() and app._RESULT["win"].destroy())
        app._finish_job(job)


def test_analysis_hotkey_dispatch() -> None:
    """Клавиши перевода и анализа ведут в разные обработчики."""
    calls: list[tuple[str, str]] = []
    original = {
        "select": app.select_region_interactive,
        "analysis": app.process_analysis,
        "translate": app.process_image,
        "mode": app.CONFIG.get("capture_mode"),
    }
    app.CONFIG["capture_mode"] = "builtin"
    app.select_region_interactive = lambda job=None: (sample_image(), "ok")
    app.process_analysis = lambda image, job=None: calls.append(("analysis", job.purpose))
    app.process_image = lambda image, job=None: calls.append(("translate", job.purpose))
    try:
        app.on_hotkey_analyze()
        check("клавиша анализа: уходит в анализ", calls == [("analysis", "analyze")], repr(calls))
        calls.clear()
        app.on_hotkey()
        check("клавиша перевода: уходит в перевод", calls == [("translate", "translate")],
              repr(calls))
        calls.clear()
        app.on_hotkey_analyze()
        app.on_hotkey()
        check("обе клавиши не мешают друг другу",
              calls == [("analysis", "analyze"), ("translate", "translate")], repr(calls))
    finally:
        app.select_region_interactive = original["select"]
        app.process_analysis = original["analysis"]
        app.process_image = original["translate"]
        app.CONFIG["capture_mode"] = original["mode"]
        app.cancel_current_job("завершение теста")


def main() -> int:
    print("Проверка интерфейса (Tk)")
    image = sample_image()
    job = app._new_job("hotkey")
    gen = job.gen

    # Ожидание выделения: плашка прогресса с кнопкой «Отмена».
    progress = app._progress()
    progress.build("Ожидание выделения…", owner=gen)
    progress.update(30, "Выделите область", owner=gen)
    check("прогресс: плашка создана", wait_for(lambda: progress.win is not None))
    check("прогресс: кнопка «Отмена» есть", wait_for(lambda: progress._cancel_btn is not None))
    check("прогресс: чужая задача не управляет плашкой",
          progress._belongs_to(gen + 1000) is False)

    # Живой вывод перевода: окно появляется на первом куске текста.
    app._result_partial(gen, "Привет", image=image, focus=False, on_top=False)
    check("окно перевода: создано", wait_for(app._result_alive))
    app._result_partial(gen, "Привет, мир", image=image, focus=False, on_top=False)
    check("окно перевода: текст обновился",
          wait_for(lambda: (app._RESULT.get("translation") or "") == "Привет, мир"))

    # Промежуточный текст не должен показываться для неактуальной задачи.
    app._result_partial(gen + 1000, "чужой текст", image=image, focus=False, on_top=False)
    time.sleep(0.2)
    check("окно перевода: чужая задача игнорируется",
          (app._RESULT.get("translation") or "") == "Привет, мир")

    # Финализация: перевод + оригинал и скрытая кнопка отмены.
    app._result_done(gen, image, "Привет, мир", "HELLO WORLD", focus=False, on_top=False)
    check("окно перевода: финальный текст",
          wait_for(lambda: (app._RESULT.get("translation") or "") == "Привет, мир"))
    check("окно перевода: оригинал сохранён",
          wait_for(lambda: (app._RESULT.get("source") or "") == "HELLO WORLD"))
    cancel_button = app._RESULT.get("cancel")
    check("окно перевода: кнопка отмены скрыта после завершения",
          cancel_button is not None and wait_for(lambda: not cancel_button.winfo_ismapped()))

    # Отметка «отменено»: частичный перевод остаётся на экране.
    app._mark_result_cancelled(gen)
    time.sleep(0.2)
    check("окно перевода: частичный текст не выбрасывается", app._result_alive())

    # Пустое окно (текста не нашлось) закрывается.
    empty_job = app._new_job("live")
    empty_gen = empty_job.gen
    app._result_done(empty_gen, image, "", "", focus=False, on_top=False)
    check("окно перевода: пустое окно создано", wait_for(app._result_alive))
    app._result_abort(empty_gen)
    check("окно перевода: пустое окно закрыто", wait_for(lambda: not app._result_alive()))
    app._finish_job(empty_job)

    # Подсказки вместо модальных окон: показываются и сами исчезают.
    app._toast("Проверка подсказки", kind="info", seconds=0.8)
    check("подсказка: показана", wait_for(lambda: app._TOAST_STATE.get("win") is not None))
    check("подсказка: исчезла сама",
          wait_for(lambda: app._TOAST_STATE.get("win") is None, timeout=4.0))

    progress.close(owner=gen)
    check("прогресс: плашка закрыта", wait_for(lambda: progress.win is None))
    app._finish_job(job)
    check("задачи: текущей нет", app._current_job() is None)

    test_hotkey_wait_loop()
    test_capture_region_builtin()
    test_region_overlay()
    test_overlay_speed()
    test_analysis_window()
    test_analysis_hotkey_dispatch()
    test_overlay_window()
    test_overlay_controller()
    test_settings_window()

    print(f"\nПровалов: {len(FAILURES)}")
    for failure in FAILURES:
        print(f"  FAIL {failure}")
    return 1 if FAILURES else 0


if __name__ == "__main__":
    import os
    import sys as _sys

    _code = main()
    _sys.stdout.flush()
    _sys.stderr.flush()
    # Tk живёт в отдельном потоке: обычное завершение иногда подвисает на
    # финализации интерпретатора, поэтому выходим явно.
    os._exit(_code)
