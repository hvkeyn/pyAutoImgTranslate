"""Офлайн-проверки логики pyAutoImgTranslate: без сети, без GUI, без хоткеев.

Запуск:
    .venv\\Scripts\\python.exe tests\\test_offline.py         (Windows)
    python3 tests/test_offline.py                            (Linux)

Проверяются самые «нежные» места: фильтр «на снимке нет текста», потоковый
разбор ответа (SSE), разбор блоков «ПЕРЕВОД/ОРИГИНАЛ», отмена задач и
сборка строк из данных Tesseract.
"""

from __future__ import annotations

import glob
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import numpy as np  # noqa: E402
from PIL import Image, ImageDraw, ImageFilter, ImageFont  # noqa: E402

import main as app  # noqa: E402

FAILURES: list[str] = []
REPORT: list[str] = []


def check(name: str, condition: bool, detail: str = "") -> None:
    if condition:
        REPORT.append(f"  ok   {name}")
    else:
        FAILURES.append(f"  FAIL {name} {detail}")
        REPORT.append(f"  FAIL {name} {detail}")


# ---------------------------------------------------------------------------
# Синтетические картинки
# ---------------------------------------------------------------------------
FONT_CANDIDATES = (
    r"C:\Windows\Fonts\segoeui.ttf",
    r"C:\Windows\Fonts\arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
)


def _font(size: int):
    """Настоящий шрифт, если он есть: на битмап-шрифте Pillow OCR читается плохо."""
    for path in FONT_CANDIDATES:
        if Path(path).is_file():
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:                     # старые Pillow
        return ImageFont.load_default()


def text_image(lines: list[str], size=(760, 360), font_size=26) -> Image.Image:
    image = Image.new("RGB", size, "white")
    draw = ImageDraw.Draw(image)
    font = _font(font_size)
    y = 24
    for line in lines:
        draw.text((28, y), line, fill="black", font=font)
        y += int(font_size * 1.6)
    return image


def solid_image(size=(760, 360)) -> Image.Image:
    return Image.new("RGB", size, (32, 64, 96))


def gradient_image(size=(760, 360)) -> Image.Image:
    x = np.linspace(0, 255, size[0], dtype=np.uint8)
    gradient = np.tile(x, (size[1], 1))
    rgb = np.dstack([gradient, np.roll(gradient, 40, axis=1), 255 - gradient])
    return Image.fromarray(rgb, "RGB")


def noise_image(size=(760, 360), seed=1) -> Image.Image:
    rng = np.random.default_rng(seed)
    data = rng.integers(0, 256, size=(size[1], size[0], 3), dtype=np.uint8)
    return Image.fromarray(data, "RGB")


def photo_like_image(size=(760, 360), seed=2, blur=6) -> Image.Image:
    """Мягкие пятна, как на фотографии: ни строк, ни глифов."""
    rng = np.random.default_rng(seed)
    small = rng.integers(40, 210, size=(max(2, size[1] // 30), max(2, size[0] // 30), 3),
                         dtype=np.uint8)
    image = Image.fromarray(small, "RGB").resize(size, Image.Resampling.BICUBIC)
    return image.filter(ImageFilter.GaussianBlur(blur))


def photo_with_rects_image(size=(760, 360), seed=5) -> Image.Image:
    """Фотография с резкими прямоугольниками — «сложный» случай без текста."""
    image = photo_like_image(size, seed=seed, blur=2)
    rng = np.random.default_rng(seed)
    draw = ImageDraw.Draw(image)
    for _ in range(14):
        x, y = rng.integers(0, size[0] - 80), rng.integers(0, size[1] - 60)
        w, h = rng.integers(20, 80), rng.integers(15, 60)
        draw.rectangle((x, y, x + w, y + h), fill=tuple(int(v) for v in rng.integers(0, 255, 3)))
    return image


def icon_like_image(size=(760, 360)) -> Image.Image:
    """Крупные геометрические фигуры без текста."""
    image = Image.new("RGB", size, (245, 245, 245))
    draw = ImageDraw.Draw(image)
    draw.ellipse((80, 80, 260, 260), fill=(200, 60, 60))
    draw.rectangle((340, 90, 520, 250), fill=(60, 120, 200))
    draw.polygon([(600, 250), (680, 90), (740, 250)], fill=(60, 180, 90))
    return image


def ui_mock_image(size=(760, 360)) -> Image.Image:
    """Интерфейс без надписей: панели, кнопки, полосы."""
    image = Image.new("RGB", size, (240, 242, 246))
    draw = ImageDraw.Draw(image)
    draw.rectangle((20, 20, 740, 60), fill=(210, 214, 222))
    for i in range(3):
        draw.rectangle((20, 90 + i * 60, 200, 130 + i * 60), fill=(180, 186, 196))
        draw.rectangle((230, 90 + i * 60, 740, 130 + i * 60), fill=(228, 230, 236))
    draw.rectangle((20, 300, 120, 340), fill=(90, 130, 220))
    draw.rectangle((140, 300, 240, 340), fill=(150, 155, 165))
    return image


def two_letters_image() -> Image.Image:
    return text_image(["OK"], size=(240, 120), font_size=48)


def grain_photo_image(width=1600, height=1200, blur=4, grain=12, seed=7) -> Image.Image:
    """Зернистая «фотография»: сложный случай для любой эвристики."""
    rng = np.random.default_rng(seed)
    small = rng.integers(40, 210, size=(max(2, height // 30), max(2, width // 30), 3),
                         dtype=np.uint8)
    image = Image.fromarray(small, "RGB").resize((width, height), Image.Resampling.BICUBIC)
    image = image.filter(ImageFilter.GaussianBlur(blur))
    data = np.asarray(image).astype(np.int16) + rng.normal(0, grain, (height, width, 3))
    return Image.fromarray(np.clip(data, 0, 255).astype(np.uint8), "RGB")


def big_text_image(lines: int = 40, size=(3840, 2160), font_size=28) -> Image.Image:
    """Полноэкранный снимок с текстом — проверка на большом разрешении."""
    image = Image.new("RGB", size, "white")
    draw = ImageDraw.Draw(image)
    font = _font(font_size)
    for i in range(lines):
        draw.text((60, 60 + i * int(font_size * 1.8)),
                  "Строка текста Some english text 123", fill="black", font=font)
    return image


def test_text_presence() -> None:
    """Снимки с текстом обрабатываются, снимки без текста — пропускаются."""
    samples: list[tuple[str, Image.Image, bool]] = [
        ("три строки текста", text_image(["Привет, мир!", "Это тестовая строка", "Третья строка"]), True),
        ("одно слово", text_image(["Настройки"]), True),
        ("два слова", text_image(["Настройки", "Интерфейс"]), True),
        ("мелкий текст", text_image(["Мелкий текст на строке"], size=(400, 120), font_size=14), True),
        ("короткое «OK»", two_letters_image(), True),
        ("однотонная заливка", solid_image(), False),
        ("градиент", gradient_image(), False),
        ("шум", noise_image(), False),
        ("фотография", photo_like_image(), False),
        ("зернистая фотография", grain_photo_image(), False),
        ("фото с прямоугольниками", photo_with_rects_image(), False),
        ("иконки и фигуры", icon_like_image(), False),
        ("интерфейс без надписей", ui_mock_image(), False),
    ]
    print("\nПроверка «есть ли текст» (оценка / уверенно / деталей / итог / ожидание):")
    for name, image, expected in samples:
        score, confident, glyphs = app.analyse_text_presence(image)
        verdict = app.image_has_text(image)
        print(f"  {name:24} score={score:5.0f} уверенно={str(confident):5} деталей={glyphs:5d} "
              f"итог={str(verdict):5} ожидалось={expected}")
        check(f"фильтр: {name} → {expected}", verdict == expected, f"получено {verdict}")

    # Слишком маленький снимок не считаем текстом.
    tiny = Image.new("RGB", (4, 4), "white")
    check("фильтр: снимок 4x4 без текста", not app.image_has_text(tiny))

    # Большой снимок с текстом: проверка работает и на 4K.
    check("фильтр: текст на 4K", app.image_has_text(big_text_image()))

    # Реальная картинка из репозитория: не должна падать и должна считаться текстом.
    sample = ROOT / "debug_deskewed.png"
    if sample.is_file():
        image = Image.open(sample)
        score, confident, glyphs = app.analyse_text_presence(image)
        print(f"  {sample.name:24} score={score:5.0f} уверенно={confident} деталей={glyphs}")
        check("фильтр: реальный скриншот — текст", app.image_has_text(image))


def test_real_photos() -> None:
    """Реальные фотографии без текста (обои Windows), если они есть на машине."""
    photos = sorted(glob.glob(r"C:\Windows\Web\Wallpaper\**\*.jpg", recursive=True))[:8]
    if not photos:
        print("\nРеальные фотографии не найдены — проверка пропущена.")
        return
    print("\nРеальные фотографии (ожидается «нет текста»):")
    positives = 0
    for path in photos:
        image = Image.open(path).convert("RGB")
        verdict = app.image_has_text(image)
        positives += 1 if verdict else 0
        print(f"  {Path(path).name:20} текст={verdict}")
    check("фильтр: фотографии в основном без текста", positives <= 2,
          f"текстом признано {positives} из {len(photos)}")


def test_split_blocks() -> None:
    cases = [
        ("ПЕРЕВОД:\nпривет\nОРИГИНАЛ:\nhello", ("привет", "hello")),
        ("ОРИГИНАЛ:\nhello\nПЕРЕВОД:\nпривет", ("привет", "hello")),
        ("перевод: привет\nоригинал: hello", ("привет", "hello")),
        ("просто перевод", ("просто перевод", "")),
        ("ПЕРЕВОД:\nтолько перевод", ("только перевод", "")),
        ("ОРИГИНАЛ:\nтолько оригинал", ("", "только оригинал")),
        ("", ("", "")),
    ]
    for raw, expected in cases:
        got = app._split_original_vs_translation(raw)
        check(f"блоки: {raw[:24]!r}", got == expected, f"-> {got!r} != {expected!r}")


def test_no_text_answers() -> None:
    for value in ("NO_TEXT", "no text", "No_Text", "НЕТ ТЕКСТА", "", "   "):
        check(f"«нет текста»: {value!r}", app._is_no_text_answer(value))
    for value in ("Привет", "Hello world", "NO_TEXT и перевод"):
        check(f"есть текст: {value!r}", not app._is_no_text_answer(value))


def test_stream_visible_text() -> None:
    steps = [
        ("ПЕРЕ", ""),
        ("ПЕРЕВОД:\nПривет", "Привет"),
        ("ПЕРЕВОД:\nПривет, мир", "Привет, мир"),
        ("ПЕРЕВОД:\nПривет, мир\nОРИГ", "Привет, мир"),
        ("NO_T", ""),
        ("NO_TEXT", ""),
        ("обычный текст без маркеров", "обычный текст без маркеров"),
    ]
    for raw, expected in steps:
        got = app._stream_visible_text(raw)
        check(f"поток: {raw[:22]!r}", got == expected, f"-> {got!r} != {expected!r}")


class FakeResponse:
    """Минимальная замена requests.Response для SSE."""

    def __init__(self, lines: list[bytes]) -> None:
        self._lines = lines

    def iter_lines(self, chunk_size: int = 0, decode_unicode: bool = False):
        yield from self._lines


def test_stream_completion() -> None:
    def chunk(text: str) -> bytes:
        return ("data: " + __import__("json").dumps(
            {"choices": [{"delta": {"content": text}}]})).encode("utf-8")

    lines = [chunk("При"), chunk("вет"), b"", b"data: [DONE]", b"data: {broken"]
    deltas: list[str] = []
    text = app._stream_completion(FakeResponse(lines), job=None, on_delta=deltas.append)
    check("SSE: собранный текст", text == "Привет", repr(text))
    check("SSE: промежуточные куски", deltas == ["При", "Привет"], repr(deltas))

    # Отмена внутри потока прерывает чтение.
    cancelled_job = app._new_job("hotkey")
    cancelled_job.cancel.set()
    try:
        app._stream_completion(FakeResponse([chunk("x")]), job=cancelled_job, on_delta=None)
        check("SSE: отмена прерывает чтение", False)
    except app.Cancelled:
        check("SSE: отмена прерывает чтение", True)
    finally:
        app._finish_job(cancelled_job)


def test_jobs() -> None:
    first = app._new_job("hotkey")
    check("задачи: первая активна", app._is_active_job(first.gen))
    second = app._new_job("live")
    check("задачи: новый захват отменяет предыдущий", first.cancelled())
    check("задачи: новая задача не отменена", not second.cancelled())
    check("задачи: результат старой не показывается", not app._is_active_job(first.gen))
    check("задачи: результат новой актуален", app._is_active_job(second.gen))
    check("задачи: отмена текущей", app.cancel_current_job("тест"))
    check("задачи: повторная отмена не срабатывает", not app.cancel_current_job("тест"))
    app._finish_job(second)
    check("задачи: после завершения активных нет", app._current_job() is None)
    check("задачи: нечего отменять", not app.cancel_current_job("тест"))


def test_ocr_data_grouping() -> None:
    data = {
        "text": ["Привет,", "мир", "", "42", "x"],
        "conf": ["92", "88", "-1", "77", "15"],
        "block_num": [1, 1, 1, 1, 2],
        "par_num": [1, 1, 1, 1, 1],
        "line_num": [1, 1, 1, 2, 1],
    }
    result = app._ocr_result_from_data(data)
    check("OCR: строки собраны", result.text == "Привет, мир\n42", repr(result.text))
    check("OCR: слова посчитаны", result.words == 4, repr(result.words))
    check("OCR: уверенность усреднена", 60 <= result.confidence <= 95, repr(result.confidence))
    check("OCR: текст признан текстом", result.looks_like_text)

    junk = app._ocr_result_from_data({
        "text": ["~", "|"], "conf": ["12", "5"],
        "block_num": [1, 1], "par_num": [1, 1], "line_num": [1, 1],
    })
    check("OCR: мусор не считается текстом", not junk.looks_like_text, repr(junk))


def test_api_key_sources() -> None:
    """Ключ API: порядок поиска и понятное сообщение при 401."""
    check("ключ: несуществующая переменная пуста",
          app.env_var("НЕТ_ТАКОЙ_ПЕРЕМЕННОЙ_12345") == "")
    check("ключ: пустое имя переменной", app.env_var("") == "")

    original_env = os.environ.get("PAIT_TEST_KEY")
    try:
        os.environ["PAIT_TEST_KEY"] = "  test-key-123  "
        check("ключ: берётся из окружения процесса", app.env_var("PAIT_TEST_KEY") == "test-key-123")
    finally:
        if original_env is None:
            os.environ.pop("PAIT_TEST_KEY", None)
        else:
            os.environ["PAIT_TEST_KEY"] = original_env

    section = app.get_provider("deepseek")
    original_key = (section or {}).get("api_key")
    try:
        section["api_key"] = "key-from-settings"
        check("ключ: настройки приоритетнее окружения",
              app.get_api_key("deepseek") == "key-from-settings")
        check("ключ: источник — настройки",
              app.key_source_hint("deepseek") == "из настроек (config.json)")
    finally:
        section["api_key"] = original_key

    check("ключ: источник для deepseek определён",
          app.key_source_hint("deepseek") in (
              "из переменной окружения DEEPSEEK_API_KEY",
              "НЕ НАЙДЕН (используется плейсхолдер — перевода не будет)"),
          app.key_source_hint("deepseek"))

    class FakeResponse:
        status_code = 401
        text = '{"error": "invalid key"}'

        def raise_for_status(self) -> None:
            raise AssertionError("не должно вызываться для 401")

    try:
        app._check_response(FakeResponse(), "model", provider_id="deepseek", api_key="")
        check("401 без ключа: понятное сообщение", False, "исключения не было")
    except RuntimeError as exc:
        check("401 без ключа: понятное сообщение", "не найден" in str(exc), str(exc))

    try:
        app._check_response(FakeResponse(), "model", provider_id="deepseek", api_key="sk-real")
        check("401 с ключом: обычное сообщение", False, "исключения не было")
    except RuntimeError as exc:
        check("401 с ключом: обычное сообщение", "401" in str(exc), str(exc))


def test_capture_mode_helpers() -> None:
    """Режим выделения области и действующая клавиша захвата."""
    default_mode = app.capture_mode_of({})
    check("режим: значение по умолчанию", default_mode in ("builtin", "system"), default_mode)
    check("режим: явный system", app.capture_mode_of({"capture_mode": "system"}) == "system")
    check("режим: явный builtin", app.capture_mode_of({"capture_mode": "builtin"}) == "builtin")
    check("режим: мусор → по умолчанию",
          app.capture_mode_of({"capture_mode": "???"}) == default_mode)

    check("подсказка: своё выделение не занимает Win+Shift+S",
          app.hotkey_hint("builtin", "windows+shift+s").endswith("занята системой"),
          app.hotkey_hint("builtin", "windows+shift+s"))
    check("подсказка: системный режим оставляет Win+Shift+S",
          app.hotkey_hint("system", "windows+shift+s") == "клавиша: windows+shift+s")
    check("подсказка: своя клавиша не подменяется",
          app.hotkey_hint("builtin", "ctrl+alt+q") == "клавиша: ctrl+alt+q")

    original = (app.CONFIG.get("capture_mode"), app.CONFIG.get("hotkey"))
    try:
        app.CONFIG["capture_mode"] = "builtin"
        app.CONFIG["hotkey"] = "windows+shift+s"
        check("клавиша: в своём режиме Win+Shift+S заменяется",
              app.effective_hotkey() == app._default_builtin_hotkey(),
              app.effective_hotkey())
        app.CONFIG["capture_mode"] = "system"
        check("клавиша: в системном режиме остаётся из настроек",
              app.effective_hotkey() == "windows+shift+s")
        app.CONFIG["capture_mode"] = "builtin"
        app.CONFIG["hotkey"] = "ctrl+alt+q"
        check("клавиша: своё сочетание не подменяется", app.effective_hotkey() == "ctrl+alt+q")
    finally:
        app.CONFIG["capture_mode"], app.CONFIG["hotkey"] = original


def test_autostart_helpers() -> None:
    command = app._autostart_command()
    check("автозапуск: команда не пустая", bool(command.strip()), command)
    check("автозапуск: файлы команды существуют", app.autostart_command_is_valid(command), command)
    check("автозапуск: несуществующий путь — невалиден",
          not app.autostart_command_is_valid(r'"C:\NoSuchFolder\pyAutoImgTranslate.exe"'))
    check("автозапуск: пустая команда — невалидна", not app.autostart_command_is_valid(""))
    check("автозапуск: устаревание не падает", isinstance(app.autostart_is_stale(), bool))


def test_analysis_parsing() -> None:
    """Разбор ответа модели для анализа: описание, запросы, источники."""
    cases = [
        ("Кот на клавиатуре.\n\nПОИСК:\n- кот на клавиатуре\n- почему коты лежат на клавиатуре",
         ("Кот на клавиатуре.", ["кот на клавиатуре", "почему коты лежат на клавиатуре"])),
        ("Кот на клавиатуре. ПОИСК: - кот на клавиатуре", ("Кот на клавиатуре.", ["кот на клавиатуре"])),
        ("Просто описание без поиска.", ("Просто описание без поиска.", [])),
        ("Описание...\nПОИСК: НЕТ", ("Описание...", [])),
        ("Описание\nПОИСК:\n1. товар\n2) цена\n• отзывы\n- отзывы",
         ("Описание", ["товар", "цена", "отзывы"])),
        ("", ("", [])),
    ]
    for raw, expected in cases:
        got = app.split_description_and_queries(raw)
        check(f"анализ: деление {raw[:24]!r}", got == expected, f"-> {got!r} != {expected!r}")

    _, queries = app.split_description_and_queries(
        "Описание\nПОИСК:\n- раз\n- два\n- три\n- четыре")
    check("анализ: не больше трёх запросов", len(queries) == 3, repr(queries))

    check("анализ: незавершённый маркер не ломает показ",
          app.visible_analysis_text("Кот на клавиатуре. ПОИ") == "Кот на клавиатуре. ПОИ")
    check("анализ: блок поиска отрезан",
          app.visible_analysis_text("Кот.\nПОИСК:\n- кот") == "Кот.")
    check("анализ: inline-блок отрезан",
          app.visible_analysis_text("Кот на клавиатуре. ПОИСК:\n- кот") == "Кот на клавиатуре.")

    check("анализ: компоновка с ответом",
          app.compose_analysis_text("Кот.", "Это животное.") == "Кот.\n\nПодробнее:\nЭто животное.")
    check("анализ: компоновка без ответа", app.compose_analysis_text("Кот.", "") == "Кот.")

    sources = [
        {"title": "A &amp; B", "url": "https://a.b/c"},
        {"title": "C", "url": "https://d.e/x?y=1"},
    ]
    formatted = app.format_sources(sources)
    check("анализ: источники отформатированы",
          formatted.startswith("1. A & B\n   https://a.b/c"), repr(formatted))

    merged = app.merge_search_results(
        [[{"url": "https://a/#frag", "title": "A"}],
         [{"url": "https://a/%23frag", "title": "A2"}, {"url": "https://b", "title": "B"}],
         [{"url": "https://b/", "title": "B2"}]])
    check("анализ: повторы источников убраны", [item["url"] for item in merged] ==
          ["https://a/#frag", "https://b"], repr(merged))

    check("анализ: мусорный результат отброшен",
          not app.result_is_relevant("Sony WH-1000XM5 цена",
                                     {"title": "Протаргол — цена", "snippet": "инструкция",
                                      "url": "https://medum.ru"}))
    check("анализ: результат по теме оставлен",
          app.result_is_relevant("Илон Маск",
                                 {"title": "Илон Маск — биография", "snippet": "Илон Маск",
                                  "url": "https://rbc.ru"}))
    check("анализ: длинное слово запроса достаточно",
          app.result_is_relevant("наушники Sony", {"title": "Наушники купить", "snippet": "",
                                                   "url": "https://ozon.ru"}))

    rss = ("<?xml version='1.0'?><rss><channel>"
           "<item><title>Заголовок &amp; ещё</title><link>https://site/1</link>"
           "<description>&lt;b&gt;Сниппет&lt;/b&gt; текста</description></item>"
           "<item><title>Второй</title><link>https://site/2</link>"
           "<description>Текст</description></item>"
           "</channel></rss>")
    parsed = app.parse_bing_rss(rss)
    check("поиск: RSS разобран", len(parsed) == 2, repr(parsed))
    check("поиск: сущности и теги убраны",
          bool(parsed) and parsed[0]["title"] == "Заголовок & ещё"
          and parsed[0]["snippet"] == "Сниппет текста", repr(parsed[0] if parsed else None))

    ddg = ('<a class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fsite%2Fx">'
           'Заголовок</a><a class="result__snippet">Описание</a>')
    parsed_ddg = app.parse_ddg_html(ddg)
    check("поиск: DDG разобран",
          bool(parsed_ddg) and parsed_ddg[0]["url"] == "https://site/x", repr(parsed_ddg))

    vision = app._build_analysis_messages(Image.new("RGB", (10, 10), "white"))
    has_image = any(isinstance(part, dict) and part.get("type") == "image_url"
                    for message in vision for part in
                    (message["content"] if isinstance(message["content"], list) else []))
    check("анализ: в промпте есть картинка", has_image)
    check("анализ: в промпте просят блок ПОИСК", "ПОИСК" in vision[0]["content"])

    with_sources = app._build_search_answer_messages("Описание", sources)
    no_sources = app._build_search_answer_messages("Описание", [])
    check("анализ: источники попали в запрос", "https://a.b/c" in with_sources[1]["content"])
    check("анализ: без источников просят быть честным",
          "не нашлось" in no_sources[1]["content"])


def test_search_layer() -> None:
    """Слой поиска: кэш, объединение источников, пустой запрос."""
    check("поиск: пустой запрос", app.web_search("") == [])

    original_wiki = app.wikipedia_search
    original_bing = app.bing_search
    app._SEARCH_CACHE.clear()
    calls = {"wiki": 0, "bing": 0}

    def fake_wiki(query: str, limit: int = 2):
        calls["wiki"] += 1
        return [{"title": f"{query} — Википедия", "url": f"https://wiki/{query}",
                 "snippet": "кратко"}]

    def fake_bing(query: str, limit: int = 4):
        calls["bing"] += 1
        return [{"title": f"{query} — сайт", "url": f"https://site/{query}",
                 "snippet": "описание"}]

    app.wikipedia_search = fake_wiki
    app.bing_search = fake_bing
    try:
        first = app.web_search("тестовый запрос", 4)
        second = app.web_search("тестовый запрос", 4)
        check("поиск: источники собраны", len(first) == 2, repr(first))
        check("поиск: повторный вызов из кэша",
              second == first and calls["wiki"] == 1 and calls["bing"] == 1, repr(calls))
        app.web_search("", 4)
        check("поиск: пустой запрос не ходит в сеть", calls["wiki"] == 1, repr(calls))
    finally:
        app.wikipedia_search = original_wiki
        app.bing_search = original_bing
        app._SEARCH_CACHE.clear()


def main() -> int:
    print("Офлайн-проверки pyAutoImgTranslate")
    for test in (test_text_presence, test_real_photos, test_split_blocks, test_no_text_answers,
                 test_stream_visible_text, test_stream_completion, test_jobs,
                 test_ocr_data_grouping, test_api_key_sources, test_capture_mode_helpers,
                 test_analysis_parsing, test_search_layer, test_autostart_helpers):
        print(f"\n== {test.__name__}")
        test()
    print("\n".join(REPORT))
    print(f"\nИтого: успешно {len(REPORT) - len(FAILURES)}, провалов {len(FAILURES)}")
    for failure in FAILURES:
        print(failure)
    return 1 if FAILURES else 0


if __name__ == "__main__":
    raise SystemExit(main())
