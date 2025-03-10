import tkinter as tk
from tkinter import scrolledtext
from PIL import Image, ImageGrab, ImageTk, ImageDraw
import io
import requests
import keyboard
import mouse
import time
import pytesseract
import win32clipboard
import win32con
import struct
import threading
import pystray
import sys
import os

# Если tesseract установлен не в PATH, укажите путь явно:
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Цветовая палитра
PRIMARY_COLOR = "#F54B64"   # Основной акцент
PRIMARY_COLOR_2 = "#F78361" # Вторая часть градиента (при желании можно использовать в других местах)
SECONDARY_COLOR = "#FFD42B" # Второстепенный акцент
DARK_GREY = "#4E586E"       # Тёмно-серый для фона
WHITE = "#FFFFFF"           # Белый

# Глобальная переменная для окна результата
result_window = None

def get_image_from_clipboard():
    """
    Пытается получить изображение из буфера обмена.
    Сначала использует ImageGrab.grabclipboard(), а если не получается –
    обращается к данным формата CF_DIB, добавляет BMP-заголовок и создаёт объект Image.
    """
    image = ImageGrab.grabclipboard()
    if image and isinstance(image, Image.Image):
        return image

    try:
        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData(win32con.CF_DIB)
        except Exception as e:
            print("Ошибка получения CF_DIB:", e)
            data = None
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        print("Ошибка доступа к буферу обмена:", e)
        data = None

    if data is None:
        return None

    try:
        bmp_header = b'BM'
        size = len(data) + 14  # общий размер BMP-файла
        bmp_header += struct.pack("<I", size)
        bmp_header += b'\x00\x00'
        bmp_header += b'\x00\x00'
        bmp_header += b'\x36\x00\x00\x00'
        bmp_data = bmp_header + data
        stream = io.BytesIO(bmp_data)
        image = Image.open(stream)
        return image
    except Exception as e:
        print("Ошибка создания изображения из CF_DIB:", e)
        return None

def wait_for_clipboard_image(timeout=5):
    """
    Ожидает появления изображения в буфере обмена в течение timeout секунд.
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        image = get_image_from_clipboard()
        if image and isinstance(image, Image.Image):
            return image
        time.sleep(0.2)
    return None

def ocr_image(image):
    """
    Извлекает текст из изображения с помощью pytesseract.
    """
    text = pytesseract.image_to_string(image, lang="eng")
    return text

def send_text_for_translation(text):
    """
    Отправляет текст на локальный сервер LM Studio по эндпоинту /v1/chat/completions.
    Используются два сообщения:
      - Системное сообщение: инструкция, чтобы модель переводила текст на русский
        и возвращала только один краткий результат перевода без дополнительных строк, комментариев, повторов или форматирования.
      - Пользовательское сообщение: исходный текст для перевода.
    """
    url = "http://127.0.0.1:1234/v1/chat/completions"
    payload = {
        "model": "llama-translate",  # замените на имя вашей модели
        "messages": [
            {"role": "system",
             "content": (
                    "Переведи данный текст на русский язык. Верни только один краткий результат перевода, "
                    "без каких-либо дополнительных строк, комментариев, повторов или форматирования. "
                    "Ответ должен состоять только из перевода."
                )
             },
            {"role": "user", "content": text}
        ],
        "temperature": 0.0,
        "stop": ["###", "<|end_of_text|>"]
    }
    headers = {"Content-Type": "application/json"}
    try:
        response = requests.post(url, json=payload, headers=headers)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print("Ошибка при отправке текста:", e)
        return {}


def show_result(image, translation):
    """
    Окно с «плоским» оформлением и новой цветовой гаммой:
      - Фон окна и фреймов: DARK_GREY
      - Область перевода: белый фон (WHITE) с тёмно-серым текстом (DARK_GREY)
      - Прокрутка и рамки в «плоском» стиле
      - Возможность копировать любую часть текста через правый клик
    """
    global result_window

    # Закрываем предыдущее окно, если открыто
    if result_window is not None:
        try:
            result_window.destroy()
        except Exception:
            pass

    result_window = tk.Tk()
    result_window.title("Скриншот и перевод")

    # Ставим окно поверх
    result_window.attributes("-topmost", True)
    result_window.focus_force()

    # Закрытие по Esc
    result_window.bind("<Escape>", lambda e: result_window.destroy())

    # Фон окна — тёмно-серый
    result_window.configure(bg=DARK_GREY)

    # Настраиваем сетку для масштабирования
    result_window.rowconfigure(0, weight=1)
    result_window.columnconfigure(0, weight=1)

    # Создаём фрейм (фон — DARK_GREY, без объёмных рамок)
    frame = tk.Frame(result_window, bg=DARK_GREY, bd=0, highlightthickness=0)
    frame.grid(sticky="nsew", padx=10, pady=10)
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)  # Колонка для текста будет растягиваться

    # Изображение слева
    img_tk = ImageTk.PhotoImage(image)
    label_img = tk.Label(frame, image=img_tk, bg=DARK_GREY, bd=0, highlightthickness=0, relief="flat")
    label_img.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
    label_img.image = img_tk  # чтобы не удалилось сборщиком мусора

    # Фрейм для текстовой области с прокруткой
    text_frame = tk.Frame(frame, bg=DARK_GREY, bd=0, highlightthickness=0)
    text_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
    text_frame.rowconfigure(0, weight=1)
    text_frame.columnconfigure(0, weight=1)

    # Горизонтальная/вертикальная прокрутка (если нужно, можно добавить и xscrollbar)
    scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL)
    scrollbar.grid(row=0, column=1, sticky="ns")

    # Текстовый виджет — белый фон, тёмно-серый текст, «плоский» стиль
    text_widget = tk.Text(
        text_frame,
        wrap="word",
        yscrollcommand=scrollbar.set,
        bg=WHITE,
        fg=DARK_GREY,
        font=("Arial", 12),
        bd=0,
        highlightthickness=0
    )
    text_widget.grid(row=0, column=0, sticky="nsew")
    scrollbar.config(command=text_widget.yview)

    # Вставляем перевод и делаем виджет только для чтения
    text_widget.insert(tk.END, translation)
    text_widget.config(state="disabled")

    # Настраиваем цвет выделения текста (PRIMARY_COLOR)
    # Чтобы сработало, нужно временно разблокировать текст
    text_widget.config(state="normal")
    text_widget.tag_configure("sel", background=PRIMARY_COLOR, foreground=WHITE)
    text_widget.config(state="disabled")

    # Функция копирования выделенного текста
    def copy_selected():
        try:
            # Временно включаем виджет, чтобы получить выделенный фрагмент
            text_widget.config(state="normal")
            selected = text_widget.get(tk.SEL_FIRST, tk.SEL_LAST)
            text_widget.config(state="disabled")
            result_window.clipboard_clear()
            result_window.clipboard_append(selected)
            print("Скопировано:", selected)
        except tk.TclError:
            pass

    # Контекстное меню (правый клик) для копирования выделенного текста
    context_menu = tk.Menu(text_widget, tearoff=0, bg=WHITE, fg=DARK_GREY, bd=0)
    context_menu.add_command(label="Копировать", command=copy_selected)

    # Отображаем меню по правому клику
    def show_context_menu(event):
        context_menu.tk_popup(event.x_root, event.y_root)

    text_widget.bind("<Button-3>", show_context_menu)

    result_window.mainloop()
    result_window = None

def wait_for_left_click(timeout=10):
    """
    Ждёт последовательность: нажатие (down) левой кнопки мыши, затем её отпускание (up).
    Если до завершения последовательности нажата любая клавиша – ожидание отменяется.
    Возвращает True, если последовательность выполнена, иначе False.
    """
    left_pressed = False
    left_released = False
    cancelled = False

    def on_mouse_event(event):
        nonlocal left_pressed, left_released
        try:
            if event.event_type == 'down' and event.button == 'left':
                left_pressed = True
            elif event.event_type == 'up' and event.button == 'left' and left_pressed:
                left_released = True
        except AttributeError:
            pass

    def on_key_event(event):
        nonlocal cancelled
        cancelled = True

    mouse.hook(on_mouse_event)
    keyboard.hook(on_key_event)
    start_time = time.time()
    while time.time() - start_time < timeout:
        if left_pressed and left_released:
            break
        if cancelled:
            break
        time.sleep(0.05)
    mouse.unhook(on_mouse_event)
    keyboard.unhook(on_key_event)
    return left_pressed and left_released and not cancelled

process_pending = False

def on_hotkey():
    """
    Обработчик нажатия Win+Shift+S.
    После отпускания горячих клавиш предлагается выбрать область стандартным способом:
    нажать и отпустить левую кнопку мыши, затем ждётся, пока изображение появится в буфере обмена.
    После этого выполняется OCR, отправка на LM Studio и вывод результата.
    Если последовательность не завершена, никакого результата не показывается.
    """
    global process_pending
    if process_pending:
        return
    process_pending = True
    print("Нажата комбинация Win+Shift+S.")
    print("Пожалуйста, нажмите и, удерживая, отпустите левую кнопку мыши для выделения области...")

    if wait_for_left_click():
        print("Область выбрана. Ждём, пока Windows добавит скриншот в буфер обмена...")
        image = wait_for_clipboard_image(timeout=5)
        if image:
            print("Изображение обнаружено в буфере обмена. Выполняется OCR...")
            extracted_text = ocr_image(image)
            if not extracted_text.strip():
                print("На изображении не найден текст.")
                process_pending = False
                return
            print("Извлечённый текст:", extracted_text)
            print("Отправляю текст на перевод в LM Studio (/v1/chat/completions)...")
            result = send_text_for_translation(extracted_text)
            translation = result.get("choices", [{}])[0].get("message", {}).get("content", "Перевод не получен")
            print("Получен перевод:", translation)
            show_result(image, translation)
        else:
            print("В буфере обмена не появилось изображения.")
    else:
        print("Ожидание отменено – нажата другая клавиша.")
    process_pending = False

def create_image_for_tray():
    """
    Пытается загрузить значок из файла translator.ico.
    Если не удаётся, создается дефолтное изображение.
    """
    try:
        icon_image = Image.open("translator.png")
        return icon_image
    except Exception as e:
        print("Не удалось загрузить translator.png, используется дефолтный значок:", e)
        image = Image.new('RGB', (64, 64), color=(50, 100, 150))
        draw = ImageDraw.Draw(image)
        draw.rectangle((16, 16, 48, 48), fill=(200, 200, 0))
        return image

def on_exit(icon, item):
    icon.stop()
    keyboard.unhook_all()
    os._exit(0)

def start_tray_icon():
    menu = pystray.Menu(pystray.MenuItem('Выход', on_exit))
    icon = pystray.Icon("pyAutoImgTranslate", create_image_for_tray(), "Переводчик ИИ", menu)
    icon.run()

if __name__ == "__main__":
    tray_thread = threading.Thread(target=start_tray_icon)
    tray_thread.daemon = True
    tray_thread.start()
    print("Ожидается нажатие Win+Shift+S для создания скриншота стандартным способом Windows...")
    keyboard.add_hotkey('windows+shift+s', on_hotkey)
    keyboard.wait()