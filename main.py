import sys
import os
import time
import threading
import struct
import io

import tkinter as tk
from tkinter import scrolledtext

import keyboard
import mouse
import requests
import pytesseract
import win32clipboard
import win32con
import pystray
import numpy as np
import cv2

from PIL import Image, ImageGrab, ImageTk, ImageDraw

# Если tesseract установлен не в PATH, укажите путь явно:
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Глобальная переменная для окна результата
result_window = None

# Цветовая палитра (пример, можно изменить при необходимости)
PRIMARY_COLOR = "#F54B64"
PRIMARY_COLOR_2 = "#F78361"
SECONDARY_COLOR = "#FFD42B"
DARK_GREY = "#4E586E"
WHITE = "#FFFFFF"


# --------------------- Улучшенная функция deskew_image ---------------------
def deskew_image(pil_image):
    """
    Принимает PIL-изображение, инвертирует его, определяет угол поворота
    и выравнивает (deskew) изображение с использованием OpenCV.
    Возвращает скорректированное PIL-изображение без обрезки текста.
    """
    # Конвертируем PIL -> OpenCV (BGR)
    cv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)

    # Переводим в оттенки серого и инвертируем, чтобы текст стал ярким на темном фоне
    gray = cv2.cvtColor(cv_image, cv2.COLOR_BGR2GRAY)
    gray = cv2.bitwise_not(gray)

    # Бинаризация с использованием порога Оцу
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

    # Находим координаты всех белых пикселей
    coords = np.column_stack(np.where(thresh > 0))
    if len(coords) == 0:
        return pil_image

    # Определяем минимальный прямоугольник, окружающий текст
    angle = cv2.minAreaRect(coords)[-1]

    # Корректируем угол (minAreaRect возвращает угол в диапазоне [-90, 0))
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle

    # Вычисляем новые размеры, чтобы избежать обрезки после поворота
    (h, w) = cv_image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    cos = np.abs(M[0, 0])
    sin = np.abs(M[0, 1])
    nW = int((h * sin) + (w * cos))
    nH = int((h * cos) + (w * sin))
    # Корректируем матрицу поворота с учетом сдвига
    M[0, 2] += (nW / 2) - center[0]
    M[1, 2] += (nH / 2) - center[1]

    rotated = cv2.warpAffine(cv_image, M, (nW, nH), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    rotated_pil = Image.fromarray(cv2.cvtColor(rotated, cv2.COLOR_BGR2RGB))
    return rotated_pil


# --------------------- Функции для работы с буфером обмена ---------------------
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


# --------------------- OCR и перевод ---------------------
def ocr_image(image):
    """
    Выпрямляет изображение (deskew) и извлекает текст с помощью pytesseract.
    """
    deskewed = deskew_image(image)
    #deskewed.save("debug_deskewed.png")
    config = r"--psm 6 --oem 3"
    upscaled = upscale_image(deskewed, scale=2)
    text = pytesseract.image_to_string(upscaled, lang="eng", config=config)
    return text

def upscale_image(pil_image, scale=2):
    cv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    # Увеличиваем в 2 раза
    upscaled = cv2.resize(cv_image, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    return Image.fromarray(cv2.cvtColor(upscaled, cv2.COLOR_BGR2RGB))

def send_text_for_translation(text):
    """
    Отправляет текст на локальный сервер LM Studio по эндпоинту /v1/chat/completions.
    Использует системное и пользовательское сообщения для запроса чистого перевода.
    """
    url = "http://127.0.0.1:1234/v1/chat/completions"
    payload = {
        "model": "llama-translate",  # замените на имя вашей модели
        "messages": [
            {
                "role": "system",
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


# --------------------- Окно результата ---------------------
def show_result(image, translation):
    """
    Отображает окно с изображением и переводом.
    Интерфейс оформлен в плоском стиле с заданной цветовой гаммой.
    Обеспечивается масштабирование элементов, область с переводом имеет прокрутку,
    а пользователь может копировать выделенные фрагменты через контекстное меню.
    """
    global result_window
    if result_window is not None:
        try:
            result_window.destroy()
        except Exception:
            pass

    result_window = tk.Tk()
    result_window.title("Скриншот и перевод ИИ")
    result_window.attributes("-topmost", True)
    result_window.focus_force()
    result_window.bind("<Escape>", lambda e: result_window.destroy())

    # Применяем цветовую гамму
    result_window.configure(bg=DARK_GREY)

    # Настройка grid для масштабирования
    result_window.rowconfigure(0, weight=1)
    result_window.columnconfigure(0, weight=1)

    frame = tk.Frame(result_window, bg=DARK_GREY, bd=0, highlightthickness=0)
    frame.grid(sticky="nsew", padx=10, pady=10)
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)

    # Изображение слева
    img_tk = ImageTk.PhotoImage(image)
    label_img = tk.Label(frame, image=img_tk, bg=DARK_GREY, bd=0, highlightthickness=0, relief="flat")
    label_img.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
    label_img.image = img_tk

    # Текстовая область справа
    text_frame = tk.Frame(frame, bg=DARK_GREY, bd=0, highlightthickness=0)
    text_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
    text_frame.rowconfigure(0, weight=1)
    text_frame.columnconfigure(0, weight=1)

    scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL)
    scrollbar.grid(row=0, column=1, sticky="ns")

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

    text_widget.insert(tk.END, translation)
    text_widget.config(state="disabled")

    # Настройка цвета выделения текста
    text_widget.config(state="normal")
    text_widget.tag_configure("sel", background=PRIMARY_COLOR, foreground=WHITE)
    text_widget.config(state="disabled")

    # Контекстное меню для копирования выделенного фрагмента
    def copy_selected():
        try:
            text_widget.config(state="normal")
            selected = text_widget.get(tk.SEL_FIRST, tk.SEL_LAST)
            text_widget.config(state="disabled")
            result_window.clipboard_clear()
            result_window.clipboard_append(selected)
            print("Скопировано:", selected)
        except tk.TclError:
            pass

    context_menu = tk.Menu(text_widget, tearoff=0, bg=WHITE, fg=DARK_GREY)
    context_menu.add_command(label="Копировать", command=copy_selected)

    def show_context_menu(event):
        context_menu.tk_popup(event.x_root, event.y_root)

    text_widget.bind("<Button-3>", show_context_menu)

    result_window.mainloop()
    result_window = None


# --------------------- Горячая клавиша и логика ---------------------
process_pending = False


def wait_for_left_click(timeout=10):
    """
    Ждёт последовательность: нажатие (down) левой кнопки мыши, затем её отпускание (up).
    Если до завершения последовательности нажата любая другая клавиша – ожидание отменяется.
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


def on_hotkey():
    """
    Обработчик нажатия Win+Shift+S.
    После отпускания горячих клавиш предлагается выбрать область стандартным способом:
    нажать и отпустить левую кнопку мыши, затем ждётся, пока изображение появится в буфере обмена.
    После этого выполняется OCR (с deskew), отправка на LM Studio и вывод результата.
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


# --------------------- Системный трей (иконка и выход) ---------------------
def create_image_for_tray():
    """
    Пытается загрузить значок из файла translator.png.
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
    from pystray import Icon, Menu, MenuItem
    menu = Menu(MenuItem('Выход', on_exit))
    icon = Icon("pyAutoImgTranslate", create_image_for_tray(), "Скриншот и перевод ИИ", menu)
    icon.run()


if __name__ == "__main__":
    tray_thread = threading.Thread(target=start_tray_icon)
    tray_thread.daemon = True
    tray_thread.start()

    print("Ожидается нажатие Win+Shift+S для создания скриншота стандартным способом Windows...")
    keyboard.add_hotkey('windows+shift+s', on_hotkey)
    keyboard.wait()