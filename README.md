pyAutoImgTranslate is a utility that automatically translates text from screenshots through OCR and AI.

## Main functions

- Screenshot of the area: Activated by pressing Win+Shift+S to select the desired area of the screen.
- OCR with rectification: Text recognition with automatic tilt correction (deskew) using Tesseract and OpenCV, with garbage-line cleanup (`eng+rus`).
- OCR/translation: By default the **neural network reads the image directly** (vision) and returns a faithful translation — no Tesseract needed. A local **Tesseract** mode is also available. Toggle with `ocr_mode` (`ai` / `tesseract`).
- Translation: Any number of AI providers over one OpenAI-compatible interface (DeepSeek cloud, local LM Studio, custom servers). Providers and their models are added/edited/removed live from the settings window — nothing is hardcoded. The prompt asks for a **faithful, close-to-original** translation that preserves structure.
- Live mode ("на лету"): when enabled, any new image that lands in the clipboard is OCR'd and translated automatically — no need to press the hotkey first. Toggle from the tray or with `Win+Shift+A`.
- Progress window: a small always-on-top "hourglass" panel with a stage label, a smoothly animated percentage (0→100%), a progress bar, and an **ETA** (elapsed + estimated time remaining), shown while capturing, recognizing and translating.
- Single settings window: opening settings from the tray or from the result window never creates a duplicate — it raises the existing one.
- Balance display: for cloud providers (DeepSeek) the settings window shows the current API balance (e.g. "Баланс API: 12.34 USD") with a refresh button.
- User-friendly interface: a compact result window with the screenshot on the left and a tidy, scrollable translation block on the right; switch between **Перевод / Оригинал** (`Ctrl+O`) and copy everything (`Ctrl+C`) or a selection.
- Autostart: Optional launch together with Windows, toggled from the tray menu or the Settings window (stored in `HKCU\...\CurrentVersion\Run`).
- Working in the background: The application runs in the system tray; single-instance guard prevents duplicates.
- Logging: Rotating log file `pyAutoImgTranslate.log` next to the executable.

## Technologies used

- Python
- Tkinter – graphical interface
- Tesseract OCR – text recognition
- OpenCV & NumPy – image tilt correction
- DeepSeek API and LM Studio API translation
- Pystray – icon in the system tray
- Keyboard & Mouse – global keyboard shortcuts

## Design system

The UI is built on a small, explicit **design system** defined in `main.py` — no ad-hoc
colors or spacings:

- `Color` — semantic color roles (`PRIMARY`, `SURFACE`, `TEXT`, `BORDER`, `SUCCESS`, …).
- `Space` — spacing scale, multiples of 4 (`XS`=4, `SM`=8, `MD`=12, `LG`=16, `XL`=24, `XXL`=32).
- `Font` — typography scale (`CAPTION`, `BODY`, `TITLE`, `MONO`).
- Reusable components: `_button(...)` (variants `primary` / `secondary` / `ghost` / `danger`),
  `_styled_entry(...)`, `_segmented(...)`.

All windows use only these tokens, keeping spacing, color and typography consistent.

## Platforms

Works on **Windows** and **Linux (Debian/Ubuntu and other distros)**.

| Feature | Windows | Linux |
| --- | --- | --- |
| Capture area | `Win+Shift+S` (native snipping overlay) | `Ctrl+Alt+S` (calls grim/gnome-screenshot/spectacle/scrot/maim/flameshot) |
| Clipboard image | WinAPI / CF_DIB | `xclip` / `xsel` / `wl-paste` |
| Global hotkeys | `keyboard` | `pynput` (X11) |
| Autostart | Registry `HKCU\...\Run` | `~/.config/autostart/*.desktop` |
| Tray | pystray (win32) | pystray (xorg/appindicator) |

On Debian/Ubuntu install the basics:

```bash
sudo apt-get update
sudo apt-get install -y python3 python3-pip python3-venv python3-tk \
    libgl1 libglib2.0-0 tesseract-ocr tesseract-ocr-eng tesseract-ocr-rus \
    xclip
# optional screenshot backends: gnome-screenshot / spectacle / scrot / maim / flameshot / grim+slurp
```

## Building a release

```bash
# Linux
./build.sh            # -> dist/pyAutoImgTranslate

# Windows
build.bat             # -> dist\pyAutoImgTranslate.exe
```

## Configuration

Everything is **data-driven**: there is no hardcoded list of providers or models. The
provider catalog lives in the `providers` array of `config.json` and is fully editable at
runtime from the tray → **Настройки…** window (also reachable via the **Настройки**
button in the result window).

The settings window lets you:

- **Add / remove / rename providers** (any number, each with its own name, ID, base URL,
  API key and optional key env-variable).
- **Add / remove models** per provider (`+` / `−`).
- **Pick the active provider and model** and switch between them.
- **Сделать активным** — make the selected provider the active one instantly.
- **Сохранить** — saves and **reconnects on the fly**, no restart required.
- **Перезагрузить из файла** — re-reads `config.json` from disk and applies it live.
- **Проверить** — sends a short test request and shows the result/error inline.
- **Открыть config.json** — saves and opens the config file for manual editing.

All windows run inside a single dedicated Tk thread, so opening settings from the tray
or the result window is safe (no second `Tk()` root).

### Config schema

| Option | Description | Default |
| --- | --- | --- |
| `active_provider` | ID of the selected provider | `deepseek` |
| `active_model` | Selected model for the active provider | `deepseek-chat` |
| `providers[]` | Array of providers (see below) | DeepSeek / LM Studio / custom |
| `ocr_mode` | `ai` (vision через нейросеть) или `tesseract` (локальный OCR) | `ai` |
| `ocr_lang` | Tesseract languages (только для `tesseract`) | `eng+rus` |
| `tesseract_cmd` | Path to `tesseract.exe` | `C:\Program Files\Tesseract-OCR\tesseract.exe` |
| `source_lang` | Source language hint | `авто` |
| `target_lang` | Translation target language | `русский` |
| `request_timeout` | HTTP timeout, seconds | `120` |
| `add_to_autostart` | Launch with Windows | `false` |
| `keep_image_on_top` | Result window always on top | `false` |
| `hotkey` | Global hotkey: capture area → OCR → translate | `windows+shift+s` |
| `live_hotkey` | Toggle live mode | `windows+shift+a` |

Each entry in `providers[]` looks like:

```json
{
  "id": "deepseek",
  "name": "DeepSeek (облако)",
  "base_url": "https://api.deepseek.com",
  "api_key": "",
  "api_key_env": "DEEPSEEK_API_KEY",
  "models": ["deepseek-chat", "deepseek-reasoner"]
}
```

The endpoint is always `<base_url>/chat/completions` (OpenAI-compatible). Add as many
providers as you like — the app picks them up without a code change.

> Old-style configs (`provider` + `deepseek`/`local`/`custom` blocks) are **migrated
automatically** on load to the new `providers[]` format.

### API key resolution

For each provider the key is resolved in this order:

1. `api_key` field of that provider in `config.json`
2. The environment variable named in `api_key_env`
3. For the `deepseek` provider only: `DEEPSEEK_API_KEY`, then the in-code default
   (`DEFAULT_DEEPSEEK_API_KEY` in `main.py`) — **replace it with your own key**.

The key is sent as `Authorization: Bearer <key>`. Never commit a real key to a public repository.

---

The project is designed for fast and effective translation of texts from the screen without unnecessary actions.
