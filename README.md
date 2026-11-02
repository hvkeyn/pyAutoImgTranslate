pyAutoImgTranslate is a utility that automatically translates text from screenshots through OCR and AI — and analyses a screenshot (describes it, looks things up on the web) on a separate hotkey.

## Main functions

- Screenshot of the area: pressing the hotkey freezes the screen and shows a full-screen overlay where you drag the region — the built-in selector, just like Win+Shift+S (all monitors, `Esc` / right click cancels, the size is shown while dragging). The system capture (Windows Snipping Tool / `grim`, `scrot`, … on Linux) stays available as the `system` mode and is used automatically if the screen cannot be grabbed. Mode: `capture_mode` (`builtin` / `system`).
- **Image analysis (`Win+Shift+I`)**: the same area selector, but the capture is **described** by the vision model, and when the frame shows something worth checking (a product, model, place, person, game, app, error, event) the app **searches the web** and answers from the found sources. The answer window has **Ответ** and **Источники** tabs, and the links are clickable. Toggle the search with `analyze_web_search`.
- OCR with rectification: Text recognition with automatic tilt correction (deskew) using Tesseract and OpenCV, with garbage-line cleanup (`eng+rus`).
- OCR/translation: By default the **neural network reads the image directly** (vision) and returns a faithful translation — no Tesseract needed. A local **Tesseract** mode is also available. Toggle with `ocr_mode` (`ai` / `tesseract`).
- No-text guard: before any API request the capture is checked locally, but it **only skips obviously empty captures** (solid fills, gradients, empty UI blocks, images without detail). A structural check looks for text-like lines; when it finds none but the capture has plenty of detail (text on a game or photo background), the capture is still sent — a missed translation is much worse than one extra request. Genuinely small/faint captures are double-checked with Tesseract over several image preparations (as-is, local contrast, Otsu binarization). When the image really has no text, the model answers `NO_TEXT` and no window appears. Toggle with `skip_no_text`.
- Cancellable: an **Отмена** button in the progress plate, an **Отменить** button in the result window and **Отменить обработку** in the tray menu all abort the current job; the HTTP connection is closed immediately, so there is no wait for the end of generation. Pressing the capture hotkey again cancels the previous job and starts a new capture (re-shoot).
- Streaming translation: the answer is read as a stream (SSE) and the text grows in the result window as it is generated; in vision mode the `ПЕРЕВОД` block comes first, so the translation is what you see, while the original stays on the **Оригинал** tab. Toggle with `stream_results`.
- Translation: Any number of AI providers over one OpenAI-compatible interface (DeepSeek cloud, local LM Studio, custom servers). Providers and their models are added/edited/removed live from the settings window — nothing is hardcoded. The prompt asks for a **faithful, close-to-original** translation that preserves structure.
- Live mode ("на лету"): when enabled, any new image that lands in the clipboard is OCR'd and translated automatically — no need to press the hotkey first. A new capture immediately supersedes the one being processed; identical images are not reprocessed. Toggle from the tray or with `Win+Shift+A`.
- Progress window: a small always-on-top "hourglass" panel with a stage label, a smoothly animated percentage (0→100%), a progress bar, an **ETA** (elapsed + estimated time remaining) and a **Cancel** button, shown while capturing, recognizing and translating.
- Toast notifications: messages ("no text on the capture", "processing cancelled", errors) appear as a small non-blocking plate in the corner, never steal focus and disappear by themselves.
- Single settings window: opening settings from the tray or from the result window never creates a duplicate — it raises the existing one.
- Balance display: for cloud providers (DeepSeek) the settings window shows the current API balance (e.g. "Баланс API: 12.34 USD") with a refresh button.
- User-friendly interface: a compact result window with the screenshot on the left and a tidy, scrollable translation block on the right; switch between **Перевод / Оригинал** (`Ctrl+O`) and copy everything (`Ctrl+C`) or a selection. Sizes and position are computed from the work area of the monitor under the cursor, and on narrow screens the image and the text stack vertically.
- Autostart: Optional launch together with Windows, toggled from the tray menu or the Settings window (stored in `HKCU\...\CurrentVersion\Run`). A stale entry (the app moved) is rewritten automatically at startup.
- Working in the background: The application runs in the system tray; single-instance guard prevents duplicates.
- Logging: Rotating log file `pyAutoImgTranslate.log` next to the executable. Recognized screen text is **not** written to it by default (`log_recognized_text`); the log records stages, sizes and errors, and never the API key.
- Speed: the built-in selector opens in ~0.15 s (screen grab) with no full-screen image copied into the UI, so dragging the frame costs well under a millisecond per mouse move; captures do not accumulate memory.

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
| Built-in area selector | yes (full-screen frozen overlay, `Win+Shift+X` by default) | yes on X11 (falls back to system tools if the screen cannot be grabbed) |
| Capture area (system mode) | `Win+Shift+S` (native snipping overlay) | `Ctrl+Alt+S` (calls grim/gnome-screenshot/spectacle/scrot/maim/flameshot) |
| Clipboard image | WinAPI / CF_DIB | `xclip` / `xsel` / `wl-paste` |
| Global hotkeys | `keyboard` | `pynput` (X11) |
| Autostart | Registry `HKCU\...\Run` | `~/.config/autostart/*.desktop` |
| Tray | pystray (win32) | pystray (xorg/appindicator) |

> `Win+Shift+S` belongs to Windows itself — it cannot be re-bound, and pressing it
> opens the system Snipping Tool. In `builtin` mode the app therefore uses
> `Win+Shift+X` (or `Ctrl+Alt+X` on Linux); set your own combination in the
> `hotkey` field of `config.json`.

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
./build_deb.sh 1.8.1  # -> dist/pyautoimgtranslate_1.8.1_<arch>.deb

# Windows
build.bat             # -> dist\pyAutoImgTranslate.exe
```

### Install the .deb (Debian/Ubuntu)

```bash
sudo apt install ./pyautoimgtranslate_1.8.1_amd64.deb
# then run:  pyautoimgtranslate   (or launch from the applications menu)
```

The `.deb` installs the binary to `/usr/lib/pyautoimgtranslate/`, a launcher to
`/usr/bin/pyautoimgtranslate`, plus a desktop entry and icon. Dependencies
(`python3-tk`, `libgl1`, `libglib2.0-0`) are pulled in automatically; Tesseract,
`xclip` and a screenshot tool are recommended.

## Release checklist

1. `python tools/check_sensitive.py` — no API keys, tokens, logs, screenshots,
   local settings or build artifacts in the tree **or in the git history**.
   Secret values are printed masked, never in full.
2. `python tests/test_offline.py && python tests/test_gui_smoke.py && python tests/test_translation_flow.py`
3. Bump `APP_VERSION` in `main.py`, the default version in `build_deb.sh`, and add a
   `RELEASE_NOTES.md` section.
4. Build the executable (`build.bat`) and, on Linux, the `.deb` (`./build_deb.sh X.Y.Z`).
5. `git add -A && git commit && git tag -a vX.Y.Z -m "pyAutoImgTranslate X.Y.Z" && git push origin main --follow-tags`

The guard can run automatically before every push — enable it once per clone:

```bash
git config core.hooksPath .githooks
```

**Never commit**: `config.json` (may hold the API key), `*.log` (may hold text from
the screen), screenshots (`debug_*.png`), `build/` and `dist/`. All of these are in
`.gitignore`; the checker verifies it.

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

Check boxes in the settings window control `add_to_autostart`, `keep_image_on_top`,
`live_on_top`, `skip_no_text`, `stream_results` and `analyze_web_search`.

### Analysis (description + web search)

The analysis hotkey runs two model calls and a search in between:

1. the vision model describes the capture and, when it makes sense, appends a
   `ПОИСК:` block with 1–3 short search queries (or `ПОИСК: НЕТ`);
2. the queries are searched **in parallel** (Wikipedia articles and extracts first,
   then Bing RSS), every hit is checked against the query, and decoy/unrelated
   results are dropped;
3. the model receives the description plus the sources and writes the answer with
   `[1]`, `[2]` references.

Results are cached for 10 minutes, and the search bypasses a dead proxy
automatically (it tries a direct connection first, then the proxy from the
environment). If nothing relevant is found, the model answers from its own
knowledge and says so instead of inventing facts.

> Privacy: the generated search queries leave the machine (Wikipedia/Bing) — the
> screenshot itself never does. Disable `analyze_web_search` to keep the analysis
> fully local apart from the model request.

### Cancelling and re-shooting

Processing can be interrupted at any moment: with the **Отмена** button in the
progress plate, with **Отменить** in the result window, from the tray
(**Отменить обработку**) or by simply pressing the capture hotkey again — a new
capture always supersedes the previous one, and the open HTTP request is closed
immediately.

## Tests

Offline checks (no network needed, Tesseract is used locally for the no-text
filter):

```bash
python tests/test_offline.py          # no-text filter, stream/analysis parsing, search layer, jobs
python tests/test_gui_smoke.py        # result window (translation + analysis), selector, settings
python tests/test_translation_flow.py # streaming, cancellation and analysis against a stub server
```

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
| `min_ocr_confidence` | Minimum mean OCR confidence to treat a capture as text | `40` |
| `tesseract_cmd` | Path to `tesseract.exe` | `C:\Program Files\Tesseract-OCR\tesseract.exe` |
| `source_lang` | Source language hint | `авто` |
| `target_lang` | Translation target language | `русский` |
| `request_timeout` | HTTP timeout, seconds | `120` |
| `skip_no_text` | Skip obviously empty captures (no API request); text on a busy background is still sent | `true` |
| `stream_results` | Show the translation while it is being generated | `true` |
| `live_on_top` | Keep the live-mode result window always on top | `true` |
| `capture_timeout` | Seconds to wait for the area selection | `60` |
| `capture_mode` | `builtin` — own full-screen selector, `system` — system screenshot tool | `builtin` (Windows) |
| `capture_dim` | How much the screen is dimmed while selecting (`0` — no dimming) | `0.45` |
| `log_recognized_text` | Write recognized screen text to the log file | `false` |
| `add_to_autostart` | Launch with Windows | `false` |
| `keep_image_on_top` | Result window always on top | `false` |
| `hotkey` | Global hotkey: capture area → OCR → translate (in `builtin` mode `Win+Shift+S` is replaced by `Win+Shift+X`) | `windows+shift+s` |
| `analyze_hotkey` | Global hotkey: capture area → describe + web search | `windows+shift+i` |
| `analyze_web_search` | Search the web during analysis | `true` |
| `search_results` | Results per search query | `4` |
| `search_timeout` | Timeout of one search request, seconds | `8` |
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

On Windows the environment variable is also looked up in the registry (`HKCU\Environment`,
then the machine environment) when the process itself does not see it — that happens
when the app was started by a shortcut that predates the variable, or by Task
Scheduler. At startup the log records where the key came from (never the key itself),
and a 401 without any key is reported as "API-ключ не найден".

---

The project is designed for fast and effective translation of texts from the screen without unnecessary actions.
