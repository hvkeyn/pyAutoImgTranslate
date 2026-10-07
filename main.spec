# -*- mode: python ; coding: utf-8 -*-
import sys

is_windows = sys.platform.startswith("win")

# Скрытые импорты зависят от ОС: бэкенд трея и библиотека хоткеев.
if is_windows:
    hidden = ["pystray._win32", "keyboard", "win32clipboard", "win32con", "win32api"]
else:
    hidden = ["pystray._xorg", "pystray._appindicator", "pynput"]

# Иконка: на Windows — .ico/.png, на Linux иконка встраивается в data.
icon_arg = ["translator.png"] if is_windows else None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('translator.png', '.')],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='pyAutoImgTranslate',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_arg,
)
