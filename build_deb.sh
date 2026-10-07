#!/usr/bin/env bash
# Сборка .deb-пакета pyAutoImgTranslate для Debian/Ubuntu.
#
# Использует dist/pyAutoImgTranslate (уже собранный PyInstaller-бинарник).
# Результат: dist/pyautotimgtranslate_<VERSION>_<ARCH>.deb
#
# Требуется: dpkg-deb (пакет dpkg-dev). Сборку запускать в Debian/Ubuntu.
set -euo pipefail

VERSION="${1:-1.6.0}"
ARCH="$(dpkg --print-architecture)"
PKG="pyautoimgtranslate"
ROOT="$(cd "$(dirname "$0")" && pwd)"
BIN="${ROOT}/dist/pyAutoImgTranslate"
OUT="${ROOT}/dist/${PKG}_${VERSION}_${ARCH}.deb"

if [[ ! -f "${BIN}" ]]; then
    echo "Не найден бинарник ${BIN}. Сначала соберите релиз: ./build.sh" >&2
    exit 1
fi

STAGE="$(mktemp -d)"
trap 'rm -rf "${STAGE}"' EXIT

# --- Файловая структура пакета ---
install -Dm755 "${BIN}" "${STAGE}/usr/lib/${PKG}/pyAutoImgTranslate"
install -Dm644 "${ROOT}/translator.png" "${STAGE}/usr/share/icons/hicolor/256x256/apps/${PKG}.png"

# Обёртка в /usr/bin (обёртка, а не симлинк, чтобы корректно указывать имя)
install -Dm755 /dev/stdin "${STAGE}/usr/bin/${PKG}" <<'EOF'
#!/bin/sh
exec /usr/lib/pyautoimgtranslate/pyAutoImgTranslate "$@"
EOF

# .desktop — запись в меню приложений
install -Dm644 /dev/stdin "${STAGE}/usr/share/applications/${PKG}.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=pyAutoImgTranslate
Comment=Перевод текста с экрана (OCR + нейросеть)
Exec=${PKG}
Icon=${PKG}
Terminal=false
Categories=Utility;Accessibility;Graphics;
Keywords=translate;ocr;screenshot;перевод;скриншот;
EOF

# --- Метаданные DEBIAN/control ---
mkdir -p "${STAGE}/DEBIAN"
cat > "${STAGE}/DEBIAN/control" <<EOF
Package: ${PKG}
Version: ${VERSION}
Section: utils
Priority: optional
Architecture: ${ARCH}
Maintainer: hvkeyn <https://github.com/hvkeyn>
Depends: python3-tk, libgl1, libglib2.0-0
Recommends: tesseract-ocr, tesseract-ocr-eng, tesseract-ocr-rus, xclip
Suggests: gnome-screenshot | spectacle | scrot | maim | flameshot | grim
Description: Screen text translation via OCR and AI
 pyAutoImgTranslate captures a screen area, recognizes the text
 (Tesseract or a vision-capable neural network) and translates it
 close to the original, showing the result in a small window.
 Supports DeepSeek, LM Studio and any OpenAI-compatible server.
EOF

# --- postinst: обновить кеши иконок/меню, если есть инструменты ---
cat > "${STAGE}/DEBIAN/postinst" <<'EOF'
#!/bin/sh
set -e
if command -v update-desktop-database >/dev/null 2>&1; then
    update-desktop-database -q /usr/share/applications || true
fi
if command -v gtk-update-icon-cache >/dev/null 2>&1; then
    gtk-update-icon-cache -q -t -f /usr/share/icons/hicolor || true
fi
exit 0
EOF
chmod 755 "${STAGE}/DEBIAN/postinst"

# --- prerm: почистить автозапуск пользователя не нужно (он в $HOME) ---
cat > "${STAGE}/DEBIAN/prerm" <<'EOF'
#!/bin/sh
set -e
exit 0
EOF
chmod 755 "${STAGE}/DEBIAN/prerm"

echo "== Сборка .deb =="
dpkg-deb --build --root-owner-group "${STAGE}" "${OUT}"
echo "Готово: ${OUT}"
