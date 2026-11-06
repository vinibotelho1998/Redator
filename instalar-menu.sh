#!/bin/sh
set -eu
base=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
dest="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
mkdir -p "$dest"
python3 - "$base" "$dest/br.redator.app.desktop" <<'PY'
from pathlib import Path
import sys
base, dest = map(Path, sys.argv[1:])
# Escape Desktop Entry strings and Exec arguments, including literal field codes.
def scalar(s):
    return str(s).replace('\\', '\\\\').replace('\n', '\\n').replace('\r', '\\r').replace('\t', '\\t')
command = str(base / 'Iniciar-Redator.sh').replace('%', '%%')
for char in ['\\', '"', '$', '`']:
    command = command.replace(char, '\\' + char)
dest.write_text('\n'.join([
    '[Desktop Entry]', 'Type=Application', 'Name=Redator', 'Comment=Editor de documentos',
    'Exec="' + scalar(command) + '" %F', 'Icon=' + scalar(base / 'app/icons/redator.png'),
    'Terminal=false', 'Categories=Office;WordProcessor;', 'StartupWMClass=Redator',
    'MimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document;application/msword;'
    'application/vnd.oasis.opendocument.text;application/pdf;application/rtf;text/rtf;text/markdown;'
    'text/plain;text/csv;text/tab-separated-values;application/epub+zip;text/html;message/rfc822;', ''
]), encoding='utf-8')
PY
if command -v update-desktop-database >/dev/null 2>&1; then update-desktop-database "$dest"; fi
printf '%s\n' "Redator registrado no menu e em Abrir com."
