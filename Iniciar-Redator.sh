#!/bin/sh
set -eu
base=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
programa="$base/out/Redator-linux-$(uname -m)/Redator"
case "$(uname -m)" in
    x86_64) programa="$base/out/Redator-linux-x64/Redator" ;;
    aarch64) programa="$base/out/Redator-linux-arm64/Redator" ;;
esac
if [ ! -x "$programa" ]; then
    printf '%s\n' "Pacote ausente. Execute npm ci e npm run package na pasta do Redator." >&2
    exit 1
fi
exec "$programa" "$@"
