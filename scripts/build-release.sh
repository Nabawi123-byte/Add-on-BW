#!/usr/bin/env bash
# Membangun paket instalasi MASTER BW (dist/MasterBW) — bisa dijalankan di Linux/macOS/Windows (Git Bash).
# Hasil: dist/MasterBW-<versi>.zip berisi add-in, mesin AI (win-x64, self-contained), fungsi Excel, dan skrip instal.
set -euo pipefail
cd "$(dirname "$0")/.."
VERSION=$(grep -oP '(?<=<Version>)[^<]+' Directory.Build.props)
OUT=dist/MasterBW
rm -rf dist && mkdir -p "$OUT"/{addin,engine,scripts,models,docs}

echo "==> Uji pustaka inti"
dotnet test tests/MasterBW.Core.Tests -c Release --nologo -v q

echo "==> Add-in Office (Word & Excel)"
dotnet build src/MasterBW.Office -c Release --nologo -v q -o "$OUT/addin"

echo "==> Fungsi sel Excel (Excel-DNA)"
dotnet build src/MasterBW.Functions -c Release --nologo -v q
mkdir -p "$OUT/addin/functions"
cp src/MasterBW.Functions/bin/Release/net48/publish/*-packed.xll "$OUT/addin/functions/"

echo "==> Mesin AI (win-x64, self-contained)"
dotnet publish src/MasterBW.Engine -c Release -r win-x64 --self-contained true --nologo -v q \
  -p:PublishSingleFile=false -p:DebugType=None -o "$OUT/engine"

# buang runtime non-Windows yang ikut terbawa paket Whisper/LLamaSharp
find "$OUT/engine/runtimes" -mindepth 1 -maxdepth 1 ! -name 'win-x64' -exec rm -rf {} +
rm -f "$OUT/engine"/*.metal "$OUT/addin"/*.deps.json

cp scripts/install.ps1 scripts/uninstall.ps1 scripts/download-model.ps1 "$OUT/scripts/"
cp scripts/INSTALL.cmd scripts/UNINSTALL.cmd "$OUT/"
cp docs/*.md "$OUT/docs/" 2>/dev/null || true
cp README.md "$OUT/BACA-SAYA.md"
echo "Letakkan file model .gguf di folder ini, atau jalankan scripts\\download-model.ps1" > "$OUT/models/TARUH-MODEL-DI-SINI.txt"

(cd dist && zip -qr "MasterBW-$VERSION.zip" MasterBW)
echo "==> Selesai: dist/MasterBW-$VERSION.zip ($(du -h dist/MasterBW-$VERSION.zip | cut -f1))"
