#!/bin/bash
# Instalasi Master untuk Word & Excel di macOS.
set -e
cd "$(dirname "$0")"
DIR="$(pwd)"

if ! command -v node >/dev/null 2>&1; then
  echo "Node.js belum terpasang. Unduh versi LTS dari https://nodejs.org lalu jalankan lagi."
  exit 1
fi

echo "[1/4] Memasang komponen (butuh internet, hanya sekali)..."
npm install --no-fund --no-audit

echo "[2/4] Membuat sertifikat HTTPS lokal (masukkan password Mac bila diminta)..."
npx office-addin-dev-certs install --days 3650

echo "[3/4] Mendaftarkan Master ke Excel dan Word..."
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
cp manifest-excel.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/master-excel.xml
cp manifest-word.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/master-word.xml

echo "[4/4] Menjalankan server Master otomatis saat login..."
NODE_BIN="$(command -v node)"
PLIST=~/Library/LaunchAgents/com.master.addin.plist
mkdir -p ~/Library/LaunchAgents
cat > "$PLIST" <<PL
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>com.master.addin</string>
  <key>ProgramArguments</key><array><string>$NODE_BIN</string><string>$DIR/server.js</string></array>
  <key>WorkingDirectory</key><string>$DIR</string>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>StandardOutPath</key><string>$DIR/master-server.log</string>
  <key>StandardErrorPath</key><string>$DIR/master-server.log</string>
</dict></plist>
PL
launchctl unload "$PLIST" 2>/dev/null || true
launchctl load "$PLIST"

echo ""
echo "SELESAI. Tutup lalu buka lagi Excel / Word."
echo "Buka: Sisipkan (Insert) > Add-in Saya (My Add-ins) > Master."
