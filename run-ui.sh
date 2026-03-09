#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

if [ ! -d node_modules ]; then
  npm install
fi

echo

echo "Запускаю Web UI на http://localhost:3000"
echo "Для зупинки натисніть Ctrl+C у цьому терміналі."
echo

if command -v xdg-open >/dev/null 2>&1; then
  xdg-open "http://localhost:3000" >/dev/null 2>&1 || true
elif command -v open >/dev/null 2>&1; then
  open "http://localhost:3000" >/dev/null 2>&1 || true
fi

node server.js
