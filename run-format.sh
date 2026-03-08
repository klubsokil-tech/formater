#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

echo "=============================================="
echo "  Academic DOCX Formatter (macOS/Linux)"
echo "=============================================="

echo "[1/4] Перевірка Node.js..."
if ! command -v node >/dev/null 2>&1; then
  echo "ПОМИЛКА: Node.js не знайдено у PATH."
  echo "Встановіть LTS-версію з https://nodejs.org/ і запустіть скрипт знову."
  exit 1
fi

echo "[2/4] Перевірка npm..."
if ! command -v npm >/dev/null 2>&1; then
  echo "ПОМИЛКА: npm не знайдено у PATH."
  echo "Перевстановіть Node.js LTS з https://nodejs.org/ і повторіть спробу."
  exit 1
fi

if [ ! -d node_modules ]; then
  echo "[3/4] Встановлення залежностей (npm install)..."
  npm install
else
  echo "[3/4] Залежності вже встановлені."
fi

INPUT="Розділ_3_Практичний_аналіз_функціонування_іншомовної_лексики_в_корейських.docx"
OUTPUT="output.docx"

if [ ! -f "$INPUT" ]; then
  echo "ПОМИЛКА: вхідний файл не знайдено: $INPUT"
  echo "Помістіть файл у папку проєкту і запустіть скрипт ще раз."
  exit 1
fi

echo "[4/4] Форматування документа..."
node formatDocx.js "$INPUT" "$OUTPUT"

echo "Перевірка результату..."
if ! node verify.js "$OUTPUT"; then
  echo "УВАГА: форматування виконано, але verify.js знайшов невідповідності."
  exit 2
fi

echo "Готово! Створено файл: $OUTPUT"
