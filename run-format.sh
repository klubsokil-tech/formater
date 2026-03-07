#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
if [ ! -d node_modules ]; then
  npm install
fi
INPUT="Розділ_3_Практичний_аналіз_функціонування_іншомовної_лексики_в_корейських.docx"
OUTPUT="output.docx"
node formatDocx.js "$INPUT" "$OUTPUT"
node verify.js "$OUTPUT"
