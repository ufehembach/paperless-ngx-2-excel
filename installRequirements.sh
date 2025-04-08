#!/bin/bash

# Installiert Pakete aus requirements.txt
# Optional: mit --break-system-packages wenn "--break" als Argument übergeben wird

REQ_FILE="requirements.txt"

if [ ! -f "$REQ_FILE" ]; then
  echo "❌ $REQ_FILE nicht gefunden."
  exit 1
fi

# Finde passenden pip-Befehl
if command -v pip >/dev/null 2>&1; then
  PIP_BIN="pip"
elif command -v pip3 >/dev/null 2>&1; then
  PIP_BIN="pip3"
else
  echo "❌ Weder pip noch pip3 gefunden."
  exit 1
fi

PIP_CMD="$PIP_BIN install -r $REQ_FILE"

if [ "$1" == "--break" ]; then
  PIP_CMD="$PIP_CMD --break-system-packages"
fi

echo "🚀 Führe aus: $PIP_CMD"
eval $PIP_CMD

