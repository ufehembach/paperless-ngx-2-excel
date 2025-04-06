#!/bin/bash

# Installiert Pakete aus requirements.txt
# Optional: mit --break-system-packages wenn "--break" als Argument übergeben wird

REQ_FILE="requirements.txt"

if [ ! -f "$REQ_FILE" ]; then
  echo "❌ $REQ_FILE nicht gefunden."
  exit 1
fi

PIP_CMD="pip install -r $REQ_FILE"

if [ "$1" == "--break" ]; then
  PIP_CMD="$PIP_CMD --break-system-packages"
fi

echo "🚀 Führe aus: $PIP_CMD"
eval $PIP_CMD

