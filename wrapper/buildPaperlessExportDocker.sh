#!/usr/bin/env bash
set -euo pipefail

IMAGE_NAME="${IMAGE_NAME:-ufe/paperless-export:latest}"

# Skript-eigener Ordner (wrapper/) und Repo-Root (eine Ebene höher) ermitteln,
# damit es egal ist, von wo aus das Script aufgerufen wird
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONTEXT_PATH="$(cd "$SCRIPT_DIR/.." && pwd)"
DOCKERFILE_PATH="$SCRIPT_DIR/Dockerfile"

TIMESTAMP_FILE="$SCRIPT_DIR/.last_docker_build"
WATCHLIST_FILE="$SCRIPT_DIR/.docker-build-watchlist"

# Standard-Dateien, die beim Build überwacht werden (relativ zu CONTEXT_PATH = Repo-Root)
DEFAULT_WATCH=(
  "wrapper/Dockerfile"
  "wrapper/doPaperlessExports.sh"
  "script/createRequirements.sh"
  "script/installRequirements.sh"
  "script/LICENSE"
  "script/paperless-ngx-2-excel.ini"
  "script/paperless-ngx-2-excel.py"
  "script/.version"
  "script/paperless-ngx-2-excel.ufe.ini"
  "script/requirements.txt"
)

# Falls eine eigene Watchlist-Datei existiert, diese verwenden
if [[ -f "$WATCHLIST_FILE" ]]; then
  mapfile -t WATCH < "$WATCHLIST_FILE"
else
  WATCH=("${DEFAULT_WATCH[@]}")
fi

# Wenn Timestamp-Datei fehlt, sofort bauen
if [[ ! -f "$TIMESTAMP_FILE" ]]; then
  echo "⚠️  Kein vorheriger Build-Zeitstempel gefunden. Baue neues Image…"
  docker build -t "$IMAGE_NAME" -f "$DOCKERFILE_PATH" "$CONTEXT_PATH"
  date +%s > "$TIMESTAMP_FILE"
  exit 0
fi

last_build=$(<"$TIMESTAMP_FILE")
needs_build=false
newer_example=""

echo "📦 Letzter Build-Zeitstempel: $last_build"
echo "📁 Kontext: $CONTEXT_PATH"
echo "📄 Überwachte Dateien:"

for f in "${WATCH[@]}"; do
  full_path="$CONTEXT_PATH/$f"
  if [[ ! -f "$full_path" ]]; then
    echo "  ❌ $full_path (existiert nicht)"
    continue
  fi

  file_mtime=$(stat -c %Y "$full_path" || echo 0)
  echo "  ⏱ $full_path (mtime=$file_mtime)"

  if (( file_mtime > last_build )); then
    needs_build=true
    newer_example="$full_path"
    break
  fi
done

if $needs_build; then
  echo "⚠️  Neuere Datei gefunden (z. B. $newer_example). Starte neuen Build…"
  docker build -t "$IMAGE_NAME" -f "$DOCKERFILE_PATH" "$CONTEXT_PATH"
  date +%s > "$TIMESTAMP_FILE"
else
  echo "✅  Kein neuer Build nötig – alle Dateien älter als letzter Build."
fi
