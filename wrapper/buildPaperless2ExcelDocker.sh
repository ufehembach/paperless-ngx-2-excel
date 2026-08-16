#!/usr/bin/env bash
set -euo pipefail

IMAGE_NAME="${IMAGE_NAME:-ufe/paperless2excel:latest}"
TIMESTAMP_FILE=".last_docker_build"
WATCHLIST_FILE=".docker-build-watchlist"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONTEXT_PATH="$(cd "$SCRIPT_DIR/.." && pwd)"
DOCKERFILE_PATH="$SCRIPT_DIR/Dockerfile"

# Standard-Dateien, die beim Build überwacht werden
DEFAULT_WATCH=(
  "wrapper/Dockerfile"
  "wrapper/doPaperlessExports.sh"
  "script/createRequirements.sh"
  "script/installRequirements.sh"
  "script/LICENSE"
  "script/paperless2excel.ini"
  "script/paperless2excel.py"
  ".version"
  "script/paperless2excel.ufe.ini"
  "script/requirements.txt"
  "script/paperless_export"
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

# Timestamp des letzten Builds lesen
last_build=$(<"$TIMESTAMP_FILE")
needs_build=false
newer_example=""

# Dateien prüfen
# Wenn Timestamp-Datei fehlt, sofort bauen
if [[ ! -f "$TIMESTAMP_FILE" ]]; then
  echo "⚠️  Kein vorheriger Build-Zeitstempel gefunden. Baue neues Image…"
  docker build -t "$IMAGE_NAME" -f "$DOCKERFILE_PATH" "$CONTEXT_PATH"
  date +%s > "$TIMESTAMP_FILE"
  exit 0
fi

# Timestamp des letzten Builds lesen
last_build=$(<"$TIMESTAMP_FILE")
needs_build=false
newer_example=""

echo "📦 Letzter Build-Zeitstempel: $last_build"
echo "📁 Kontext: $CONTEXT_PATH"
echo "📄 Überwachte Dateien:"

# Dateien und Ordner prüfen
for f in "${WATCH[@]}"; do
  full_path="$CONTEXT_PATH/$f"

  if [[ -d "$full_path" ]]; then
    # Ordner-Eintrag: rekursiv die zuletzt geänderte Datei darin finden
    newest_in_dir=$(find "$full_path" -type f -printf '%T@ %p\n' 2>/dev/null | sort -rn | head -1)
    if [[ -z "$newest_in_dir" ]]; then
      echo "  ❌ $full_path/ (leer oder existiert nicht)"
      continue
    fi
    file_mtime="${newest_in_dir%% *}"
    file_mtime="${file_mtime%.*}"  # Nachkommastellen von %T@ kappen
    newest_file="${newest_in_dir#* }"
    echo "  ⏱ $full_path/ (neueste Datei: $newest_file, mtime=$file_mtime)"
    if (( file_mtime > last_build )); then
      needs_build=true
      newer_example="$newest_file"
      break
    fi
    continue
  fi

  # Datei existiert nicht → überspringen
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

# Build auslösen, wenn nötig
if $needs_build; then
  echo "⚠️  Neuere Datei gefunden (z. B. $newer_example). Starte neuen Build…"
  docker build -t "$IMAGE_NAME" -f "$DOCKERFILE_PATH" "$CONTEXT_PATH"
  date +%s > "$TIMESTAMP_FILE"
else
  echo "✅  Kein neuer Build nötig – alle Dateien älter als letzter Build."
fi
##
