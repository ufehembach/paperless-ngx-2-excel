#!/usr/bin/env bash
set -euo pipefail

# --- names/paths you may tune ---
IMAGE="ufe/paperless2excel:latest"
NAME="paperless-export"
NET="paperless_default"
EXPORTS_HOST="/volume1/paperless-export/exports"   # host path mounted into container at /exports
WORKDIR="$(cd "$(dirname "$0")" && pwd)"
cd "$WORKDIR"

# --- file/lock naming like before ---
filename="$(basename -- "$0")"
filename="${filename%.*}"
logfile="$WORKDIR/$filename.log"
lockFile="$WORKDIR/$filename.lock"

# --- lock ---
HEARTBEAT_INTERVAL=60   # Sekunden zwischen Heartbeats

if [[ -f "$lockFile" ]]; then
  echo "lockfile found   : $lockFile"
  cat "$lockFile"
  exit 0
fi

{
  echo "start : $(date '+%Y-%m-%d %H:%M:%S')"
  echo "pid   : $$"
} > "$lockFile"

heartbeatPid=""

start_heartbeat() {
  (
    while true; do
      sleep "$HEARTBEAT_INTERVAL"
      echo "heartbeat: $(date '+%Y-%m-%d %H:%M:%S')" >> "$lockFile" 2>/dev/null || true
    done
  ) &
  heartbeatPid=$!
  disown "$heartbeatPid" 2>/dev/null || true
}

stop_heartbeat() {
  if [[ -n "$heartbeatPid" ]]; then
    kill "$heartbeatPid" 2>/dev/null || true
    wait "$heartbeatPid" 2>/dev/null || true
    heartbeatPid=""
  fi
}

finalize() {
  local exit_code=$?
  stop_heartbeat
  rm -f "$lockFile"

  # altes ok/KO Log dieses Jobs entfernen, bevor neu benannt wird
  rm -f "$WORKDIR/${filename}.ok.log" "$WORKDIR/${filename}.KO.log" 2>/dev/null || true

  if [[ -f "$logfile" ]]; then
    if [[ $exit_code -eq 0 ]]; then
      mv -f "$logfile" "$WORKDIR/${filename}.ok.log"
    else
      echo "$(date '+%Y-%m-%d %H:%M:%S') | ABBRUCH mit Exit-Code $exit_code" >> "$logfile"
      mv -f "$logfile" "$WORKDIR/${filename}.KO.log"
    fi
  fi
  exit $exit_code
}

trap finalize EXIT
trap 'exit 130' INT

start_heartbeat

# --- helpers ---
progressF="__Current__${filename}__$(date --iso).progress.txt"
function log_section() { echo "-----------------------" | tee -a "$logfile"; }
function log_msg() { printf '%s | %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" | tee -a "$logfile"; }

# --- network ensure ---
if ! docker network inspect "$NET" >/dev/null 2>&1; then
  echo "[net] creating $NET ..."
  docker network create "$NET" >/dev/null
fi

# --- doTheJob (merged) ---
doTheJob() {
  if [[ $# -ne 2 ]]; then
    echo "usage doTheJob <para1> <para2>"
    return 1
  fi

  : > "$logfile"
  : > "$progressF"

  log_msg "Job gestartet (para1=$1, para2=$2)"

  {
    echo "para1: $1"
    echo "para2: $2"
    echo "pwd: $(pwd)"
    echo
    echo "=== docker env ==="
    which bash || true
    which docker || true
    docker ps || true

    echo
    log_section
    echo "$(date) start"

    # ---------- BUILD  ----------
    log_msg "Build: ./buildPaperlessExportDocker.sh"
   	./buildPaperlessExportDocker.sh
    # ---------- RUN (one-shot) ----------
    echo
    echo "[run] removing old container (if any)"
    docker rm -f "$NAME" >/dev/null 2>&1 || true

    log_msg "Export-Ordner: ${EXPORTS_HOST}"
    ls  "${EXPORTS_HOST}" 

    echo
    log_msg "Starte Container $IMAGE (Netz: $NET)"
    docker run --rm \
      --name "$NAME" \
      --network "$NET" \
      -e TZ=Europe/Berlin \
      -e EXPORT_DIR=/exports \
      -e FORCE_COPY=1 \
      -v "${EXPORTS_HOST}:/exports" \
      -v "${EXPORTS_HOST}:/app/exports" \
      "$IMAGE"

    log_msg "Container-Export beendet"

    # ---------- post-processing on host ----------
    echo
    echo "[post] cleanup marker + rotate last files"
    myDate="$(date --iso)"

    # remove any old markers in /volume1/paperless-export (host)
    rm -f /volume1/paperless-export/___${filename}*.last.txt 2>/dev/null || true

    # ZIP rotation (keep at least 7; delete >7d old beyond that)
    echo
    echo "[post] rotate ZIPs (keep at least 7)"
    MIN_FILES=7
    TOTAL_FILES="$(find . -name "*.zip" -type f | wc -l || echo 0)"
    if (( TOTAL_FILES <= MIN_FILES )); then
      echo "Abbruch: Es gibt nur $TOTAL_FILES ZIP-Dateien. Mindestens $MIN_FILES müssen verbleiben."
    else
      # delete .zip older than 7 days, but keep MIN_FILES newest overall
      # Safe approach: list all, sort by mtime ascending, delete all but last MIN_FILES
      mapfile -t ALL_ZIPS < <(find . -name "*.zip" -type f -mtime +7 -printf "%T@ %p\n" | sort -n | awk '{ $1=""; sub(/^ /,""); print }')
      if (( ${#ALL_ZIPS[@]} > MIN_FILES )); then
        DEL_COUNT=$(( ${#ALL_ZIPS[@]} - MIN_FILES ))
        for ((i=0; i<DEL_COUNT; i++)); do
          f="${ALL_ZIPS[$i]}"
          echo "Lösche: $f"
          rm -f -- "$f"
        done
      else
        echo "Keine ZIPs zu löschen (älter als 7 Tage minus Mindestanzahl)."
      fi
    fi

    echo
    echo "[post] chown ufe:users -R ."
    chown ufe:users . -R || true

    echo
    echo "========================="
    ls -al /volume1/docker/paperless || true
    echo "+++++++++++++++++++++++++"
    ls -al /volume1/paperless-export || true
    echo "+++++++++++++++++++++++++"

  } | tee -a myLog.log 2> myError.log

  # finalize last log like before
  rm -f "${filename}"*.last.log 2>/dev/null || true
  myDate="$(date --iso)"
  myLastFile="${filename}__${myDate}.last.log"
  {
    echo "------ $(date --iso-8601=seconds)"
    echo "++++++Error++++++"
    cat myError.log
    echo "------Log--------"
    cat myLog.log
    echo "......LS....."
    ls -l
    echo "...end. $(date) ..."
  } > "$myLastFile"
  rm -f myError.log myLog.log "$progressF"

  # -------- copy_folders (host-side) --------
  source_folder="/volume1/paperless-export/exports"
  base_folder="/volume1/rkp-nebenkosten/belege"
  log_msg "Kopiere NK/Miete-Ordner: $source_folder -> $base_folder"
  copy_nk_and_miete_folders "$source_folder" "$base_folder"
  cd $source_folder
  /volume1/paperless-export/exports/folder_summary.sh
  cd -

  log_msg "Job fertig"
}

copy_nk_and_miete_folders() {
  local source_dir="$1"   # z.B. /volume1/paperless-export/exports
  local target_dir="$2"   # z.B. /volume1/rkp-nebenkosten/belege

  echo "--------------------------------------------------"
  echo "[INFO] Starte gefilterte Kopie"
  echo "[INFO] Quelle: $source_dir"
  echo "[INFO] Ziel  : $target_dir"
  echo "--------------------------------------------------"

  shopt -s nullglob

  mkdir -p "$target_dir"

  # 1) Alle passenden Ordner im source einsammeln
  local dir name
  local -a selected=()

  for dir in "$source_dir"/*/; do
    name="${dir%/}"
    name="${name##*/}"

    if [[ "$name" == *Nebenkosten*  || \
          "$name" == *NK*           || \
          "$name" == *Mietvertrag*  || \
          "$name" == *Miete*        ]]; then
      echo "[INFO] Gefunden: $name"
      selected+=("$name")
    else
      echo "[DEBUG] Ignoriere: $name"
    fi
  done

  if (( ${#selected[@]} == 0 )); then
    echo "[WARN] Keine passenden Ordner gefunden. Ziel wird geleert."
  fi

  # 2) Im Ziel Ordner löschen, die nicht (mehr) in selected sind
  local existing base keep
  for existing in "$target_dir"/*/; do
    base="${existing%/}"
    base="${base##*/}"
    keep=false

    for name in "${selected[@]}"; do
      if [[ "$base" == "$name" ]]; then
        keep=true
        break
      fi
    done

    if [[ "$keep" == false ]]; then
      echo "[INFO] Entferne veralteten Ziel-Ordner: $base"
      rm -rf -- "$existing"
    fi
  done

  # 3) Für jeden ausgewählten Ordner rsync mit --delete
  for name in "${selected[@]}"; do
    echo "[INFO] Synchronisiere: $name"
    rsync -av --delete \
      "$source_dir/$name/" \
      "$target_dir/$name/"
  done

  # 4) ACLs vom Parent des Zielordners übernehmen (Synology)
  if command -v synoacltool >/dev/null 2>&1; then
    local base_dir
    base_dir="$(dirname "$target_dir")"
    echo "[INFO] Kopiere ACLs von $base_dir → $target_dir"
    synoacltool -copy "$base_dir" "$target_dir" || true
  else
    echo "[INFO] synoacltool nicht gefunden – ACL-Kopie übersprungen."
  fi

  echo "[INFO] Fertig."
}

doTheJob Job Done
