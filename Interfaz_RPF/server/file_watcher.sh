#!/bin/bash
# file_watcher.sh — Detecta nuevos archivos en rpf-proyecto-datos/ y notifica n8n
#
# Requiere: sudo apt-get install -y inotify-tools
# Instalado como servicio: systemd/rpf-watcher.service
#
# Eventos detectados: close_write (archivo completado, no write parcial)
# Filtra: solo .xlsx, .xls, .csv, .json, .txt, .pdf

DATA_DIR="${DATA_ROOT:-/home/joselozano/rpf-proyecto-datos}"
N8N_URL="${N8N_WEBHOOK:-http://localhost:5678/webhook/rpf-new-file}"
LOG_FILE="${SYNC_LOG:-/home/joselozano/rpf-ejecucion/logs/sync.log}"
WATCHER_LOG="$(dirname "$LOG_FILE")/watcher.log"

mkdir -p "$(dirname "$WATCHER_LOG")"

echo "[$(date -Iseconds)] RPF File Watcher iniciado — observando: $DATA_DIR" | tee -a "$WATCHER_LOG"

inotifywait \
    --monitor \
    --recursive \
    --event close_write \
    --format '%w%f' \
    "$DATA_DIR" 2>/dev/null | \
while IFS= read -r FILEPATH; do
    # Filtrar solo extensiones relevantes
    EXT="${FILEPATH##*.}"
    EXT_LOWER=$(echo "$EXT" | tr '[:upper:]' '[:lower:]')

    case "$EXT_LOWER" in
        xlsx|xls|csv|json|txt|pdf)
            FILENAME=$(basename "$FILEPATH")
            TS=$(date -Iseconds)

            echo "[$TS] Nuevo archivo: $FILEPATH" | tee -a "$WATCHER_LOG"

            # Clasificar tipo de archivo para n8n
            if echo "$FILEPATH" | grep -q "01_INFO"; then
                TIPO="evento"
            elif echo "$FILEPATH" | grep -q "02_DATOS"; then
                TIPO="scada"
            elif echo "$FILEPATH" | grep -q "DATOS_EXTRAIDOS"; then
                TIPO="mapeo"
            else
                TIPO="desconocido"
            fi

            # Notificar a n8n con metadata
            PAYLOAD=$(printf '{"file":"%s","filename":"%s","tipo":"%s","ts":"%s"}' \
                "$FILEPATH" "$FILENAME" "$TIPO" "$TS")

            curl --silent \
                 --max-time 5 \
                 --retry 2 \
                 -X POST "$N8N_URL" \
                 -H "Content-Type: application/json" \
                 -d "$PAYLOAD" >> "$WATCHER_LOG" 2>&1

            echo "" >> "$WATCHER_LOG"
            ;;
    esac
done
