#!/bin/bash
# file_watcher.sh — Detecta nuevos archivos en /srv/rpf/datos/ y notifica n8n
#
# Requiere: sudo apt-get install -y inotify-tools
# Instalado como servicio: systemd/rpf-watcher.service
#
# Eventos detectados: close_write (archivo completado, no write parcial)
# Filtra: solo .xlsx, .csv, .json, .emf

DATA_DIR="/srv/rpf/datos"
N8N_URL="${N8N_WEBHOOK:-http://localhost:5678/webhook/rpf-new-file}"
LOG_FILE="/srv/rpf/sync/logs/watcher.log"

mkdir -p "$(dirname "$LOG_FILE")"

echo "[$(date -Iseconds)] RPF File Watcher iniciado — observando: $DATA_DIR" | tee -a "$LOG_FILE"

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
        xlsx|csv|json|emf)
            FILENAME=$(basename "$FILEPATH")
            TS=$(date -Iseconds)

            echo "[$TS] Nuevo archivo: $FILEPATH" | tee -a "$LOG_FILE"

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
                 -d "$PAYLOAD" >> "$LOG_FILE" 2>&1

            echo "" >> "$LOG_FILE"
            ;;
    esac
done
