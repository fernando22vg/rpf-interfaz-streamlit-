#!/bin/bash
# setup_ollama_rpf.sh — Configura Ollama para servidor con VRAM limitada
# y crea el modelo personalizado rpf-cobee

set -e

echo "======================================================"
echo "  SETUP: Modelo RPF COBEE + Configuración VRAM"
echo "======================================================"

# ── 1. Variables de entorno Ollama (guardar en /etc/systemd/system/ollama.service.d/) ──
echo ""
echo "[1/4] Configurando variables de entorno Ollama..."

sudo mkdir -p /etc/systemd/system/ollama.service.d/

sudo tee /etc/systemd/system/ollama.service.d/override.conf > /dev/null << 'EOF'
[Service]
# Solo un modelo cargado a la vez (ahorra VRAM)
Environment="OLLAMA_MAX_LOADED_MODELS=1"
# Solo una petición paralela (evita OOM)
Environment="OLLAMA_NUM_PARALLEL=1"
# Descarga el modelo de VRAM si no se usa en 10 minutos (libera memoria)
Environment="OLLAMA_KEEP_ALIVE=10m"
# Tiempo máximo de respuesta: 10 minutos (evita que quede colgado)
Environment="OLLAMA_REQUEST_TIMEOUT=600"
# Flashattention para reducir uso de VRAM en modelos compatibles
Environment="OLLAMA_FLASH_ATTENTION=1"
EOF

echo "  ✓ Override systemd creado"

# ── 2. Recargar systemd y reiniciar Ollama ─────────────────────────────────
echo ""
echo "[2/4] Reiniciando servicio Ollama..."
sudo systemctl daemon-reload
sudo systemctl restart ollama
sleep 5

# Verificar que Ollama está corriendo
if ! systemctl is-active --quiet ollama; then
    echo "  [ERROR] Ollama no se reinició correctamente"
    sudo systemctl status ollama --no-pager | tail -10
    exit 1
fi
echo "  ✓ Ollama activo"

# ── 3. Crear modelo rpf-cobee ──────────────────────────────────────────────
echo ""
echo "[3/4] Creando modelo rpf-cobee desde Modelfile..."

MODELFILE_PATH="/home/joselozano/rpf-ejecucion/Modelfile.rpf"

if [ ! -f "$MODELFILE_PATH" ]; then
    echo "  [ERROR] No se encontró: $MODELFILE_PATH"
    echo "  Copia primero el Modelfile.rpf al servidor"
    exit 1
fi

ollama create rpf-cobee -f "$MODELFILE_PATH"
echo "  ✓ Modelo rpf-cobee creado"

# ── 4. Prueba rápida ───────────────────────────────────────────────────────
echo ""
echo "[4/4] Prueba rápida del modelo..."
RESPUESTA=$(ollama run rpf-cobee "En una sola línea: ¿cuál es tu especialidad?" 2>/dev/null || echo "TIMEOUT")

if [ "$RESPUESTA" = "TIMEOUT" ]; then
    echo "  [WARN] El modelo no respondió en tiempo — puede necesitar más tiempo la primera vez"
else
    echo "  Respuesta: $RESPUESTA"
    echo "  ✓ Modelo funcionando"
fi

echo ""
echo "======================================================"
echo "  COMPLETADO"
echo "  Modelo disponible: rpf-cobee"
echo "  En Open WebUI: selecciona 'rpf-cobee' como modelo"
echo "  y activa el knowledge base 'RPF COBEE'"
echo "======================================================"
echo ""
echo "  Modelos disponibles:"
ollama list
