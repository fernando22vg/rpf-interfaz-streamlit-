# deploy_to_server.ps1 — Despliega los archivos de Capa 1 al servidor Ubuntu
#
# Uso desde PowerShell en tu PC de desarrollo:
#   cd "c:\Programas Python\Interfaz_RPF\server"
#   .\deploy_to_server.ps1
#
# Requiere: OpenSSH instalado en Windows (ya viene con Windows 10/11)
# Ajusta SERVER_USER y SERVER_HOST si es necesario.

$SERVER_USER = "jose"
$SERVER_HOST = "192.168.0.92"
$SERVER = "$SERVER_USER@$SERVER_HOST"
$REMOTE_DIR = "/srv/rpf/sync"
$SYSTEMD_DIR = "/etc/systemd/system"

Write-Host "=== RPF Capa 1 — Deploy al servidor $SERVER_HOST ===" -ForegroundColor Cyan

# ── 1. Crear estructura de directorios en el servidor ─────────────────────────
Write-Host "`n[1/6] Creando directorios en el servidor..." -ForegroundColor Yellow
ssh $SERVER @"
sudo mkdir -p /srv/rpf/sync/logs
sudo mkdir -p /srv/rpf/datos/01_INFO_CNDC_RPF
sudo mkdir -p /srv/rpf/datos/02_DATOS_CNDC_RPF
sudo mkdir -p /srv/rpf/datos/DATOS_EXTRAIDOS
sudo mkdir -p /srv/rpf/codigo
sudo chown -R jose:jose /srv/rpf
sudo chmod 700 /srv/rpf/sync
echo 'Directorios creados OK'
"@

# ── 2. Copiar scripts Python ──────────────────────────────────────────────────
Write-Host "`n[2/6] Copiando scripts Python..." -ForegroundColor Yellow
scp "sharepoint_client_standalone.py" "${SERVER}:${REMOTE_DIR}/"
scp "sp_sync_daemon.py"               "${SERVER}:${REMOTE_DIR}/"
scp "requirements_server.txt"          "${SERVER}:${REMOTE_DIR}/"
scp "file_watcher.sh"                  "${SERVER}:${REMOTE_DIR}/"

# ── 3. Crear .env con template (solo si no existe ya) ─────────────────────────
Write-Host "`n[3/6] Creando .env template en el servidor..." -ForegroundColor Yellow
ssh $SERVER @"
if [ ! -f /srv/rpf/sync/.env ]; then
    cat > /srv/rpf/sync/.env << 'ENVEOF'
# RPF Intelligence Bridge — Variables de entorno
# IMPORTANTE: Completar SP_PASSWORD antes de activar el servicio
SP_PASSWORD=COMPLETAR_AQUI
N8N_WEBHOOK=http://localhost:5678/webhook/rpf-sync-done
SYNC_LOG=/srv/rpf/sync/logs/sync.log
DATA_ROOT=/srv/rpf/datos
N8N_WEBHOOK_FILES=http://localhost:5678/webhook/rpf-new-file
ENVEOF
    chmod 600 /srv/rpf/sync/.env
    echo '.env creado (recuerda completar SP_PASSWORD)'
else
    echo '.env ya existe, no se sobreescribe'
fi
"@

# ── 4. Instalar dependencias Python ──────────────────────────────────────────
Write-Host "`n[4/6] Instalando dependencias Python..." -ForegroundColor Yellow
ssh $SERVER "cd /srv/rpf/sync && pip3 install -r requirements_server.txt --quiet"

# ── 5. Instalar inotify-tools y copiar servicios systemd ─────────────────────
Write-Host "`n[5/6] Instalando inotify-tools y servicios systemd..." -ForegroundColor Yellow
scp "systemd/rpf-sync.service"    "${SERVER}:/tmp/"
scp "systemd/rpf-sync.timer"      "${SERVER}:/tmp/"
scp "systemd/rpf-watcher.service" "${SERVER}:/tmp/"

ssh $SERVER @"
sudo apt-get install -y inotify-tools --quiet
sudo cp /tmp/rpf-sync.service    $SYSTEMD_DIR/
sudo cp /tmp/rpf-sync.timer      $SYSTEMD_DIR/
sudo cp /tmp/rpf-watcher.service $SYSTEMD_DIR/
sudo chmod +x /srv/rpf/sync/file_watcher.sh
sudo systemctl daemon-reload
echo 'Servicios instalados (NO activados aún — completa SP_PASSWORD primero)'
"@

# ── 6. Crear base de datos PostgreSQL ────────────────────────────────────────
Write-Host "`n[6/6] Configurando base de datos PostgreSQL..." -ForegroundColor Yellow
scp "setup_db.sql" "${SERVER}:/tmp/"
ssh $SERVER @"
# Intentar crear la base de datos si no existe
sudo -u postgres psql -tc "SELECT 1 FROM pg_database WHERE datname = 'rpf_intelligence'" | grep -q 1 || \
    sudo -u postgres psql -c "CREATE DATABASE rpf_intelligence OWNER jose"
# Aplicar schema
psql -d rpf_intelligence -f /tmp/setup_db.sql
echo 'Base de datos rpf_intelligence configurada'
"@

# ── Instrucciones finales ─────────────────────────────────────────────────────
Write-Host @"

=== Deploy completado ===

PRÓXIMOS PASOS (ejecutar en el servidor via SSH):

1. Completar la contraseña de SharePoint:
   nano /srv/rpf/sync/.env
   # Cambiar: SP_PASSWORD=COMPLETAR_AQUI  →  SP_PASSWORD=tu_contraseña

2. Probar el sync en modo dry-run:
   cd /srv/rpf/sync
   python3 sp_sync_daemon.py --dry-run

3. Si el dry-run funciona, hacer el primer sync real:
   python3 sp_sync_daemon.py

4. Activar los servicios systemd:
   sudo systemctl enable --now rpf-sync.timer
   sudo systemctl enable --now rpf-watcher.service

5. Verificar que el timer está activo:
   systemctl list-timers rpf-sync.timer

6. Ver logs:
   tail -f /srv/rpf/sync/logs/sync.log
   journalctl -u rpf-sync -f

"@ -ForegroundColor Green
