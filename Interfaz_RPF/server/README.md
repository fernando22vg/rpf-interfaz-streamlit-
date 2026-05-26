# RPF Intelligence Bridge — Capa 1: Data Bridge

Sincronización incremental de datos SharePoint → Ubuntu Server 192.168.0.92

## Estructura de archivos

```
server/
├── sharepoint_client_standalone.py  # Cliente SP sin dependencia de Streamlit
├── sp_sync_daemon.py                # Daemon de sync incremental
├── requirements_server.txt          # Dependencias mínimas (requests, python-dotenv)
├── file_watcher.sh                  # Watcher inotifywait → n8n webhook
├── setup_db.sql                     # Schema PostgreSQL (Capa 1)
├── deploy_to_server.ps1             # Script de despliegue desde Windows
└── systemd/
    ├── rpf-sync.service             # Servicio oneshot para el sync
    ├── rpf-sync.timer               # Timer: cada 15 minutos
    └── rpf-watcher.service          # Servicio watcher continuo
```

## Despliegue rápido (desde Windows)

```powershell
cd "c:\Programas Python\Interfaz_RPF\server"
.\deploy_to_server.ps1
```

## Estructura en el servidor después del deploy

```
/srv/rpf/
├── sync/
│   ├── sharepoint_client_standalone.py
│   ├── sp_sync_daemon.py
│   ├── file_watcher.sh
│   ├── requirements_server.txt
│   ├── .env                    ← SP_PASSWORD aquí (chmod 600)
│   └── logs/
│       ├── sync.log
│       └── watcher.log
├── datos/
│   ├── 01_INFO_CNDC_RPF/       ← espejo de SharePoint 01_INFO CNDC_RPF
│   ├── 02_DATOS_CNDC_RPF/      ← espejo de SharePoint 02_DATOS CNDC_RPF
│   └── DATOS_EXTRAIDOS/        ← archivos de mapeo loc_names_*.xlsx
└── codigo/                     ← repo GitHub (git clone manual)
```

## Comandos útiles en el servidor

```bash
# Sync manual (dry-run primero)
python3 /srv/rpf/sync/sp_sync_daemon.py --dry-run
python3 /srv/rpf/sync/sp_sync_daemon.py

# Sync solo una carpeta
python3 /srv/rpf/sync/sp_sync_daemon.py --folder 02

# Ver estado de servicios
systemctl status rpf-sync.timer
systemctl status rpf-watcher.service
systemctl list-timers rpf-sync.timer

# Ver logs en tiempo real
tail -f /srv/rpf/sync/logs/sync.log
journalctl -u rpf-sync -f

# Activar/desactivar servicios
sudo systemctl enable --now rpf-sync.timer
sudo systemctl enable --now rpf-watcher.service
sudo systemctl stop rpf-watcher.service

# Verificar PostgreSQL
psql -d rpf_intelligence -c "SELECT * FROM v_rpf_sync_history LIMIT 5;"
psql -d rpf_intelligence -c "SELECT * FROM v_rpf_pending_files LIMIT 10;"
```

## Verificación end-to-end

1. `python3 sp_sync_daemon.py --dry-run` → lista archivos SP sin descargar
2. `python3 sp_sync_daemon.py` → primer sync real (puede tardar varios minutos)
3. `ls /srv/rpf/datos/01_INFO_CNDC_RPF/` → debe mostrar semestres
4. `systemctl list-timers rpf-sync.timer` → próxima ejecución en ~15 min
5. `touch /srv/rpf/datos/test.xlsx` → n8n debe recibir webhook
6. `psql -d rpf_intelligence -c "SELECT count(*) FROM rpf_file_log;"` → archivos registrados

## Notas importantes

- `sharepoint_client_standalone.py` es idéntico al original excepto:
  - `_sp_password()` lee de `os.environ["SP_PASSWORD"]` (no Streamlit Secrets)
  - Sin decoradores `@st.cache_data`
- El daemon **nunca elimina** archivos locales (solo descarga)
- Si SharePoint cambia la sesión, el retry automático en `_get_session()` la renueva
- Los logs rotan automáticamente cuando superan ~10MB (configurado en logrotate)
