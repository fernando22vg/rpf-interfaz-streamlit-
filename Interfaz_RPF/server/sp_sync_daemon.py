#!/usr/bin/env python3
"""
sp_sync_daemon.py — Sincronización incremental SharePoint → /srv/rpf/datos/

Sincroniza 3 carpetas de SharePoint al servidor Ubuntu de forma incremental:
  1. 01_INFO CNDC_RPF/           → /srv/rpf/datos/01_INFO_CNDC_RPF/
  2. 02_DATOS CNDC_RPF/{año}/    → /srv/rpf/datos/02_DATOS_CNDC_RPF/
  3. DATOS EXTRAIDOS.../loc_name → /srv/rpf/datos/DATOS_EXTRAIDOS/

Solo descarga archivos nuevos o con tamaño distinto al local.
Al finalizar notifica a n8n vía webhook POST.

Uso:
    python3 sp_sync_daemon.py             # sync normal
    python3 sp_sync_daemon.py --dry-run   # solo lista cambios sin descargar
    python3 sp_sync_daemon.py --folder 01 # sync solo carpeta 01_INFO

Variables de entorno (en .env):
    SP_PASSWORD   — contraseña del enlace compartido de SharePoint
    N8N_WEBHOOK   — URL del webhook n8n (opcional)
    SYNC_LOG      — ruta al archivo de log (opcional)
    DATA_ROOT     — ruta raíz de datos en el servidor (default: /srv/rpf/datos)
"""

import argparse
import json
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import requests

# ── Cargar .env (antes de importar sharepoint_client_standalone) ──────────────
_HERE = Path(__file__).parent
_ENV_FILE = _HERE / ".env"

def _load_dotenv(path: Path):
    """Mini cargador de .env sin dependencia externa."""
    if not path.exists():
        return
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key not in os.environ:   # no sobreescribir si ya está en el env
                os.environ[key] = val

_load_dotenv(_ENV_FILE)

# ── Configuración desde entorno ───────────────────────────────────────────────
DATA_ROOT   = Path(os.environ.get("DATA_ROOT", "/srv/rpf/datos"))
N8N_WEBHOOK = os.environ.get("N8N_WEBHOOK", "")
SYNC_LOG    = os.environ.get("SYNC_LOG", str(_HERE / "logs" / "sync.log"))

# ── Logging ───────────────────────────────────────────────────────────────────
_log_path = Path(SYNC_LOG)
_log_path.parent.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_log_path, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("rpf-sync")

# ── Importar cliente de SharePoint (standalone, sin Streamlit) ────────────────
sys.path.insert(0, str(_HERE))
from sharepoint_client_standalone import (
    _get_session,
    _list_files,
    _list_folders,
    _download_file,
    _sp_path,
    _INNER_FOLDER,
    _LOC_FOLDER,
    _DATOS_FOLDER,
)

# ── Estructura de directorios locales ─────────────────────────────────────────
LOCAL_01_INFO  = DATA_ROOT / "01_INFO_CNDC_RPF"
LOCAL_02_DATOS = DATA_ROOT / "02_DATOS_CNDC_RPF"
LOCAL_EXTRAIDOS = DATA_ROOT / "DATOS_EXTRAIDOS"


# ── Sync incremental ──────────────────────────────────────────────────────────

def sync_folder_recursive(
    session,
    site_url: str,
    sp_path: str,
    local_path: Path,
    dry_run: bool = False,
    max_depth: int = 10,
    _depth: int = 0,
) -> dict:
    """
    Sincroniza sp_path → local_path de forma incremental.
    Devuelve estadísticas: {downloaded, skipped, errors, bytes_downloaded}
    """
    stats = {"downloaded": 0, "skipped": 0, "errors": 0, "bytes_downloaded": 0}

    if _depth > max_depth:
        log.warning(f"Profundidad máxima alcanzada en: {sp_path}")
        return stats

    local_path.mkdir(parents=True, exist_ok=True)

    # ── Sincronizar archivos en este nivel ────────────────────────────────────
    try:
        files = _list_files(session, site_url, sp_path)
    except Exception as e:
        log.error(f"Error listando archivos en '{sp_path}': {e}")
        stats["errors"] += 1
        return stats

    for f in files:
        fname = f["Name"]
        local_file = local_path / fname
        sp_size = int(f.get("Length", 0))
        srv_url = f["ServerRelativeUrl"]

        # ¿Archivo ya existe con tamaño correcto?
        if local_file.exists() and local_file.stat().st_size == sp_size:
            stats["skipped"] += 1
            continue

        action = "DRY" if dry_run else "↓"
        size_kb = sp_size / 1024
        log.info(f"[{action}] {local_file.relative_to(DATA_ROOT)}  ({size_kb:.1f} KB)")

        if not dry_run:
            try:
                dl_url = (
                    f"{site_url}/_api/web"
                    f"/GetFileByServerRelativeUrl('{_sp_path(srv_url)}')/$value"
                )
                _download_file(session, dl_url, local_file)
                stats["downloaded"] += 1
                stats["bytes_downloaded"] += sp_size
            except Exception as e:
                log.error(f"Error descargando '{fname}': {e}")
                stats["errors"] += 1
                # Eliminar archivo parcial si existe
                if local_file.exists():
                    local_file.unlink(missing_ok=True)
        else:
            stats["downloaded"] += 1  # cuenta como "would download" en dry-run

    # ── Recurrir en subcarpetas ───────────────────────────────────────────────
    try:
        folders = _list_folders(session, site_url, sp_path)
    except Exception as e:
        log.error(f"Error listando carpetas en '{sp_path}': {e}")
        stats["errors"] += 1
        return stats

    for folder in folders:
        sub_stats = sync_folder_recursive(
            session, site_url,
            f"{sp_path}/{folder['Name']}",
            local_path / folder["Name"],
            dry_run=dry_run,
            max_depth=max_depth,
            _depth=_depth + 1,
        )
        for k in stats:
            stats[k] += sub_stats[k]

    return stats


def sync_01_info(session, site_url: str, root_path: str, dry_run: bool) -> dict:
    """Sincroniza 01_INFO CNDC_RPF → /srv/rpf/datos/01_INFO_CNDC_RPF/"""
    sp_path = f"{root_path}/{_INNER_FOLDER}"
    log.info(f"=== Sincronizando 01_INFO CNDC_RPF → {LOCAL_01_INFO} ===")
    return sync_folder_recursive(session, site_url, sp_path, LOCAL_01_INFO, dry_run)


def sync_02_datos(session, site_url: str, root_path: str, dry_run: bool) -> dict:
    """Sincroniza 02_DATOS CNDC_RPF/{año actual y anterior} → /srv/rpf/datos/02_DATOS_CNDC_RPF/"""
    sp_base = f"{root_path}/{_DATOS_FOLDER}"
    stats_total = {"downloaded": 0, "skipped": 0, "errors": 0, "bytes_downloaded": 0}

    # Listar años disponibles y sincronizar los 2 más recientes
    try:
        year_folders = _list_folders(session, site_url, sp_base)
    except Exception as e:
        log.error(f"Error listando años en 02_DATOS CNDC_RPF: {e}")
        stats_total["errors"] += 1
        return stats_total

    current_year = datetime.now().year
    # Ordenar años por cercanía al año actual (priorizar más recientes)
    year_names = sorted(
        [f["Name"] for f in year_folders if f["Name"].isdigit()],
        key=lambda n: abs(int(n) - current_year)
    )
    years_to_sync = year_names[:2]  # año actual + año anterior

    for year in years_to_sync:
        sp_year_path = f"{sp_base}/{year}"
        local_year = LOCAL_02_DATOS / year
        log.info(f"=== Sincronizando 02_DATOS CNDC_RPF/{year} → {local_year} ===")
        stats = sync_folder_recursive(session, site_url, sp_year_path, local_year, dry_run)
        for k in stats_total:
            stats_total[k] += stats[k]

    return stats_total


def sync_extraidos(session, site_url: str, root_path: str, dry_run: bool) -> dict:
    """Sincroniza DATOS EXTRAIDOS.../loc_name → /srv/rpf/datos/DATOS_EXTRAIDOS/"""
    sp_path = f"{root_path}/{_LOC_FOLDER}"
    log.info(f"=== Sincronizando DATOS EXTRAIDOS → {LOCAL_EXTRAIDOS} ===")
    return sync_folder_recursive(session, site_url, sp_path, LOCAL_EXTRAIDOS, dry_run)


# ── Notificación a n8n ────────────────────────────────────────────────────────

def notify_n8n(stats_by_folder: dict, dry_run: bool):
    """Envía resumen de sync a n8n webhook."""
    if not N8N_WEBHOOK:
        return
    payload = {
        "event": "rpf_sync_done",
        "timestamp": datetime.now().isoformat(),
        "dry_run": dry_run,
        "stats": stats_by_folder,
        "total_downloaded": sum(s["downloaded"] for s in stats_by_folder.values()),
        "total_errors": sum(s["errors"] for s in stats_by_folder.values()),
        "total_bytes": sum(s["bytes_downloaded"] for s in stats_by_folder.values()),
    }
    try:
        r = requests.post(N8N_WEBHOOK, json=payload, timeout=10)
        log.info(f"n8n notificado: HTTP {r.status_code}")
    except Exception as e:
        log.warning(f"No se pudo notificar a n8n: {e}")


# ── Punto de entrada ──────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="RPF SharePoint Sync Daemon")
    parser.add_argument("--dry-run", action="store_true",
                        help="Lista cambios sin descargar")
    parser.add_argument("--folder", choices=["01", "02", "extraidos", "all"],
                        default="all", help="Carpeta a sincronizar (default: all)")
    args = parser.parse_args()

    if not os.environ.get("SP_PASSWORD"):
        log.error("SP_PASSWORD no está configurado. Revisa /srv/rpf/sync/.env")
        sys.exit(1)

    dry_label = " [DRY-RUN]" if args.dry_run else ""
    log.info(f"{'='*60}")
    log.info(f"RPF Sync iniciado{dry_label}  —  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info(f"DATA_ROOT: {DATA_ROOT}")

    t_start = time.time()
    stats_total = {}

    try:
        session, site_url, root_path = _get_session()
        log.info(f"Sesión SharePoint OK — site: {site_url}")

        if args.folder in ("01", "all"):
            stats_total["01_INFO"] = sync_01_info(session, site_url, root_path, args.dry_run)

        if args.folder in ("02", "all"):
            stats_total["02_DATOS"] = sync_02_datos(session, site_url, root_path, args.dry_run)

        if args.folder in ("extraidos", "all"):
            stats_total["EXTRAIDOS"] = sync_extraidos(session, site_url, root_path, args.dry_run)

    except Exception as e:
        log.error(f"Error fatal durante sync: {e}", exc_info=True)
        sys.exit(1)

    # ── Resumen ───────────────────────────────────────────────────────────────
    elapsed = time.time() - t_start
    total_dl  = sum(s["downloaded"] for s in stats_total.values())
    total_sk  = sum(s["skipped"]    for s in stats_total.values())
    total_err = sum(s["errors"]     for s in stats_total.values())
    total_mb  = sum(s["bytes_downloaded"] for s in stats_total.values()) / (1024 ** 2)

    log.info(f"{'─'*60}")
    log.info(f"Sync completado en {elapsed:.1f}s{dry_label}")
    log.info(f"  Descargados : {total_dl}")
    log.info(f"  Sin cambios : {total_sk}")
    log.info(f"  Errores     : {total_err}")
    log.info(f"  MB bajados  : {total_mb:.2f}")

    if not args.dry_run:
        notify_n8n(stats_total, args.dry_run)

    sys.exit(1 if total_err > 0 else 0)


if __name__ == "__main__":
    main()
