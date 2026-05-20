"""
sync_to_sharepoint.py
─────────────────────
Sincroniza carpetas locales → SharePoint bajo demanda.

Ejecución:
    python sync_to_sharepoint.py               # sincroniza todo
    python sync_to_sharepoint.py --dry-run     # muestra qué subiría sin subir
    python sync_to_sharepoint.py --carpeta RPF # sincroniza solo la carpeta "RPF"
    python sync_to_sharepoint.py --forzar      # re-sube aunque no haya cambios

Configuración:
    Editar sync_config.json para agregar/quitar carpetas.
    Las credenciales van en .streamlit/secrets.toml (nunca en este archivo).
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE CARPETAS A SINCRONIZAR
# Editar sync_config.json para personalizar. Este script lo crea si no existe.
# ─────────────────────────────────────────────────────────────────────────────
_CONFIG_FILE    = Path(__file__).parent / "sync_config.json"
_MANIFEST_FILE  = Path(__file__).parent / ".sync_manifest.json"  # registro de lo ya subido

# Extensiones de archivo a sincronizar
DEFAULT_EXTENSIONS = [".xlsx", ".xls", ".json", ".csv", ".txt", ".pdf", ".png", ".jpg"]

# Prefijos de archivo a ignorar (temporales de Excel, bloqueos, etc.)
IGNORE_PREFIXES = ["~$", ".tmp", ".lock", "._", "desktop.ini", "thumbs.db"]

# Config por defecto — se crea en sync_config.json si no existe
DEFAULT_CONFIG = {
    "_comentario": "Edite este archivo para agregar o quitar carpetas. No ponga contraseñas aquí.",
    "extensiones": DEFAULT_EXTENSIONS,
    "carpetas": [
        {
            "nombre":       "RPF — Info CNDC",
            "activa":       True,
            "local":        r"C:\Datos del CNDC\01_INFO CNDC_RPF",
            "sp_subcarpeta": "01_INFO CNDC_RPF",
            "comentario":   "Semestres, análisis y resultados de eventos"
        },
        {
            "nombre":       "RPF — Datos CNDC",
            "activa":       True,
            "local":        r"C:\Datos del CNDC\02_DATOS CNDC_RPF",
            "sp_subcarpeta": "02_DATOS CNDC_RPF",
            "comentario":   "Archivos SCADA y EMF por año"
        },
        {
            "nombre":       "DIgSILENT — Mapeos",
            "activa":       True,
            "local":        r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Designacion de loc_name",
            "sp_subcarpeta": "DATOS EXTRAIDOS DE DIGSILENT/Designacion de loc_name",
            "comentario":   "loc_names_gen.xlsx y archivos de mapeo PowerFactory"
        },
        {
            "nombre":       "Ejemplo — carpeta futura",
            "activa":       False,
            "local":        r"C:\ruta\a\otra\carpeta",
            "sp_subcarpeta": "nombre_en_sharepoint",
            "comentario":   "Poner activa: true para activar esta carpeta"
        }
    ]
}


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────────────────

def _load_config() -> dict:
    """Carga sync_config.json. Lo crea con valores por defecto si no existe."""
    if not _CONFIG_FILE.exists():
        with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        print(f"[INFO] Creado archivo de configuración: {_CONFIG_FILE}")
        print("[INFO] Edítelo para personalizar las carpetas antes de volver a ejecutar.\n")
    with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_manifest() -> dict:
    """Carga el registro de archivos ya sincronizados {ruta_absoluta: mtime}."""
    if _MANIFEST_FILE.exists():
        try:
            with open(_MANIFEST_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_manifest(manifest: dict):
    with open(_MANIFEST_FILE, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)


def _should_sync(path: Path, ext_filter: list[str]) -> bool:
    """Devuelve True si el archivo debe sincronizarse."""
    if path.suffix.lower() not in ext_filter:
        return False
    name_lower = path.name.lower()
    if any(name_lower.startswith(p.lower()) for p in IGNORE_PREFIXES):
        return False
    return True


def _mtime(path: Path) -> float:
    """Tiempo de modificación del archivo."""
    return path.stat().st_mtime


def _sp_folder_for(local_file: Path, local_root: Path, sp_root_path: str) -> str:
    """
    Convierte la carpeta de un archivo local a su ruta SP equivalente.
    local_file:  C:\\RPF\\2025 sem1\\Evento 1\\resultado.xlsx
    local_root:  C:\\RPF
    sp_root_path: /personal/user/Documents/Shared/01_INFO CNDC_RPF
    → /personal/user/Documents/Shared/01_INFO CNDC_RPF/2025 sem1/Evento 1
    """
    rel = local_file.parent.relative_to(local_root).as_posix()
    return f"{sp_root_path}/{rel}" if rel != "." else sp_root_path


# ─────────────────────────────────────────────────────────────────────────────
# LÓGICA DE SINCRONIZACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def sync_carpeta(carpeta: dict, sp_raiz_path: str, ext_filter: list[str],
                 manifest: dict, dry_run: bool, forzar: bool) -> dict:
    """
    Sincroniza una carpeta local con su equivalente en SharePoint.
    Devuelve estadísticas: {subidos, omitidos, errores}.
    """
    import sharepoint_client as _sp  # importación tardía

    local_root = Path(carpeta["local"])
    sp_sub     = carpeta["sp_subcarpeta"].strip("/")
    nombre     = carpeta["nombre"]

    if not local_root.exists():
        print(f"  ⚠  Carpeta no encontrada, omitida: {local_root}")
        return {"subidos": 0, "omitidos": 0, "errores": 1}

    # Ruta SP completa para esta carpeta raíz
    sp_folder_raiz = f"{sp_raiz_path}/{sp_sub}"

    stats = {"subidos": 0, "omitidos": 0, "errores": 0}
    archivos = [p for p in local_root.rglob("*") if p.is_file() and _should_sync(p, ext_filter)]

    print(f"\n  📁 {nombre}  ({len(archivos)} archivos candidatos)")
    print(f"     Local : {local_root}")
    print(f"     SP    : {sp_folder_raiz}")

    for archivo in sorted(archivos):
        clave  = str(archivo)
        mtime  = _mtime(archivo)
        ya_ok  = (manifest.get(clave) == mtime) and not forzar

        if ya_ok:
            stats["omitidos"] += 1
            continue

        sp_folder = _sp_folder_for(archivo, local_root, sp_folder_raiz)
        rel_str   = archivo.relative_to(local_root)

        if dry_run:
            print(f"     [DRY]  → {rel_str}")
            stats["subidos"] += 1
            continue

        try:
            _sp.ensure_sp_folder(sp_folder)
            _sp.upload_file(str(archivo), sp_folder)
            manifest[clave] = mtime
            stats["subidos"] += 1
            print(f"     ✔  {rel_str}")
        except Exception as exc:
            stats["errores"] += 1
            print(f"     ✘  {rel_str}  —  {exc}")

    return stats


def run(filtro_nombre: str | None = None, dry_run: bool = False, forzar: bool = False):
    """Punto de entrada principal."""
    print("=" * 62)
    print("  SYNC LOCAL → SHAREPOINT")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    if dry_run:
        print("  [MODO DRY-RUN — no se sube nada]")
    if forzar:
        print("  [MODO FORZAR — re-sube todos los archivos]")
    print("=" * 62)

    # ── Cargar configuración ──────────────────────────────────────────────────
    cfg      = _load_config()
    ext_list = cfg.get("extensiones", DEFAULT_EXTENSIONS)
    carpetas = [c for c in cfg["carpetas"] if c.get("activa", False)]

    if filtro_nombre:
        carpetas = [c for c in carpetas if filtro_nombre.lower() in c["nombre"].lower()]
        if not carpetas:
            print(f"[ERROR] No se encontró ninguna carpeta activa con nombre '{filtro_nombre}'")
            sys.exit(1)

    if not carpetas:
        print("[INFO] No hay carpetas activas en sync_config.json.")
        print("       Edite el archivo y ponga 'activa': true en las que quiera sincronizar.")
        sys.exit(0)

    # ── Conectar a SharePoint ─────────────────────────────────────────────────
    print("\nConectando a SharePoint…")
    try:
        # Necesitamos st.secrets — si no hay Streamlit, leer secrets.toml directo
        _inject_secrets()
        import sharepoint_client as _sp
        session, site_url, root_path = _sp._get_session()
        # La raíz SP es root_path (carpeta compartida), de donde cuelga todo
        sp_raiz = root_path
        print(f"  Conectado ✔  ({site_url})")
    except Exception as exc:
        print(f"  [ERROR] No se pudo conectar a SharePoint:\n  {exc}")
        sys.exit(1)

    # ── Sincronizar cada carpeta ──────────────────────────────────────────────
    manifest   = _load_manifest()
    totales    = {"subidos": 0, "omitidos": 0, "errores": 0}
    t_inicio   = time.time()

    for carpeta in carpetas:
        s = sync_carpeta(carpeta, sp_raiz, ext_list, manifest, dry_run, forzar)
        for k in totales:
            totales[k] += s[k]

    if not dry_run:
        _save_manifest(manifest)

    # ── Resumen ───────────────────────────────────────────────────────────────
    elapsed = time.time() - t_inicio
    print("\n" + "=" * 62)
    print(f"  RESUMEN — {elapsed:.1f} s")
    print(f"  ✔ Subidos  : {totales['subidos']}")
    print(f"  ⏭ Omitidos : {totales['omitidos']}  (sin cambios desde último sync)")
    print(f"  ✘ Errores  : {totales['errores']}")
    print("=" * 62)

    if totales["errores"]:
        sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# INYECCIÓN DE SECRETS (para correr fuera de Streamlit)
# ─────────────────────────────────────────────────────────────────────────────

def _inject_secrets():
    """
    Cuando el script corre fuera de Streamlit, st.secrets no existe.
    Lee .streamlit/secrets.toml manualmente y lo inyecta como atributo
    para que sharepoint_client._sp_password() lo encuentre.
    """
    try:
        import streamlit as st
        # Probar si st.secrets ya funciona
        _ = st.secrets
        return  # Streamlit ya maneja los secrets
    except Exception:
        pass

    # Leer secrets.toml manualmente (formato TOML simplificado)
    secrets_file = Path(__file__).parent.parent / ".streamlit" / "secrets.toml"
    if not secrets_file.exists():
        secrets_file = Path(__file__).parent / ".streamlit" / "secrets.toml"

    secrets = {}
    if secrets_file.exists():
        try:
            import re
            text = secrets_file.read_text(encoding="utf-8")
            for m in re.finditer(r'^(\w+)\s*=\s*["\'](.+?)["\']', text, re.MULTILINE):
                secrets[m.group(1)] = m.group(2)
        except Exception:
            pass

    if not secrets:
        print("[WARN] No se encontró .streamlit/secrets.toml — SharePoint puede rechazar la conexión.")
        return

    # Parchear st.secrets con un dict-like simple
    import types
    import streamlit as st

    class _FakeSecrets(dict):
        def get(self, key, default=None):
            return super().get(key, default)

    try:
        st.secrets = _FakeSecrets(secrets)
    except Exception:
        # Si st no permite asignación, monkeypatch el módulo
        sys.modules["streamlit"].secrets = _FakeSecrets(secrets)


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Sincroniza carpetas locales con SharePoint.",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="""
Ejemplos:
  python sync_to_sharepoint.py                    # sincroniza todas las carpetas activas
  python sync_to_sharepoint.py --dry-run          # solo muestra qué se subiría
  python sync_to_sharepoint.py --carpeta RPF      # solo carpetas cuyo nombre contenga 'RPF'
  python sync_to_sharepoint.py --forzar           # re-sube aunque no haya cambios

Para agregar carpetas: edite sync_config.json
Para cambiar la contraseña: edite .streamlit/secrets.toml
        """
    )
    parser.add_argument("--dry-run",  action="store_true",
                        help="Muestra qué se subiría sin subir nada")
    parser.add_argument("--forzar",   action="store_true",
                        help="Re-sube todos los archivos aunque no hayan cambiado")
    parser.add_argument("--carpeta",  metavar="NOMBRE",
                        help="Filtra por nombre de carpeta en sync_config.json")
    args = parser.parse_args()

    run(filtro_nombre=args.carpeta, dry_run=args.dry_run, forzar=args.forzar)
