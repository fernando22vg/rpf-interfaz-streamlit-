"""
sharepoint_client.py  (acceso via cookies de sesión, sin Azure App Registration)

Usa el sharing link de SharePoint para inicializar una sesión con cookies
que permiten llamar al REST API de SharePoint sin credenciales explícitas.
Funciona para carpetas compartidas como "Cualquiera con el enlace puede ver".
"""

import shutil
import tempfile
import threading
import time
import re
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote
from datetime import datetime

import requests
import streamlit as st

SHARE_URL = (
    "https://cobee1-my.sharepoint.com/:f:/g/personal/angel_mariscal_cobee_com"
    "/IgDQ0-3WNNN1SYksWDQKnGTeAdQNzcw0KrsBYeBuI7_NAf0?e=SBurQb"
)

# Subcarpeta dentro de la raíz compartida donde están los semestres
# Equivalente al RAIZ local: C:\Datos del CNDC\01_INFO CNDC_RPF
_INNER_FOLDER = "01_INFO CNDC_RPF"

# Otras carpetas en la raíz compartida (misma estructura que local)
_LOC_FOLDER   = "DATOS EXTRAIDOS DE DIGSILENT/Designacion de loc_name"
_DATOS_FOLDER = "02_DATOS CNDC_RPF"

_TMP_ROOT = Path(tempfile.gettempdir()) / "rpf_sharepoint"
_session_cache: dict = {}
_session_lock = threading.Lock()

_HEADERS_BROWSE = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
}
_HEADERS_API = {"Accept": "application/json;odata=nometadata"}


# ── Inicialización de sesión ──────────────────────────────────────────────────

def _get_session() -> tuple:
    """
    Retorna (session, site_url, root_folder_server_relative_path).
    Accede al sharing link para obtener cookies de sesión y la ruta raíz.
    """
    with _session_lock:
        if _session_cache.get("expires_at", 0) > time.time():
            return (
                _session_cache["session"],
                _session_cache["site_url"],
                _session_cache["root_path"],
            )

        session = requests.Session()
        session.headers.update(_HEADERS_BROWSE)

        r = session.get(SHARE_URL, timeout=30)
        r.raise_for_status()

        final_url = r.url
        parsed = urlparse(final_url)
        qs = parse_qs(parsed.query)

        # Estrategia 1: parámetro "id" en la URL de redirección
        folder_path = unquote(qs.get("id", [""])[0])

        # Estrategia 2: parámetro "RootFolder" o "listurl"
        if not folder_path:
            folder_path = unquote(qs.get("RootFolder", [""])[0])

        if not folder_path:
            raise RuntimeError(
                "SharePoint no devolvió la ruta de la carpeta en el redirect.\n"
                f"URL final: {final_url}\n"
                f"Parámetros detectados: {list(qs.keys())}\n"
                "Verifica que el enlace esté activo y sea 'Cualquiera con el enlace'."
            )

        scheme_host = f"{parsed.scheme}://{parsed.netloc}"
        if "/personal/" in folder_path:
            user_slug = folder_path.split("/personal/")[1].split("/")[0]
            site_url = f"{scheme_host}/personal/{user_slug}"
        else:
            site_url = scheme_host

        _session_cache.update({
            "session": session,
            "site_url": site_url,
            "root_path": folder_path,
            "expires_at": time.time() + 1800,
        })

        return session, site_url, folder_path


def _clear_session():
    with _session_lock:
        _session_cache.clear()


# ── SharePoint REST API ───────────────────────────────────────────────────────

def _sp_api(session, site_url: str, endpoint: str) -> dict:
    """Llama al SharePoint REST API con la sesión (cookies) actual."""
    r = session.get(
        f"{site_url}/_api/{endpoint}",
        headers=_HEADERS_API,
        timeout=20,
    )
    if r.status_code in (401, 403):
        # Sesión expirada: reinicializar una vez
        _clear_session()
        session, site_url, _ = _get_session()
        r = session.get(
            f"{site_url}/_api/{endpoint}",
            headers=_HEADERS_API,
            timeout=20,
        )
    if r.status_code in (401, 403):
        raise PermissionError(
            "SharePoint rechazó el acceso al REST API.\n"
            "El truco de cookies no funciona con la configuración de este tenant.\n"
            "Opciones: (1) App Registration en Azure, (2) Google Drive."
        )
    r.raise_for_status()
    return r.json()


def _sp_path(path: str) -> str:
    """Escapa comillas simples para rutas OData de SharePoint."""
    return path.replace("'", "''")


def _list_folders(session, site_url: str, sp_path: str) -> list:
    data = _sp_api(
        session, site_url,
        f"web/GetFolderByServerRelativeUrl('{_sp_path(sp_path)}')/Folders"
    )
    return data.get("value", [])


def _list_files(session, site_url: str, sp_path: str) -> list:
    data = _sp_api(
        session, site_url,
        f"web/GetFolderByServerRelativeUrl('{_sp_path(sp_path)}')/Files"
    )
    return data.get("value", [])


# ── API pública ───────────────────────────────────────────────────────────────

def _raiz_path() -> tuple:
    """Devuelve (session, site_url, raiz_path) donde raiz_path apunta a 01_INFO CNDC_RPF."""
    session, site_url, root_path = _get_session()
    raiz = f"{root_path}/{_INNER_FOLDER}"
    return session, site_url, raiz


@st.cache_data(ttl=300, show_spinner=False)
def listar_semestres() -> list:
    session, site_url, raiz = _raiz_path()
    folders = _list_folders(session, site_url, raiz)
    return sorted(f["Name"] for f in folders)


@st.cache_data(ttl=300, show_spinner=False)
def listar_eventos(semestre: str) -> list:
    session, site_url, raiz = _raiz_path()
    analisis_path = f"{raiz}/{semestre}/Análisis_todos_los_eventos"
    folders = _list_folders(session, site_url, analisis_path)
    return sorted(f["Name"] for f in folders)


# Rutas temporales equivalentes a las carpetas locales
TMP_RAIZ       = _TMP_ROOT                                    # ≡ C:\Datos del CNDC\01_INFO CNDC_RPF
TMP_LOC_FOLDER = _TMP_ROOT.parent / _LOC_FOLDER              # ≡ C:\Datos del CNDC\DATOS EXTRAIDOS...
TMP_DATOS      = _TMP_ROOT.parent / _DATOS_FOLDER            # ≡ C:\Datos del CNDC\02_DATOS CNDC_RPF


def descargar_archivos_estaticos():
    """
    Descarga una sola vez los archivos de mapeo (loc_names_*.xlsx) desde SharePoint.
    Estos archivos no cambian con cada evento.
    """
    dest = TMP_LOC_FOLDER
    if dest.exists() and any(dest.glob("*.xlsx")):
        return  # ya descargados

    session, site_url, root_path = _get_session()
    sp_loc_path = f"{root_path}/{_LOC_FOLDER}"
    dest.mkdir(parents=True, exist_ok=True)

    for f in _list_files(session, site_url, sp_loc_path):
        local_file = dest / f["Name"]
        if not local_file.exists():
            srv_url = f["ServerRelativeUrl"]
            dl_url = (
                f"{site_url}/_api/web"
                f"/GetFileByServerRelativeUrl('{_sp_path(srv_url)}')/$value"
            )
            _download_file(session, dl_url, local_file)


def descargar_evento(semestre: str, evento: str, progress_cb=None) -> Path:
    """
    Descarga la carpeta del evento a /tmp/ usando la sesión de SharePoint.
    También descarga los archivos directos del semestre (Tabla_Eventos_*.xlsx, etc.).
    Devuelve la ruta local equivalente a ev_path.
    """
    local_path = _TMP_ROOT / semestre / evento
    session, site_url, raiz = _raiz_path()

    # Descargar archivos del nivel semestre (p.ej. Tabla_Eventos_*.xlsx)
    sem_local = _TMP_ROOT / semestre
    sem_local.mkdir(parents=True, exist_ok=True)
    sp_sem_path = f"{raiz}/{semestre}"
    for f in _list_files(session, site_url, sp_sem_path):
        dest_f = sem_local / f["Name"]
        if not dest_f.exists():
            srv_url = f["ServerRelativeUrl"]
            dl_url = (
                f"{site_url}/_api/web"
                f"/GetFileByServerRelativeUrl('{_sp_path(srv_url)}')/$value"
            )
            _download_file(session, dl_url, dest_f)

    # Descargar carpeta del evento
    if not local_path.exists():
        sp_event_path = f"{raiz}/{semestre}/Análisis_todos_los_eventos/{evento}"
        _download_tree(session, site_url, sp_event_path, local_path, progress_cb)

    return local_path


_PAT_FALLA = re.compile(
    r'FALLA\s+(\d{2})\.(\d{2})\.(\d{2,4})\s+HRS\s*(\d{2})\.(\d{2})',
    re.IGNORECASE,
)


def descargar_scada_falla(fecha_evento) -> Path:
    """
    Busca y descarga desde SharePoint la carpeta FALLA que corresponde a
    fecha_evento: 02_DATOS CNDC_RPF/{año}/FALLA DD.MM.YYYY HRS HH.MM/

    Devuelve la ruta local bajo TMP_DATOS (misma estructura que RAIZ_DATOS local).
    """
    session, site_url, root_path = _get_session()
    yr4 = fecha_evento.year
    yr2 = yr4 % 100
    dd, mm = fecha_evento.day, fecha_evento.month
    hh, mi = fecha_evento.hour, fecha_evento.minute

    datos_base = f"{root_path}/{_DATOS_FOLDER}"

    def _buscar_en_year(year_sp_path):
        try:
            folders = _list_folders(session, site_url, year_sp_path)
        except Exception:
            return None
        candidatos = []
        for f in folders:
            m = _PAT_FALLA.match(f["Name"])
            if not m:
                continue
            fd_s, fm_s, fy_s, fh_s, fmi_s = m.groups()
            fd, fm, fh, fmi = int(fd_s), int(fm_s), int(fh_s), int(fmi_s)
            fy = int(fy_s) % 100
            if fd == dd and fm == mm and fy == yr2:
                diff = abs(fh * 60 + fmi - hh * 60 - mi)
                candidatos.append((diff, f["Name"]))
        if not candidatos:
            return None
        candidatos.sort()
        return candidatos[0][1]

    # Intentar con el año exacto primero
    falla_name = _buscar_en_year(f"{datos_base}/{yr4}")
    used_yr4 = yr4

    # Fallback: recorrer todos los años disponibles
    if falla_name is None:
        try:
            top = _list_folders(session, site_url, datos_base)
        except Exception as exc:
            raise FileNotFoundError(
                f"No se pudo acceder a '{_DATOS_FOLDER}' en SharePoint: {exc}"
            ) from exc
        for tf in sorted(top, key=lambda f: f["Name"]):
            if tf["Name"].isdigit() and int(tf["Name"]) != yr4:
                candidate = _buscar_en_year(f"{datos_base}/{tf['Name']}")
                if candidate:
                    falla_name = candidate
                    used_yr4 = int(tf["Name"])
                    break

    if falla_name is None:
        raise FileNotFoundError(
            f"No se encontró carpeta FALLA para {fecha_evento.strftime('%d.%m.%Y')} "
            f"bajo SharePoint/{_DATOS_FOLDER}"
        )

    falla_sp_path = f"{datos_base}/{used_yr4}/{falla_name}"
    local_path = TMP_DATOS / str(used_yr4) / falla_name
    if not local_path.exists():
        _download_tree(session, site_url, falla_sp_path, local_path)
    return local_path


def limpiar_cache_evento(semestre: str, evento: str):
    local_path = _TMP_ROOT / semestre / evento
    if local_path.exists():
        shutil.rmtree(local_path)


# ── Descarga recursiva ────────────────────────────────────────────────────────

def _download_tree(session, site_url: str, sp_path: str, dest: Path, progress_cb=None):
    dest.mkdir(parents=True, exist_ok=True)

    for folder in _list_folders(session, site_url, sp_path):
        _download_tree(
            session, site_url,
            f"{sp_path}/{folder['Name']}",
            dest / folder["Name"],
            progress_cb,
        )

    for f in _list_files(session, site_url, sp_path):
        local_file = dest / f["Name"]
        if not local_file.exists():
            srv_url = f["ServerRelativeUrl"]
            dl_url = (
                f"{site_url}/_api/web"
                f"/GetFileByServerRelativeUrl('{_sp_path(srv_url)}')/$value"
            )
            _download_file(session, dl_url, local_file)
        if progress_cb:
            progress_cb(f["Name"])


def _download_file(session, url: str, dest: Path):
    with session.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(65536):
                f.write(chunk)
