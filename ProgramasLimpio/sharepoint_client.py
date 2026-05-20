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


# ── Escritura en SharePoint ───────────────────────────────────────────────────

def _get_request_digest(session, site_url: str) -> str:
    """Obtiene el FormDigest necesario para operaciones de escritura."""
    r = session.post(
        f"{site_url}/_api/contextinfo",
        headers={**_HEADERS_API, "Content-Length": "0"},
        timeout=20,
    )
    r.raise_for_status()
    return r.json().get("FormDigestValue", "")


def _upload_sp_file(session, site_url: str, sp_folder: str, filename: str, content: bytes):
    """Sube (crea o sobreescribe) un archivo en una carpeta SharePoint."""
    digest = _get_request_digest(session, site_url)
    url = (
        f"{site_url}/_api/web"
        f"/GetFolderByServerRelativeUrl('{_sp_path(sp_folder)}')"
        f"/Files/add(overwrite=true,url='{filename.replace(chr(39), chr(39)*2)}')"
    )
    r = session.post(
        url,
        data=content,
        headers={
            **_HEADERS_API,
            "X-RequestDigest": digest,
            "Content-Type": "application/octet-stream",
        },
        timeout=30,
    )
    r.raise_for_status()


def upload_json(sp_folder: str, filename: str, data: dict):
    """
    Sube un dict como JSON a SharePoint.
    sp_folder: ruta relativa al servidor (server-relative path)
    """
    import json as _json
    session, site_url, _ = _get_session()
    content = _json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    _upload_sp_file(session, site_url, sp_folder, filename, content)


def download_json(sp_file_path: str) -> dict:
    """
    Descarga un archivo JSON desde SharePoint y lo devuelve como dict.
    Devuelve {} si no existe o hay error.
    """
    import json as _json
    session, site_url, _ = _get_session()
    url = (
        f"{site_url}/_api/web"
        f"/GetFileByServerRelativeUrl('{_sp_path(sp_file_path)}')/$value"
    )
    try:
        r = session.get(url, headers=_HEADERS_API, timeout=20)
        if r.status_code == 404:
            return {}
        r.raise_for_status()
        return _json.loads(r.content.decode("utf-8"))
    except Exception:
        return {}


def sp_folder_from_local(local_path: str) -> str:
    """
    Convierte una ruta local bajo TMP_RAIZ a su equivalente en SharePoint.
    Ejemplo: /tmp/rpf_sharepoint/2025 sem1/Análisis.../Evento 1
           → {sp_root}/01_INFO CNDC_RPF/2025 sem1/Análisis.../Evento 1
    """
    rel = Path(local_path).relative_to(_TMP_ROOT).as_posix()
    _, _, raiz = _raiz_path()   # raiz = {sp_root}/01_INFO CNDC_RPF
    return f"{raiz}/{rel}"


def local_path_to_sp_folder(local_path: str, raiz_local: str) -> str:
    """
    Mapea una ruta local (bajo raiz_local = RAIZ_RPF) a su carpeta SP equivalente.
    Funciona en modo local (no IS_CLOUD).

    Ejemplo:
      local_path  = C:\\Datos del CNDC\\01_INFO CNDC_RPF\\2025 sem1\\Ev 1\\E1.0\\Datos Curvas
      raiz_local  = C:\\Datos del CNDC\\01_INFO CNDC_RPF
      → {sp_root}/01_INFO CNDC_RPF/2025 sem1/Ev 1/E1.0/Datos Curvas
    """
    p = Path(local_path)
    raiz = Path(raiz_local)
    # Si es un archivo, tomar el directorio padre
    folder = p.parent if p.is_file() else p
    try:
        rel = folder.relative_to(raiz).as_posix()
    except ValueError:
        # Si no está bajo raiz_local, usar solo el nombre de la carpeta
        rel = folder.name
    _, _, raiz_sp = _raiz_path()
    return f"{raiz_sp}/{rel}" if rel != "." else raiz_sp


def upload_file(local_path: str, sp_folder: str):
    """
    Sube un archivo local a una carpeta SharePoint (crea o sobreescribe).
    sp_folder: server-relative path de la carpeta destino en SP.
    """
    p = Path(local_path)
    if not p.is_file():
        raise FileNotFoundError(f"No existe: {local_path}")
    session, site_url, _ = _get_session()
    with open(p, "rb") as fh:
        content = fh.read()
    _upload_sp_file(session, site_url, sp_folder, p.name, content)


def ensure_sp_folder(sp_folder: str):
    """
    Garantiza que una carpeta exista en SharePoint, creando todos los niveles
    necesarios. Ignora si ya existe.
    """
    session, site_url, _ = _get_session()
    # Verificar si existe
    try:
        _sp_api(session, site_url,
                f"web/GetFolderByServerRelativeUrl('{_sp_path(sp_folder)}')")
        return  # ya existe
    except Exception:
        pass

    # Crear nivel a nivel
    parts = sp_folder.strip("/").split("/")
    for i in range(1, len(parts) + 1):
        partial = "/" + "/".join(parts[:i])
        try:
            _sp_api(session, site_url,
                    f"web/GetFolderByServerRelativeUrl('{_sp_path(partial)}')")
        except Exception:
            # Crear esta carpeta
            digest = _get_request_digest(session, site_url)
            parent = "/" + "/".join(parts[:i - 1]) if i > 1 else "/"
            name   = parts[i - 1]
            url    = (f"{site_url}/_api/web"
                      f"/GetFolderByServerRelativeUrl('{_sp_path(parent)}')"
                      f"/Folders/add(url='{name.replace(chr(39), chr(39)*2)}')")
            try:
                session.post(url, data=b"",
                             headers={**_HEADERS_API, "X-RequestDigest": digest,
                                      "Content-Type": "application/octet-stream"},
                             timeout=20)
            except Exception:
                pass  # best-effort


def sp_global_cfg_folder() -> str:
    """Ruta SP de la carpeta donde se guarda unit_global_config.json."""
    _, _, root_path = _get_session()
    return f"{root_path}/{_LOC_FOLDER}"


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
    local_path = _TMP_ROOT / semestre / "Análisis_todos_los_eventos" / evento
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
    r'FALLA\s+(\d{2})[.\-](\d{2})[.\-](\d{2,4})\s+HRS\s*(\d{2})\.(\d{2})',
    re.IGNORECASE,
)

# Formato alternativo: Datos_DDMMYY_HHMM_*.xlsx  (p.ej. Datos_190325_0506_CCERI50.xlsx)
_PAT_DATOS = re.compile(
    r'Datos_(\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})_.*\.xlsx',
    re.IGNORECASE,
)


def descargar_scada_falla(fecha_evento) -> Path:
    """
    Busca y descarga desde SharePoint la carpeta/archivo SCADA para fecha_evento.
    Soporta dos estructuras en 02_DATOS CNDC_RPF/{año}/:
      A) Carpeta  FALLA DD.MM.YYYY HRS HH.MM/  (contiene archivo '1 seg.*')
      B) Archivo  Datos_DDMMYY_HHMM_*.xlsx     (mismo contenido que '1 seg', nombre distinto)

    En el caso B crea una carpeta FALLA temporal y guarda el archivo con
    nombre '1 seg.*.xlsx' para que OrdenadorDatosEvento lo encuentre.

    Devuelve la ruta local de la carpeta FALLA (real o simulada).
    """
    session, site_url, root_path = _get_session()
    yr4 = fecha_evento.year
    yr2 = yr4 % 100
    dd, mm = fecha_evento.day, fecha_evento.month
    hh, mi = fecha_evento.hour, fecha_evento.minute

    datos_base = f"{root_path}/{_DATOS_FOLDER}"

    # ── Buscar carpeta FALLA en un nivel dado ─────────────────────────────────
    def _buscar_falla_en(sp_path):
        try:
            folders = _list_folders(session, site_url, sp_path)
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

    # ── Buscar archivo Datos_DDMMYY_HHMM_*.xlsx en un nivel dado ─────────────
    def _buscar_datos_en(sp_path):
        try:
            files = _list_files(session, site_url, sp_path)
        except Exception:
            return None
        candidatos = []
        for f in files:
            m = _PAT_DATOS.match(f["Name"])
            if not m:
                continue
            fd_s, fm_s, fy_s, fh_s, fmi_s = m.groups()
            fd, fm, fh, fmi = int(fd_s), int(fm_s), int(fh_s), int(fmi_s)
            fy = int(fy_s) % 100
            if fd == dd and fm == mm and fy == yr2:
                diff = abs(fh * 60 + fmi - hh * 60 - mi)
                candidatos.append((diff, fh, fmi, f["Name"], f["ServerRelativeUrl"]))
        if not candidatos:
            return None
        candidatos.sort()
        return candidatos[0]  # (diff, fh, fmi, name, srv_url)

    # ── Listar años disponibles en datos_base ─────────────────────────────────
    try:
        top_folders = _list_folders(session, site_url, datos_base)
    except Exception as exc:
        raise FileNotFoundError(
            f"No se pudo acceder a '{_DATOS_FOLDER}' en SharePoint: {exc}"
        ) from exc
    year_names = sorted(
        (f["Name"] for f in top_folders if f["Name"].isdigit()),
        key=lambda n: abs(int(n) - yr4),   # ordenar por cercanía al año del evento
    )

    # ── Intentar año exacto primero, luego otros años ─────────────────────────
    search_years = [str(yr4)] + [y for y in year_names if y != str(yr4)]

    for year_str in search_years:
        sp_year = f"{datos_base}/{year_str}"

        # Estrategia A: carpeta FALLA
        falla_name = _buscar_falla_en(sp_year)
        if falla_name:
            falla_sp_path = f"{sp_year}/{falla_name}"
            local_path    = TMP_DATOS / year_str / falla_name
            if not local_path.exists():
                _download_tree(session, site_url, falla_sp_path, local_path)
            return local_path

        # Estrategia B: archivo Datos_DDMMYY_HHMM_*.xlsx
        datos_hit = _buscar_datos_en(sp_year)
        if datos_hit:
            _, fh, fmi, datos_name, srv_url = datos_hit
            # Crear carpeta FALLA sintética en local
            falla_nombre = f"FALLA {dd:02d}.{mm:02d}.{yr4} HRS {fh:02d}.{fmi:02d}"
            seg_nombre   = f"1 seg.{dd:02d}.{mm:02d}.{yr2:02d}_hrs.{fh:02d}.{fmi:02d}.xlsx"
            local_falla  = TMP_DATOS / year_str / falla_nombre
            local_seg    = local_falla / seg_nombre
            if not local_seg.exists():
                local_falla.mkdir(parents=True, exist_ok=True)
                dl_url = (
                    f"{site_url}/_api/web"
                    f"/GetFileByServerRelativeUrl('{_sp_path(srv_url)}')/$value"
                )
                _download_file(session, dl_url, local_seg)
            return local_falla

    raise FileNotFoundError(
        f"No se encontró carpeta FALLA ni archivo Datos_*.xlsx para "
        f"{fecha_evento.strftime('%d.%m.%Y %H:%M')} "
        f"en SharePoint/{_DATOS_FOLDER}. "
        f"Años revisados: {search_years}"
    )


def limpiar_cache_evento(semestre: str, evento: str):
    local_path = _TMP_ROOT / semestre / "Análisis_todos_los_eventos" / evento
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
