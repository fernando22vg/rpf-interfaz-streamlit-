"""
sharepoint_client_standalone.py  — versión sin Streamlit para Ubuntu Server

Copia funcional de sharepoint_client.py con dos cambios mínimos:
  1. _sp_password() lee de os.environ["SP_PASSWORD"] en lugar de st.secrets
  2. Se eliminan los decoradores @st.cache_data (no disponible fuera de Streamlit)

Uso en el servidor:
    export SP_PASSWORD="contraseña_del_link"
    from sharepoint_client_standalone import _get_session, _list_files, ...
"""

import os
import shutil
import tempfile
import threading
import time
import re
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote

import requests

SHARE_URL = (
    "https://cobee1-my.sharepoint.com/:f:/g/personal/angel_mariscal_cobee_com"
    "/IgDxSTKUmkgHSqWE7ujrvUlUASIptnjxSa00s0iqJMYCcCk?e=59xxDd"
)

# ── CAMBIO 1: Lee SP_PASSWORD de variable de entorno (no de Streamlit Secrets)
def _sp_password() -> str:
    return os.environ.get("SP_PASSWORD", "")

# Subcarpeta dentro de la raíz compartida donde están los semestres
_INNER_FOLDER = "01_INFO CNDC_RPF"
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

def _submit_password(session: requests.Session, r: requests.Response) -> requests.Response:
    import html as _html
    from html.parser import HTMLParser as _HP

    class _FormParser(_HP):
        def __init__(self):
            super().__init__()
            self.action  = ""
            self.fields  = {}
            self._in_form = False
        def handle_starttag(self, tag, attrs):
            a = dict(attrs)
            if tag == "form":
                self._in_form = True
                self.action = a.get("action", "")
            if tag == "input" and self._in_form:
                n = a.get("name", "")
                v = a.get("value", "")
                t = a.get("type", "text")
                if n and t != "submit":
                    self.fields[n] = _html.unescape(v)

    parser = _FormParser()
    parser.feed(r.text)

    pwd = _sp_password()
    for key in list(parser.fields.keys()):
        if "password" in key.lower() or "pwd" in key.lower():
            parser.fields[key] = pwd
            break
    else:
        parser.fields["txtPassword"] = pwd
        parser.fields["ctl00$PlaceHolderMain$ctl00$txtPassword"] = pwd

    action = parser.action or r.url
    if action.startswith("/"):
        p = urlparse(r.url)
        action = f"{p.scheme}://{p.netloc}{action}"
    elif not action.startswith("http"):
        action = r.url

    r2 = session.post(
        action,
        data=parser.fields,
        headers={**_HEADERS_BROWSE, "Content-Type": "application/x-www-form-urlencoded",
                 "Referer": r.url},
        timeout=30,
        allow_redirects=True,
    )
    r2.raise_for_status()
    return r2


def _needs_password(r: requests.Response) -> bool:
    """Detecta si la página tiene un campo de contraseña para enviar.
    Más robusto que la versión original: acepta comillas simples y dobles,
    y detecta el campo por name/id además de type.
    """
    txt = r.text.lower()
    return (
        'type="password"'   in txt or
        "type='password'"   in txt or
        'name="txtpassword"' in txt or
        'id="txtpassword"'   in txt or
        'name="password"'    in txt
    )


def _get_session() -> tuple:
    with _session_lock:
        if _session_cache.get("expires_at", 0) > time.time():
            return (
                _session_cache["session"],
                _session_cache["site_url"],
                _session_cache["root_path"],
            )

        # Validar que SP_PASSWORD está configurado antes de intentar conectar
        if not _sp_password():
            raise RuntimeError(
                "SP_PASSWORD no está configurado.\n"
                "Verifica que /srv/rpf/sync/.env contiene: SP_PASSWORD=tu_contraseña\n"
                "y que el daemon carga el .env antes de importar este módulo."
            )

        session = requests.Session()
        session.headers.update(_HEADERS_BROWSE)

        r = session.get(SHARE_URL, timeout=30)
        r.raise_for_status()

        # Intento 1: detectar y enviar contraseña con la lógica normal
        if _needs_password(r):
            r = _submit_password(session, r)

        # Intento 2: si seguimos en guestaccess.aspx, intentar enviar contraseña
        # aunque _needs_password no lo haya detectado (formulario cargado por JS)
        if "guestaccess.aspx" in r.url and _sp_password():
            r2 = _submit_password(session, r)
            # Solo usar r2 si avanzamos (salimos de guestaccess.aspx)
            if "guestaccess.aspx" not in r2.url:
                r = r2

        final_url = r.url
        parsed = urlparse(final_url)
        qs = parse_qs(parsed.query)

        # Estrategia 1: parámetro "id" en la URL de redirección
        folder_path = unquote(qs.get("id", [""])[0])

        # Estrategia 2: parámetro "RootFolder"
        if not folder_path:
            folder_path = unquote(qs.get("RootFolder", [""])[0])

        # Estrategia 3: buscar "ServerRelativeUrl" en el HTML (original)
        if not folder_path:
            m = re.search(r'"ServerRelativeUrl"\s*:\s*"([^"]+)"', r.text, re.IGNORECASE)
            if m:
                folder_path = unquote(m.group(1))

        # Estrategia 4: buscar patrón /personal/USER/... en el HTML
        if not folder_path:
            m = re.search(r'(/personal/[^/\s"\'<>]+(?:/[^/\s"\'<>]+){2,})', r.text)
            if m:
                candidate = unquote(m.group(1))
                # Filtrar rutas de layouts o _api, solo rutas de datos
                if "_layouts" not in candidate and "_api" not in candidate:
                    folder_path = candidate

        if not folder_path:
            pwd_status = "SÍ (len={})".format(len(_sp_password())) if _sp_password() else "NO ← ESTE ES EL PROBLEMA"
            raise RuntimeError(
                "SharePoint no devolvió la ruta de la carpeta.\n"
                f"URL final: {final_url}\n"
                f"Parámetros URL: {list(qs.keys())}\n"
                f"SP_PASSWORD configurado: {pwd_status}\n"
                "Si la contraseña está bien, el formulario del link puede haber cambiado.\n"
                f"Primeros 300 chars del HTML:\n{r.text[:300]}"
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
    r = session.get(
        f"{site_url}/_api/{endpoint}",
        headers=_HEADERS_API,
        timeout=20,
    )
    if r.status_code in (401, 403):
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
            "El truco de cookies no funciona con la configuración de este tenant."
        )
    r.raise_for_status()
    return r.json()


def _sp_path(path: str) -> str:
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


# ── Descarga ──────────────────────────────────────────────────────────────────

def _download_file(session, url: str, dest: Path):
    with session.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(65536):
                f.write(chunk)


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


# ── API pública ───────────────────────────────────────────────────────────────

def _raiz_path() -> tuple:
    session, site_url, root_path = _get_session()
    raiz = f"{root_path}/{_INNER_FOLDER}"
    return session, site_url, raiz


# ── CAMBIO 2: Sin @st.cache_data (no disponible fuera de Streamlit)
def listar_semestres() -> list:
    session, site_url, raiz = _raiz_path()
    folders = _list_folders(session, site_url, raiz)
    return sorted(f["Name"] for f in folders)


def listar_eventos(semestre: str) -> list:
    session, site_url, raiz = _raiz_path()
    analisis_path = f"{raiz}/{semestre}/Análisis_todos_los_eventos"
    folders = _list_folders(session, site_url, analisis_path)
    return sorted(
        (f["Name"] for f in folders),
        key=lambda d: int(m.group(1)) if (m := re.search(r"(\d+)$", d)) else -1)


TMP_RAIZ       = _TMP_ROOT
TMP_LOC_FOLDER = _TMP_ROOT.parent / _LOC_FOLDER
TMP_DATOS      = _TMP_ROOT.parent / _DATOS_FOLDER
