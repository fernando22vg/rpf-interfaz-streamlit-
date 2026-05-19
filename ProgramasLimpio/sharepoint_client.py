"""
sharepoint_client.py
Acceso a la carpeta de SharePoint que contiene los datos de eventos RPF.

Autenticación: Azure App Registration (client credentials flow).
Requiere en Streamlit secrets (Settings → Secrets):

  [sharepoint]
  tenant_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
  client_id     = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
  client_secret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
"""

import base64
import shutil
import tempfile
import threading
import time
from pathlib import Path
from typing import Optional

import requests
import streamlit as st

# ── Carpeta raíz compartida en SharePoint ─────────────────────────────────────
SHARE_URL = (
    "https://cobee1-my.sharepoint.com/:f:/g/personal/angel_mariscal_cobee_com"
    "/IgDQ0-3WNNN1SYksWDQKnGTeAdQNzcw0KrsBYeBuI7_NAf0?e=SBurQb"
)

_TMP_ROOT = Path(tempfile.gettempdir()) / "rpf_sharepoint"
_token_cache: dict = {}
_token_lock = threading.Lock()

GRAPH = "https://graph.microsoft.com/v1.0"


# ── Autenticación ─────────────────────────────────────────────────────────────

def _get_access_token() -> str:
    with _token_lock:
        if _token_cache.get("expires_at", 0) > time.time() + 60:
            return _token_cache["access_token"]

        try:
            import msal
        except ImportError:
            raise RuntimeError(
                "Falta el paquete 'msal'. Agrégalo a requirements.txt."
            )

        try:
            sp = st.secrets["sharepoint"]
            tenant_id = sp["tenant_id"]
            client_id = sp["client_id"]
            client_secret = sp["client_secret"]
        except (KeyError, AttributeError):
            raise RuntimeError(
                "Configura los secrets de Streamlit Cloud:\n"
                "  [sharepoint]\n"
                "  tenant_id     = '...'\n"
                "  client_id     = '...'\n"
                "  client_secret = '...'"
            )

        app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if "access_token" not in result:
            raise RuntimeError(
                f"Error de autenticación SharePoint: {result.get('error_description', result)}"
            )

        _token_cache["access_token"] = result["access_token"]
        _token_cache["expires_at"] = time.time() + result.get("expires_in", 3600)
        return _token_cache["access_token"]


def _headers() -> dict:
    return {"Authorization": f"Bearer {_get_access_token()}", "Accept": "application/json"}


# ── Graph API helpers ─────────────────────────────────────────────────────────

def _encode_share_url(url: str) -> str:
    b64 = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    return f"u!{b64}"


@st.cache_data(ttl=300, show_spinner=False)
def _get_root_item() -> dict:
    encoded = _encode_share_url(SHARE_URL)
    r = requests.get(f"{GRAPH}/shares/{encoded}/driveItem", headers=_headers(), timeout=20)
    r.raise_for_status()
    return r.json()


def _children(drive_id: str, item_id: str) -> list:
    items, url = [], f"{GRAPH}/drives/{drive_id}/items/{item_id}/children"
    while url:
        r = requests.get(url, headers=_headers(), timeout=20)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _item_by_path(drive_id: str, root_id: str, rel_path: str) -> Optional[dict]:
    r = requests.get(
        f"{GRAPH}/drives/{drive_id}/items/{root_id}:/{rel_path}",
        headers=_headers(), timeout=20
    )
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


# ── API pública ───────────────────────────────────────────────────────────────

@st.cache_data(ttl=300, show_spinner=False)
def listar_semestres() -> list:
    root = _get_root_item()
    drive_id = root["parentReference"]["driveId"]
    kids = _children(drive_id, root["id"])
    return sorted(c["name"] for c in kids if "folder" in c)


@st.cache_data(ttl=300, show_spinner=False)
def listar_eventos(semestre: str) -> list:
    root = _get_root_item()
    drive_id = root["parentReference"]["driveId"]
    analisis = _item_by_path(
        drive_id, root["id"], f"{semestre}/Análisis_todos_los_eventos"
    )
    if not analisis:
        return []
    kids = _children(drive_id, analisis["id"])
    return sorted(c["name"] for c in kids if "folder" in c)


def descargar_evento(semestre: str, evento: str, progress_cb=None) -> Path:
    """
    Descarga la carpeta del evento desde SharePoint a /tmp/.
    Devuelve la ruta local equivalente a ev_path.
    Si ya existe en caché, retorna sin re-descargar.
    """
    local_path = _TMP_ROOT / semestre / evento
    if local_path.exists():
        return local_path

    root = _get_root_item()
    drive_id = root["parentReference"]["driveId"]
    rel = f"{semestre}/Análisis_todos_los_eventos/{evento}"
    event_item = _item_by_path(drive_id, root["id"], rel)
    if not event_item:
        raise FileNotFoundError(f"Evento no encontrado en SharePoint: {rel}")

    _download_tree(drive_id, event_item["id"], local_path, progress_cb)
    return local_path


def limpiar_cache_evento(semestre: str, evento: str):
    local_path = _TMP_ROOT / semestre / evento
    if local_path.exists():
        shutil.rmtree(local_path)


# ── Descarga recursiva ────────────────────────────────────────────────────────

def _download_tree(drive_id: str, item_id: str, dest: Path, progress_cb=None):
    dest.mkdir(parents=True, exist_ok=True)
    for child in _children(drive_id, item_id):
        name = child["name"]
        if "folder" in child:
            _download_tree(drive_id, child["id"], dest / name, progress_cb)
        elif "file" in child:
            local_file = dest / name
            if not local_file.exists():
                _download_file(child["@microsoft.graph.downloadUrl"], local_file)
            if progress_cb:
                progress_cb(name)


def _download_file(url: str, dest: Path):
    r = requests.get(url, stream=True, timeout=120)
    r.raise_for_status()
    with open(dest, "wb") as f:
        for chunk in r.iter_content(chunk_size=65536):
            f.write(chunk)
