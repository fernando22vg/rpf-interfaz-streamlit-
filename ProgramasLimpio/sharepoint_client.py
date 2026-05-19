"""
sharepoint_client.py  (implementación Google Drive)
Accede a la carpeta de Google Drive con los datos de eventos RPF.

Requiere en Streamlit secrets (Settings → Secrets):

  [gdrive]
  api_key   = "AIzaSy..."          # Google API key (lectura, sin OAuth)
  folder_id = "1ABCxyz..."         # ID de la carpeta raíz en Google Drive

La carpeta debe estar compartida como "Cualquiera con el enlace puede ver".
Estructura esperada en Google Drive:
  📁 raíz/
    └── 📁 Semestre_1/
          └── 📁 Análisis_todos_los_eventos/
                └── 📁 Evento_001/
                      ├── 📁 Resultados_COBEE/
                      ├── 📁 Graficas Registro 1SEG COBEE/
                      ├── 📁 E1.0/Datos Curvas/
                      └── 📁 E1.1/Datos Curvas/
"""

import shutil
import tempfile
from pathlib import Path
from typing import Optional

import requests
import streamlit as st

_TMP_ROOT = Path(tempfile.gettempdir()) / "rpf_gdrive"
_DRIVE_API = "https://www.googleapis.com/drive/v3"
_FOLDER_MIME = "application/vnd.google-apps.folder"


# ── Credenciales ──────────────────────────────────────────────────────────────

def _creds() -> tuple[str, str]:
    """Devuelve (api_key, root_folder_id) desde Streamlit secrets."""
    try:
        cfg = st.secrets["gdrive"]
        return cfg["api_key"], cfg["folder_id"]
    except (KeyError, AttributeError):
        raise RuntimeError(
            "Configura los secrets de Streamlit Cloud:\n"
            "  [gdrive]\n"
            "  api_key   = 'AIzaSy...'\n"
            "  folder_id = '1ABCxyz...'"
        )


# ── Google Drive API helpers ──────────────────────────────────────────────────

def _list_children(folder_id: str, api_key: str) -> list:
    """Lista todos los archivos/carpetas hijos de una carpeta."""
    items, page_token = [], None
    while True:
        params = {
            "q": f"'{folder_id}' in parents and trashed=false",
            "key": api_key,
            "fields": "nextPageToken,files(id,name,mimeType)",
            "pageSize": 200,
        }
        if page_token:
            params["pageToken"] = page_token
        r = requests.get(f"{_DRIVE_API}/files", params=params, timeout=20)
        if r.status_code == 403:
            raise PermissionError(
                "Google Drive API key inválida o carpeta no es pública. "
                "Verifica que la carpeta esté compartida como 'Cualquiera con el enlace'."
            )
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("files", []))
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return items


def _find_child(folder_id: str, name: str, api_key: str) -> Optional[dict]:
    """Busca un hijo por nombre exacto."""
    for item in _list_children(folder_id, api_key):
        if item["name"] == name:
            return item
    return None


# ── API pública (misma interfaz que sharepoint_client original) ───────────────

@st.cache_data(ttl=300, show_spinner=False)
def listar_semestres() -> list:
    """Lista semestres (carpetas de primer nivel en la raíz de GDrive)."""
    api_key, root_id = _creds()
    children = _list_children(root_id, api_key)
    return sorted(c["name"] for c in children if c["mimeType"] == _FOLDER_MIME)


@st.cache_data(ttl=300, show_spinner=False)
def listar_eventos(semestre: str) -> list:
    """Lista eventos dentro de semestre/Análisis_todos_los_eventos/."""
    api_key, root_id = _creds()

    sem_item = _find_child(root_id, semestre, api_key)
    if not sem_item:
        return []

    analisis_item = _find_child(sem_item["id"], "Análisis_todos_los_eventos", api_key)
    if not analisis_item:
        return []

    children = _list_children(analisis_item["id"], api_key)
    return sorted(c["name"] for c in children if c["mimeType"] == _FOLDER_MIME)


def descargar_evento(semestre: str, evento: str, progress_cb=None) -> Path:
    """
    Descarga la carpeta del evento desde Google Drive a /tmp/.
    Devuelve la ruta local equivalente a ev_path.
    Si ya existe en caché local, no re-descarga.
    """
    local_path = _TMP_ROOT / semestre / evento
    if local_path.exists():
        return local_path

    api_key, root_id = _creds()

    sem_item = _find_child(root_id, semestre, api_key)
    if not sem_item:
        raise FileNotFoundError(f"Semestre no encontrado en Google Drive: {semestre}")

    analisis_item = _find_child(sem_item["id"], "Análisis_todos_los_eventos", api_key)
    if not analisis_item:
        raise FileNotFoundError(f"Carpeta 'Análisis_todos_los_eventos' no encontrada en {semestre}")

    event_item = _find_child(analisis_item["id"], evento, api_key)
    if not event_item:
        raise FileNotFoundError(f"Evento no encontrado en Google Drive: {evento}")

    _download_tree(event_item["id"], local_path, api_key, progress_cb)
    return local_path


def limpiar_cache_evento(semestre: str, evento: str):
    """Elimina el caché local de un evento para forzar re-descarga."""
    local_path = _TMP_ROOT / semestre / evento
    if local_path.exists():
        shutil.rmtree(local_path)


# ── Descarga recursiva ────────────────────────────────────────────────────────

def _download_tree(folder_id: str, dest: Path, api_key: str, progress_cb=None):
    dest.mkdir(parents=True, exist_ok=True)
    for item in _list_children(folder_id, api_key):
        name = item["name"]
        if item["mimeType"] == _FOLDER_MIME:
            _download_tree(item["id"], dest / name, api_key, progress_cb)
        else:
            local_file = dest / name
            if not local_file.exists():
                _download_file(item["id"], local_file)
            if progress_cb:
                progress_cb(name)


def _download_file(file_id: str, dest: Path):
    """Descarga un archivo público de Google Drive."""
    url = f"https://drive.google.com/uc?id={file_id}&confirm=t&export=download"
    with requests.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(chunk_size=65536):
                f.write(chunk)
