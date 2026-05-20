"""
sync_watcher.py
───────────────
Sincronización automática local → SharePoint para modo local (IS_CLOUD=False).

Detecta cambios en RAIZ_RPF (create/modify) mediante watchdog y los sube
a SharePoint en un hilo de fondo con debounce de 2 s.

Uso típico (desde interfaz_analisis_RPF.py):
    from sync_watcher import get_watcher
    watcher = get_watcher(RAIZ_RPF)
    watcher.start(raiz_local=RAIZ_RPF)

Requisito opcional:
    pip install watchdog
"""

from __future__ import annotations

import logging
import threading
import time
from pathlib import Path

log = logging.getLogger("sync_watcher")

# ── Extensiones que se sincronizan ────────────────────────────────────────────
SYNC_EXTS    = {".xlsx", ".xls", ".json", ".csv", ".txt", ".pdf"}
IGNORE_PREFX = {"~$", ".tmp", ".lock", "._"}   # prefijos de archivo temporal

# ── Importación opcional de watchdog ─────────────────────────────────────────
try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler as _FSHandler
    WATCHDOG_OK = True
except ImportError:
    WATCHDOG_OK = False
    _FSHandler  = object          # placeholder para herencia


# ─────────────────────────────────────────────────────────────────────────────
# Handler de eventos del sistema de archivos
# ─────────────────────────────────────────────────────────────────────────────

class _SPSyncHandler(_FSHandler):
    """Recibe eventos de watchdog y los encola para upload con debounce."""

    def __init__(self, raiz_local: str, debounce_s: float = 2.0):
        if WATCHDOG_OK:
            super().__init__()
        self.raiz_local  = str(raiz_local)
        self.debounce_s  = debounce_s
        self._pending: dict[str, float] = {}   # path → tiempo_due
        self._lock       = threading.Lock()
        self._stats      = {"uploaded": 0, "errors": 0,
                            "last_file": None, "last_ts": None}

        # Hilo de flush en background
        self._flush_thread = threading.Thread(
            target=self._flush_loop, daemon=True, name="sp-sync-flush"
        )
        self._flush_thread.start()

    # ── Callbacks watchdog ────────────────────────────────────────────────────

    def on_created(self, event):
        if not event.is_directory:
            self._schedule(event.src_path)

    def on_modified(self, event):
        if not event.is_directory:
            self._schedule(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            self._schedule(event.dest_path)

    # ── Lógica interna ────────────────────────────────────────────────────────

    def _should_sync(self, path: str) -> bool:
        p = Path(path)
        if not p.is_file():
            return False
        if p.suffix.lower() not in SYNC_EXTS:
            return False
        if any(p.name.startswith(pfx) for pfx in IGNORE_PREFX):
            return False
        return True

    def _schedule(self, path: str):
        if self._should_sync(path):
            with self._lock:
                self._pending[path] = time.time() + self.debounce_s

    def _flush_loop(self):
        """Cada 0.5 s sube los archivos cuyo debounce ha expirado."""
        import sharepoint_client as _sp  # importación tardía para evitar ciclos
        while True:
            time.sleep(0.5)
            now       = time.time()
            to_upload = []
            with self._lock:
                for path, due_at in list(self._pending.items()):
                    if now >= due_at:
                        to_upload.append(path)
                        del self._pending[path]

            for path in to_upload:
                try:
                    sp_folder = _sp.local_path_to_sp_folder(path, self.raiz_local)
                    _sp.ensure_sp_folder(sp_folder)
                    _sp.upload_file(path, sp_folder)
                    self._stats["uploaded"]  += 1
                    self._stats["last_file"]  = Path(path).name
                    self._stats["last_ts"]    = time.strftime("%H:%M:%S")
                    log.info("Synced ✔ %s → %s", path, sp_folder)
                except Exception as exc:
                    self._stats["errors"] += 1
                    log.warning("Sync error: %s — %s", path, exc)

    @property
    def stats(self) -> dict:
        with self._lock:
            pending = len(self._pending)
        return {**self._stats, "pending": pending}


# ─────────────────────────────────────────────────────────────────────────────
# Clase pública: SharePointWatcher
# ─────────────────────────────────────────────────────────────────────────────

class SharePointWatcher:
    """
    Observa RAIZ_RPF en busca de cambios y los sube a SharePoint.

    Ciclo de vida:
        watcher = SharePointWatcher()
        ok = watcher.start(raiz_local)   # True si watchdog disponible
        ...
        watcher.stop()
    """

    def __init__(self):
        self._handler:  _SPSyncHandler | None = None
        self._observer: "Observer | None"      = None
        self._running   = False
        self._raiz      = ""

    def start(self, raiz_local: str) -> bool:
        """Inicia el watcher. Devuelve True si se pudo iniciar."""
        if self._running:
            return True
        if not WATCHDOG_OK:
            log.warning(
                "sync_watcher: watchdog no instalado. "
                "Instale con:  pip install watchdog"
            )
            return False
        if not Path(raiz_local).is_dir():
            log.warning("sync_watcher: directorio no existe: %s", raiz_local)
            return False

        self._raiz    = raiz_local
        self._handler = _SPSyncHandler(raiz_local)
        self._observer = Observer()
        self._observer.schedule(self._handler, raiz_local, recursive=True)
        self._observer.start()
        self._running = True
        log.info("sync_watcher iniciado: %s", raiz_local)
        return True

    def stop(self):
        if self._observer and self._running:
            self._observer.stop()
            self._observer.join(timeout=5)
            self._running = False

    def sync_file_now(self, local_path: str):
        """Fuerza la subida inmediata de un archivo (sin esperar debounce)."""
        if self._handler:
            self._handler._schedule(local_path)
            # Acortar debounce a 0 para subida inmediata
            with self._handler._lock:
                self._handler._pending[local_path] = time.time()

    @property
    def is_running(self) -> bool:
        return self._running

    @property
    def watchdog_available(self) -> bool:
        return WATCHDOG_OK

    @property
    def stats(self) -> dict:
        if self._handler:
            return self._handler.stats
        return {"uploaded": 0, "errors": 0, "last_file": None,
                "last_ts": None, "pending": 0}


# ─────────────────────────────────────────────────────────────────────────────
# Singleton por proceso (evita múltiples observers en reruns de Streamlit)
# ─────────────────────────────────────────────────────────────────────────────

_watcher: SharePointWatcher | None = None
_watcher_lock = threading.Lock()


def get_watcher() -> SharePointWatcher:
    """Devuelve la instancia singleton del watcher."""
    global _watcher
    with _watcher_lock:
        if _watcher is None:
            _watcher = SharePointWatcher()
    return _watcher
