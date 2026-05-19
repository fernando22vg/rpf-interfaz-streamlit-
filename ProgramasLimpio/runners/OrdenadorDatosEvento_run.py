import sys
import os
import json
import builtins
import re
import time

# Forzar UTF-8 en stdout/stderr para evitar UnicodeEncodeError en consolas cp1252 (Windows)
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# ── Cargar parámetros ─────────────────────────────────────────────────────────
if len(sys.argv) < 2 or not os.path.isfile(sys.argv[1]):
    print("[ERROR] Uso: python OrdenadorDatosEvento_run.py <params.json>")
    sys.exit(1)

with open(sys.argv[1], "r", encoding="utf-8") as _f:
    _p = json.load(_f)

_SEM = _p["semestre"]
_EV  = _p["evento"]
_RAIZ_RPF = _p["RAIZ_RPF"]
_RAIZ_DATOS = _p["RAIZ_DATOS"]

# ── Pre-calcular índices para elegir() ───────────────────────────────────────
try:
    _sems    = sorted(d for d in os.listdir(_RAIZ_RPF) if os.path.isdir(os.path.join(_RAIZ_RPF, d)))
    _SEM_IDX = str(_sems.index(_SEM) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] Semestre '{_SEM}' no encontrado en {_RAIZ_RPF}: {_e}")
    sys.exit(1)

_base_ev_rpf = os.path.join(_RAIZ_RPF, _SEM, "Análisis_todos_los_eventos")
try:
    _evs = sorted(d for d in os.listdir(_base_ev_rpf) if os.path.isdir(os.path.join(_base_ev_rpf, d)))
except FileNotFoundError as _e:
    print(f"[ERROR] Carpeta de eventos no encontrada: {_base_ev_rpf}: {_e}")
    sys.exit(1)

print(f"[DEBUG] Eventos disponibles en {_base_ev_rpf}: {_evs}")
print(f"[DEBUG] Buscando evento: {repr(_EV)}")

# Coincidencia exacta primero; si falla, normalizar (strip + sin distinción de mayúsculas)
try:
    _EV_IDX = str(_evs.index(_EV) + 1)
except ValueError:
    _ev_norm = _EV.strip().lower()
    _match = next((i for i, d in enumerate(_evs) if d.strip().lower() == _ev_norm), None)
    if _match is None:
        print(f"[ERROR] Evento '{_EV}' no encontrado en {_base_ev_rpf}. Disponibles: {_evs}")
        sys.exit(1)
    _EV = _evs[_match]
    _EV_IDX = str(_match + 1)
    print(f"[INFO] Coincidencia aproximada: '{_EV}' → índice {_EV_IDX}")

# ── Monkey-patch input() ──────────────────────────────────────────────────────
_elegir_count = [0]

def _auto_input(prompt=""):
    """Intercepta cada input() del script original y devuelve la respuesta correcta."""
    ps = str(prompt).strip()
    print(prompt, end="", flush=True)

    # 1. Selecciones de elegir() — identificadas por el prompt exacto
    if "Seleccionar numero:" in ps:
        _elegir_count[0] += 1
        val = _SEM_IDX if _elegir_count[0] == 1 else _EV_IDX
        print(f"{val}  [AUTO]", flush=True)
        return val

    # 2. Enter final ("Presiona Enter para cerrar...")
    if "Presiona Enter" in ps or "Presione Enter" in ps.lower():
        print("[AUTO — script finalizado]", flush=True)
        return ""

    # Fallback: cualquier input() no previsto
    print("  [WARN: input() no mapeado — devolviendo '']", flush=True)
    return ""

builtins.input = _auto_input

# ── Leer el script original y parchear las constantes ────────────────────────
_script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "OrdenadorDatosEvento.py")

if not os.path.isfile(_script_path):
    print(f"[ERROR] No se encontró: {_script_path}")
    sys.exit(1)

with open(_script_path, "r", encoding="utf-8") as _f:
    _code = _f.read()


def _patch(code, pattern, replacement, flags=re.DOTALL):
    """re.sub con reemplazo literal (sin interpretar backslashes del replacement)."""
    return re.sub(pattern, lambda _m: replacement, code, count=1, flags=flags)


# Rutas y constantes — se reemplazan con los valores del JSON
_code = _patch(_code,
    r'RAIZ_RPF\s*=\s*r?"[^"]+"',
    f'RAIZ_RPF   = r"{_RAIZ_RPF}"')

_code = _patch(_code,
    r'RAIZ_DATOS\s*=\s*r?"[^"]+"',
    f'RAIZ_DATOS = r"{_RAIZ_DATOS}"')

# ── Ejecutar el script parcheado ──────────────────────────────────────────────
print("=" * 60, flush=True)
print(f"  RUNNER OrdenadorDatosEvento: {_SEM} / {_EV}", flush=True)
print("=" * 60, flush=True)

try:
    exec(  # noqa: S102
        compile(_code, _script_path, "exec"),
        {"__name__": "__main__", "__file__": _script_path},
    )
except SystemExit as _ex:
    sys.exit(_ex.code)
except Exception as _ex:
    print(f"\n[ERROR] Excepción en el script: {type(_ex).__name__}: {_ex}", flush=True)
    import traceback
    traceback.print_exc()
    sys.exit(1)