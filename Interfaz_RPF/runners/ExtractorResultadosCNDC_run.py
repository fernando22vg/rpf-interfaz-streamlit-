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
    print("[ERROR] Uso: python ExtractorResultadosCNDC_run.py <params.json>")
    sys.exit(1)

with open(sys.argv[1], "r", encoding="utf-8") as _f:
    _p = json.load(_f)

_SEM = _p["semestre"]
_EV  = _p["evento"]
_RAIZ = _p["RAIZ"]
_CARPETA_COBEE = _p["CARPETA_COBEE"]

# ── Pre-calcular índices para elegir_idx() ───────────────────────────────────────
try:
    _sems    = sorted(d for d in os.listdir(_RAIZ) if os.path.isdir(os.path.join(_RAIZ, d)))
    _SEM_IDX = str(_sems.index(_SEM) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] Semestre '{_SEM}' no encontrado en {_RAIZ}: {_e}")
    sys.exit(1)

_base_ev_rpf = os.path.join(_RAIZ, _SEM, "Análisis_todos_los_eventos")
# Handle potential alternative folder name "Analisis_todos_los_eventos"
if not os.path.isdir(_base_ev_rpf):
    _base_ev_rpf = os.path.join(_RAIZ, _SEM, "Analisis_todos_los_eventos")

try:
    _evs    = sorted(
        (d for d in os.listdir(_base_ev_rpf) if os.path.isdir(os.path.join(_base_ev_rpf, d))),
        key=lambda d: int(m.group(1)) if (m := re.search(r"(\d+)$", d)) else -1)
    _EV_IDX = str(_evs.index(_EV) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] Evento '{_EV}' no encontrado en {_base_ev_rpf}: {_e}")
    sys.exit(1)

# ── Monkey-patch input() ──────────────────────────────────────────────────────
_elegir_idx_count = [0]

def _auto_input(prompt=""):
    """Intercepta cada input() del script original y devuelve la respuesta correcta."""
    ps = str(prompt).strip()
    print(prompt, end="", flush=True)

    # 1. Selecciones de elegir_idx() — identificadas by "Selecciona numero:"
    if "Selecciona numero:" in ps:
        _elegir_idx_count[0] += 1
        val = _SEM_IDX if _elegir_idx_count[0] == 1 else _EV_IDX
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
_script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "ExtractorResultadosCNDC.py")

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
    r'RAIZ\s*=\s*r?"[^"]+"',
    f'RAIZ             = r"{_RAIZ}"')

_code = _patch(_code,
    r'CARPETA_COBEE\s*=\s*"[^"]+"',
    f'CARPETA_COBEE    = "{_CARPETA_COBEE}"')

# ── Ejecutar el script parcheado ──────────────────────────────────────────────
print("=" * 60, flush=True)
print(f"  RUNNER ExtractorResultadosCNDC: {_SEM} / {_EV}", flush=True)
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