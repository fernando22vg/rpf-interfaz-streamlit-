"""
CondInicialesPF_run.py
Wrapper no-interactivo de CondInicialesPF.py para uso desde Streamlit u otros scripts.

Recibe un archivo JSON con parámetros como sys.argv[1] y ejecuta el script
original reemplazando todos los input() con respuestas automáticas.

Estructura del JSON de parámetros:
{
    "semestre": "SEM_I_2025",
    "evento":   "Evento 12",
    "RAIZ":     "C:\\Datos del CNDC\\01_INFO CNDC_RPF"
}
"""
import sys
import os
import json
import builtins
import re

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ── Cargar parámetros ─────────────────────────────────────────────────────────
if len(sys.argv) < 2 or not os.path.isfile(sys.argv[1]):
    print("[ERROR] Uso: python CondInicialesPF_run.py <params.json>")
    sys.exit(1)

with open(sys.argv[1], "r", encoding="utf-8") as _f:
    _p = json.load(_f)

_SEM  = _p["semestre"]
_EV   = _p["evento"]
_RAIZ = _p["RAIZ"]

# ── Pre-calcular índices para elegir() ───────────────────────────────────────
try:
    _sems    = sorted(d for d in os.listdir(_RAIZ) if os.path.isdir(os.path.join(_RAIZ, d)))
    _SEM_IDX = str(_sems.index(_SEM) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] No se pudo encontrar semestre '{_SEM}' en {_RAIZ}: {_e}")
    sys.exit(1)

_base_ev = os.path.join(_RAIZ, _SEM, "Análisis_todos_los_eventos")
try:
    _evs    = sorted(d for d in os.listdir(_base_ev) if os.path.isdir(os.path.join(_base_ev, d)))
    _EV_IDX = str(_evs.index(_EV) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] No se pudo encontrar evento '{_EV}' en {_base_ev}: {_e}")
    sys.exit(1)

# ── Monkey-patch input() ──────────────────────────────────────────────────────
_elegir_count = [0]

def _auto_input(prompt=""):
    ps = str(prompt).strip()
    print(prompt, end="", flush=True)

    if "Seleccionar numero:" in ps:
        _elegir_count[0] += 1
        val = _SEM_IDX if _elegir_count[0] == 1 else _EV_IDX
        print(f"{val}  [AUTO]", flush=True)
        return val

    # Enter final
    print("  [AUTO]", flush=True)
    return ""

builtins.input = _auto_input

# ── Ejecutar el script original ───────────────────────────────────────────────
_script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CondInicialesPF.py")

if not os.path.isfile(_script_path):
    print(f"[ERROR] No se encontró: {_script_path}")
    sys.exit(1)

with open(_script_path, "r", encoding="utf-8") as _f:
    _code = _f.read()

# Sobreescribir rutas si vienen en el JSON (opcional)
def _patch(code, pattern, replacement, flags=re.DOTALL):
    return re.sub(pattern, lambda _m: replacement, code, count=1, flags=flags)

if "LOC_GEN_PATH" in _p:
    _code = _patch(_code, r'LOC_GEN_PATH\s*=.*', f'LOC_GEN_PATH = r"{_p["LOC_GEN_PATH"]}"', flags=0)
if "LOC_CAR_PATH" in _p:
    _code = _patch(_code, r'LOC_CAR_PATH\s*=.*', f'LOC_CAR_PATH = r"{_p["LOC_CAR_PATH"]}"', flags=0)
if "LOC_XFO_PATH" in _p:
    _code = _patch(_code, r'LOC_XFO_PATH\s*=.*', f'LOC_XFO_PATH = r"{_p["LOC_XFO_PATH"]}"', flags=0)
_code = _patch(_code, r'RAIZ\s*=\s*r?"[^"]+"', f'RAIZ = r"{_RAIZ}"')

print("=" * 60, flush=True)
print(f"  RUNNER: {_SEM} / {_EV}", flush=True)
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
