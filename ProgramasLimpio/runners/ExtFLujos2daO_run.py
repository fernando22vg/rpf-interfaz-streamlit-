"""
ExtFLujos2daO_run.py
Wrapper no-interactivo de ExtFLujos2daO.py.

Recibe un archivo JSON con parámetros como sys.argv[1] y ejecuta el script
original reemplazando todos los input() con respuestas automáticas.

Estructura del JSON:
{
    "semestre":           "SEM_I_2025",
    "evento":             "Evento 12",
    "RAIZ":               "C:\\Datos del CNDC\\01_INFO CNDC_RPF",
    "LOC_NAMES_GEN_PATH": "C:\\...\\loc_names_gen.xlsx"
}
"""
import sys
import os
import json
import builtins
import re

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# ── Cargar parámetros ─────────────────────────────────────────────────────────
if len(sys.argv) < 2 or not os.path.isfile(sys.argv[1]):
    print("[ERROR] Uso: python ExtFLujos2daO_run.py <params.json>")
    sys.exit(1)

with open(sys.argv[1], "r", encoding="utf-8") as _f:
    _p = json.load(_f)

_SEM = _p["semestre"]
_EV  = _p["evento"]
_raiz = _p["RAIZ"]

# ── Pre-calcular índices ──────────────────────────────────────────────────────
try:
    _sems    = sorted(d for d in os.listdir(_raiz) if os.path.isdir(os.path.join(_raiz, d)))
    _SEM_IDX = str(_sems.index(_SEM) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] Semestre '{_SEM}' no encontrado en {_raiz}: {_e}")
    sys.exit(1)

_base_ev = os.path.join(_raiz, _SEM, "Análisis_todos_los_eventos")
try:
    _evs    = sorted(d for d in os.listdir(_base_ev) if os.path.isdir(os.path.join(_base_ev, d)))
    _EV_IDX = str(_evs.index(_EV) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] Evento '{_EV}' no encontrado en {_base_ev}: {_e}")
    sys.exit(1)

# ── Monkey-patch input() ──────────────────────────────────────────────────────
_sel_count = [0]

def _auto_input(prompt=""):
    ps = str(prompt).strip()
    print(prompt, end="", flush=True)

    # Selecciones numéricas (semestre → evento, en ese orden)
    if "Selecciona numero" in ps or "Seleccionar numero" in ps:
        _sel_count[0] += 1
        val = _SEM_IDX if _sel_count[0] == 1 else _EV_IDX
        print(f"{val}  [AUTO]", flush=True)
        return val

    # Enter final
    if "Presiona Enter" in ps or "Presione Enter" in ps:
        print("[AUTO — script finalizado]", flush=True)
        return ""

    # Fallback
    print("  [WARN: input() no mapeado — devolviendo '']", flush=True)
    return ""

builtins.input = _auto_input

# ── Parchear constantes y ejecutar ───────────────────────────────────────────
_script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "ExtFLujos2daO.py")
if not os.path.isfile(_script_path):
    print(f"[ERROR] No se encontró: {_script_path}")
    sys.exit(1)

with open(_script_path, "r", encoding="utf-8") as _f:
    _code = _f.read()

def _patch(code, pattern, replacement, flags=re.DOTALL):
    return re.sub(pattern, lambda _m: replacement, code, count=1, flags=flags)

_code = _patch(_code,
    r'RAIZ\s*=\s*r?"[^"]+"',
    f'RAIZ = r"{_p["RAIZ"]}"')

_code = _patch(_code,
    r'LOC_NAMES_GEN_PATH\s*=\s*\([\s\S]*?\)',
    f'LOC_NAMES_GEN_PATH = r"{_p["LOC_NAMES_GEN_PATH"]}"')

print("=" * 60, flush=True)
print(f"  RUNNER ExtFLujos2daO: {_SEM} / {_EV}", flush=True)
print("=" * 60, flush=True)

try:
    exec(  # noqa: S102
        compile(_code, _script_path, "exec"),
        {"__name__": "__main__", "__file__": _script_path},
    )
except SystemExit as _ex:
    sys.exit(_ex.code)
except Exception as _ex:
    print(f"\n[ERROR] {type(_ex).__name__}: {_ex}", flush=True)
    import traceback
    traceback.print_exc()
    sys.exit(1)
