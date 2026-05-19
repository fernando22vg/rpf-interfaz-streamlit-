"""
MapeoRetirosSTI_run.py
Wrapper no-interactivo de MapeoRetirosSTI_v6.py.

Recibe un archivo JSON con parámetros como sys.argv[1].

Estructura del JSON:
{
    "DATOS_PATH":        "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\DatosSINdigsilent.xlsx",
    "LOC_NAMES_XFO":     "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\Designacion de loc_name\\loc_names_xfo.xlsx",
    "DEENER_PATH":       "C:\\Datos del CNDC\\01_INFO CNDC_RPF\\...\\deener_DDMMYY.xlsx",
    "POSTOT_PATH":       "",
    "OUTPUT_DIR":        "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\Designacion de loc_name",
    "HORA_EVENTO_LABEL": "18:45"
}

POSTOT_PATH puede ser "" si no hay archivo de retiros STI disponible.
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
    print("[ERROR] Uso: python MapeoRetirosSTI_run.py <params.json>")
    sys.exit(1)

with open(sys.argv[1], "r", encoding="utf-8") as _f:
    _p = json.load(_f)

# ── Monkey-patch input() (por si acaso) ───────────────────────────────────────
def _auto_input(prompt=""):
    ps = str(prompt).strip()
    print(prompt, end="", flush=True)
    print("  [WARN: input() no mapeado — devolviendo '']", flush=True)
    return ""

builtins.input = _auto_input

# ── Parchear constantes y ejecutar ───────────────────────────────────────────
_script_path = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "Programas_1_uso_modelo", "MapeoRetirosSTI_v6.py",
)
if not os.path.isfile(_script_path):
    print(f"[ERROR] No se encontró: {_script_path}")
    sys.exit(1)

with open(_script_path, "r", encoding="utf-8") as _f:
    _code = _f.read()


def _patch(code, pattern, replacement, flags=re.DOTALL):
    return re.sub(pattern, lambda _m: replacement, code, count=1, flags=flags)


_code = _patch(_code,
    r'DATOS_PATH\s*=\s*r?"[^"]+"',
    f'DATOS_PATH       = r"{_p["DATOS_PATH"]}"')

_code = _patch(_code,
    r'LOC_NAMES_XFO\s*=\s*r?"[^"]+"',
    f'LOC_NAMES_XFO    = r"{_p["LOC_NAMES_XFO"]}"')

_code = _patch(_code,
    r'DEENER_PATH\s*=\s*r?"[^"]*"',
    f'DEENER_PATH      = r"{_p["DEENER_PATH"]}"')

_code = _patch(_code,
    r'POSTOT_PATH\s*=\s*r?"[^"]*"',
    f'POSTOT_PATH      = r"{_p.get("POSTOT_PATH", "")}"')

_code = _patch(_code,
    r'OUTPUT_DIR\s*=\s*r?"[^"]+"',
    f'OUTPUT_DIR       = r"{_p["OUTPUT_DIR"]}"')

_code = _patch(_code,
    r'HORA_EVENTO_LABEL\s*=\s*"[^"]+"',
    f'HORA_EVENTO_LABEL = "{_p.get("HORA_EVENTO_LABEL", "18:45")}"')

print("=" * 60, flush=True)
print(f"  RUNNER MapeoRetirosSTI — mapeo cargas PF → distribuidores", flush=True)
print(f"  Deener      : {os.path.basename(_p['DEENER_PATH'])}", flush=True)
print(f"  Hora evento : {_p.get('HORA_EVENTO_LABEL', '18:45')}", flush=True)
print(f"  Salida      : {_p['OUTPUT_DIR']}", flush=True)
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
