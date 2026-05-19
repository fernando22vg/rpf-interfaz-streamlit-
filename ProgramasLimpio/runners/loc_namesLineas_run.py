"""
loc_namesLineas_run.py
Wrapper no-interactivo de loc_namesLineas.py.

Recibe un archivo JSON con parámetros como sys.argv[1].

Estructura del JSON:
{
    "DATOS_PF":   "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\DatosSINdigsilent.xlsx",
    "OUTPUT_DIR": "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\Designacion de loc_name"
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
    print("[ERROR] Uso: python loc_namesLineas_run.py <params.json>")
    sys.exit(1)

with open(sys.argv[1], "r", encoding="utf-8") as _f:
    _p = json.load(_f)

# ── Monkey-patch input() ──────────────────────────────────────────────────────
def _auto_input(prompt=""):
    ps = str(prompt).strip()
    print(prompt, end="", flush=True)

    if "Presiona Enter" in ps or "Presione Enter" in ps:
        print("[AUTO — script finalizado]", flush=True)
        return ""

    print("  [WARN: input() no mapeado — devolviendo '']", flush=True)
    return ""

builtins.input = _auto_input

# ── Parchear constantes y ejecutar ───────────────────────────────────────────
_script_path = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "Programas_1_uso_modelo", "loc_namesLineas.py",
)
if not os.path.isfile(_script_path):
    print(f"[ERROR] No se encontró: {_script_path}")
    sys.exit(1)

with open(_script_path, "r", encoding="utf-8") as _f:
    _code = _f.read()


def _patch(code, pattern, replacement, flags=re.DOTALL):
    return re.sub(pattern, lambda _m: replacement, code, count=1, flags=flags)


_code = _patch(_code,
    r'DATOS_PF\s*=\s*r?"[^"]+"',
    f'DATOS_PF    = r"{_p["DATOS_PF"]}"')

_code = _patch(_code,
    r'OUTPUT_DIR\s*=\s*r?"[^"]+"',
    f'OUTPUT_DIR  = r"{_p["OUTPUT_DIR"]}"')

print("=" * 60, flush=True)
print(f"  RUNNER loc_namesLineas — catálogo de líneas PF", flush=True)
print(f"  DatosPF : {os.path.basename(_p['DATOS_PF'])}", flush=True)
print(f"  Salida  : {_p['OUTPUT_DIR']}", flush=True)
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
