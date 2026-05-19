"""
loc_namesGEN_run.py
Wrapper no-interactivo de loc_namesGEN.py.

Recibe un archivo JSON con parámetros como sys.argv[1].

Estructura del JSON:
{
    "DATOS_PF":      "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\DatosSINdigsilent.xlsx",
    "DATOS_SIN_PATH":"C:\\Datos del CNDC\\Datos_SIN_20251210.xls",
    "OUTPUT_DIR":    "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\Designacion de loc_name",
    "SIM_REF_PATH":  "C:\\Datos del CNDC\\01_INFO CNDC_RPF\\...\\datos_simulacion_*.xlsx"
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
    print("[ERROR] Uso: python loc_namesGEN_run.py <params.json>")
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
    "Programas_1_uso_modelo", "loc_namesGEN.py",
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
    f'DATOS_PF      = r"{_p["DATOS_PF"]}"')

_code = _patch(_code,
    r'DATOS_SIN_PATH\s*=\s*r?"[^"]+"',
    f'DATOS_SIN_PATH = r"{_p["DATOS_SIN_PATH"]}"')

_code = _patch(_code,
    r'OUTPUT_DIR\s*=\s*r?"[^"]+"',
    f'OUTPUT_DIR    = r"{_p["OUTPUT_DIR"]}"')

_code = _patch(_code,
    r'SIM_REF_PATH\s*=\s*\([\s\S]*?\)',
    f'SIM_REF_PATH = r"{_p["SIM_REF_PATH"]}"')

print("=" * 60, flush=True)
print(f"  RUNNER loc_namesGEN — mapeo generadores CNDC → PF", flush=True)
print(f"  DatosPF     : {os.path.basename(_p['DATOS_PF'])}", flush=True)
print(f"  Referencia  : {os.path.basename(_p['SIM_REF_PATH'])}", flush=True)
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
