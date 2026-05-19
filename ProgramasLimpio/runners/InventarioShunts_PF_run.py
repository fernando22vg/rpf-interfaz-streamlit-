"""
InventarioShunts_PF_run.py
Wrapper no-interactivo de InventarioShunts_PF.py.

Recibe un archivo JSON con parámetros como sys.argv[1].

Estructura del JSON:
{
    "PF_BASE":     "C:\\Program Files\\DIgSILENT\\PowerFactory 2025 SP2",
    "PF_PROYECTO": "PMP_NOV25_OCT29_31102025(1)",
    "CASO_BASE":   "CNDC",
    "OUTPUT_DIR":  "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\Designacion de loc_name",
    "tap_shunts":  false
}

tap_shunts: true  → activa 'Tap Adjustment of Shunts' en el Load Flow
            false → ejecuta sin ajuste (defecto)
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
    print("[ERROR] Uso: python InventarioShunts_PF_run.py <params.json>")
    sys.exit(1)

with open(sys.argv[1], "r", encoding="utf-8") as _f:
    _p = json.load(_f)

_TAP = "1" if bool(_p.get("tap_shunts", False)) else "2"

# ── Monkey-patch input() ──────────────────────────────────────────────────────
def _auto_input(prompt=""):
    ps = str(prompt).strip()
    print(prompt, end="", flush=True)

    # Selección tap adjustment of shunts [1/2]
    if "Seleccion [1/2]" in ps:
        print(f"{_TAP}  [AUTO]", flush=True)
        return _TAP

    if "Presiona Enter" in ps or "Presione Enter" in ps:
        print("[AUTO — script finalizado]", flush=True)
        return ""

    print("  [WARN: input() no mapeado — devolviendo '']", flush=True)
    return ""

builtins.input = _auto_input

# ── Parchear constantes y ejecutar ───────────────────────────────────────────
_script_path = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "Programas_1_uso_modelo", "InventarioShunts_PF.py",
)
if not os.path.isfile(_script_path):
    print(f"[ERROR] No se encontró: {_script_path}")
    sys.exit(1)

with open(_script_path, "r", encoding="utf-8") as _f:
    _code = _f.read()


def _patch(code, pattern, replacement, flags=re.DOTALL):
    return re.sub(pattern, lambda _m: replacement, code, count=1, flags=flags)


_code = _patch(_code,
    r'PF_BASE\s*=\s*r?"[^"]+"',
    f'PF_BASE     = r"{_p["PF_BASE"]}"')

_code = _patch(_code,
    r'PF_PROYECTO\s*=\s*"[^"]+"',
    f'PF_PROYECTO = "{_p["PF_PROYECTO"]}"')

_code = _patch(_code,
    r'CASO_BASE\s*=\s*"[^"]+"',
    f'CASO_BASE   = "{_p["CASO_BASE"]}"')

_code = _patch(_code,
    r'OUTPUT_DIR\s*=\s*r?"[^"]+"',
    f'OUTPUT_DIR  = r"{_p["OUTPUT_DIR"]}"')

print("=" * 60, flush=True)
print(f"  RUNNER InventarioShunts_PF — inventario de shunts y compensadores", flush=True)
print(f"  Proyecto PF  : {_p['PF_PROYECTO']}", flush=True)
print(f"  Tap Shunts   : {'SI' if _TAP == '1' else 'NO'}", flush=True)
print(f"  Salida       : {_p['OUTPUT_DIR']}", flush=True)
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
