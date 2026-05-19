"""
DatsoGENBUSLNE_run.py
Wrapper no-interactivo de DatsoGENBUSLNE.py.

Recibe un archivo JSON con parámetros como sys.argv[1].

Estructura del JSON:
{
    "PF_DIR":      "C:\\Program Files\\DIgSILENT\\PowerFactory 2025 SP2",
    "PF_PY":       "C:\\Program Files\\DIgSILENT\\PowerFactory 2025 SP2\\Python\\3.12",
    "PF_PROYECTO": "PMP_NOV25_OCT29_31102025(1)",
    "output_path": "C:\\Datos del CNDC\\DATOS EXTRAIDOS DE DIGSILENT\\DatosSINdigsilent.xlsx"
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
    print("[ERROR] Uso: python DatsoGENBUSLNE_run.py <params.json>")
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
    "Programas_1_uso_modelo", "DatsoGENBUSLNE.py",
)
if not os.path.isfile(_script_path):
    print(f"[ERROR] No se encontró: {_script_path}")
    sys.exit(1)

with open(_script_path, "r", encoding="utf-8") as _f:
    _code = _f.read()


def _patch(code, pattern, replacement, flags=re.DOTALL):
    return re.sub(pattern, lambda _m: replacement, code, count=1, flags=flags)


_code = _patch(_code,
    r'PF_DIR\s*=\s*r?"[^"]+"',
    f'PF_DIR = r"{_p["PF_DIR"]}"')

_code = _patch(_code,
    r'PF_PY\s*=\s*r?"[^"]+"',
    f'PF_PY  = r"{_p["PF_PY"]}"')

_code = _patch(_code,
    r'ActivateProject\("[^"]+"\)',
    f'ActivateProject("{_p["PF_PROYECTO"]}")')

_code = _patch(_code,
    r'output_path\s*=\s*r?"[^"]+"',
    f'output_path = r"{_p["output_path"]}"')

print("=" * 60, flush=True)
print(f"  RUNNER DatsoGENBUSLNE — extracción completa de red", flush=True)
print(f"  Proyecto PF : {_p['PF_PROYECTO']}", flush=True)
print(f"  Salida      : {_p['output_path']}", flush=True)
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
