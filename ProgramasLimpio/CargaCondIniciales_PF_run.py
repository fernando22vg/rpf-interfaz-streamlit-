"""
CargaCondIniciales_PF_run.py
Wrapper no-interactivo de CargaCondIniciales_PF.py.

Recibe un archivo JSON con parámetros como sys.argv[1] y ejecuta el script
original reemplazando todos los input() con respuestas automáticas.

Estructura del JSON de parámetros:
{
    "semestre":       "SEM_I_2025",
    "evento":         "Evento 12",
    "RAIZ":           "C:\\Datos del CNDC\\01_INFO CNDC_RPF",
    "PF_BASE":        "C:\\Program Files\\DIgSILENT\\PowerFactory 2025 SP2",
    "LOC_XFO_PATH":   "C:\\...\\loc_names_xfo.xlsx",
    "PF_PROYECTO":    "PMP_NOV25_OCT29_31102025(1)",
    "CASO_BASE":      "CNDC",
    "modo_disparo":   "1",
    "pgini_manual":   {"sym_GCH01": 80.5, "sym_GCH02": 75.0},
    "ajustar_post_lf": false
}
"""
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
    print("[ERROR] Uso: python CargaCondIniciales_PF_run.py <params.json>")
    sys.exit(1)

with open(sys.argv[1], "r", encoding="utf-8") as _f:
    _p = json.load(_f)

_SEM          = _p["semestre"]
_EV           = _p["evento"]
_MODO         = str(_p.get("modo_disparo", "1"))
_MANUAL       = _p.get("pgini_manual", {})
_AJUSTAR         = bool(_p.get("ajustar_post_lf", False))
_KEEP_PF_OPEN    = bool(_p.get("keep_pf_open", True))
_GUARDAR_ESCEN   = bool(_p.get("guardar_escenario", True))

# ── Pre-calcular índices para elegir() ───────────────────────────────────────
_raiz = _p["RAIZ"]
try:
    _sems    = sorted(d for d in os.listdir(_raiz) if os.path.isdir(os.path.join(_raiz, d)))
    _SEM_IDX = str(_sems.index(_SEM) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] No se pudo encontrar semestre '{_SEM}' en {_raiz}: {_e}")
    sys.exit(1)

_base_ev = os.path.join(_raiz, _SEM, "Análisis_todos_los_eventos")
_ev_path = os.path.join(_base_ev, _EV)
_FLAG_WAITING  = os.path.join(_ev_path, "_pf_waiting.flag")
_FLAG_CONTINUE = os.path.join(_ev_path, "_pf_continue.flag")
try:
    _evs    = sorted(d for d in os.listdir(_base_ev) if os.path.isdir(os.path.join(_base_ev, d)))
    _EV_IDX = str(_evs.index(_EV) + 1)
except (FileNotFoundError, ValueError) as _e:
    print(f"[ERROR] No se pudo encontrar evento '{_EV}' en {_base_ev}: {_e}")
    sys.exit(1)

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

    # 2. Menú de asignación del disparo [1/2/3/4]
    if "1/2/3/4" in ps or "Seleccionar [1" in ps:
        print(f"{_MODO}  [AUTO]", flush=True)
        return _MODO

    # 3. Ingreso manual de pgini por unidad  (prompt: "  loc_name  (Pmax=...) -> Nuevo valor...")
    if "Nuevo valor [MW]" in ps:
        # Extrae el loc_name al inicio del prompt (antes del paréntesis)
        m = re.match(r"\s*(\S+)", ps)
        loc = m.group(1) if m else ""
        val_num = _MANUAL.get(loc)
        val = str(round(float(val_num), 4)) if val_num is not None else ""
        print(f"{val}  [AUTO]", flush=True)
        return val

    # 4. Confirmación de continuar a pesar de diferencia >= 5 MW
    if "Continuar de todas formas" in ps or "¿Continuar" in ps:
        print("s  [AUTO]", flush=True)
        return "s"

    # 5. Ajuste post-LF
    if "Realizar ajuste" in ps or "¿Realizar" in ps:
        val = "s" if _AJUSTAR else "n"
        print(f"{val}  [AUTO]", flush=True)
        return val

    # 6. Enter final ("Presione Enter para finalizar...")
    if "Presione Enter" in ps or "presione enter" in ps.lower():
        if _KEEP_PF_OPEN:
            # Limpiar señal anterior y avisar a la interfaz que PF está abierto
            if os.path.exists(_FLAG_CONTINUE):
                os.remove(_FLAG_CONTINUE)
            with open(_FLAG_WAITING, "w") as _fw:
                _fw.write("waiting")
            print("[PF-OPEN] PowerFactory abierto. Esperando señal de cierre desde la interfaz...", flush=True)
            # Esperar hasta que la interfaz escriba el archivo de señal (máx. 30 min)
            _timeout = 1800
            _elapsed = 0
            while not os.path.exists(_FLAG_CONTINUE) and _elapsed < _timeout:
                time.sleep(1)
                _elapsed += 1
            # Limpiar flags
            for _f in (_FLAG_WAITING, _FLAG_CONTINUE):
                if os.path.exists(_f):
                    try:
                        os.remove(_f)
                    except OSError:
                        pass
            print("[PF-OPEN] Señal recibida — cerrando PowerFactory.", flush=True)
        else:
            print("[AUTO — script finalizado]", flush=True)
        return ""

    # Fallback: cualquier input() no previsto
    print("  [WARN: input() no mapeado — devolviendo '']", flush=True)
    return ""


builtins.input = _auto_input

# ── Leer el script original y parchear las constantes ────────────────────────
_script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "CargaCondIniciales_PF.py")

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
    f'RAIZ        = r"{_p["RAIZ"]}"')

_code = _patch(_code,
    r'PF_BASE\s*=\s*r?"[^"]+"',
    f'PF_BASE     = r"{_p["PF_BASE"]}"')

_code = _patch(_code,
    r'LOC_XFO_PATH\s*=\s*\([\s\S]*?\)',
    f'LOC_XFO_PATH = r"{_p["LOC_XFO_PATH"]}"')

_code = _patch(_code,
    r'PF_PROYECTO\s*=\s*"[^"]+"',
    f'PF_PROYECTO    = "{_p["PF_PROYECTO"]}"')

_code = _patch(_code,
    r'CASO_BASE\s*=\s*"[^"]+"',
    f'CASO_BASE      = "{_p["CASO_BASE"]}"')

_code = _patch(_code,
    r'AJUSTAR_POST_LF\s*=\s*\w+',
    f'AJUSTAR_POST_LF = {_AJUSTAR}')

_code = _patch(_code,
    r'GUARDAR_ESCENARIO\s*=\s*\w+',
    f'GUARDAR_ESCENARIO = {_GUARDAR_ESCEN}')

# ── Ejecutar el script parcheado ──────────────────────────────────────────────
print("=" * 60, flush=True)
print(f"  RUNNER: {_SEM} / {_EV}", flush=True)
print(f"  Modo disparo: {_MODO}  |  Ajuste post-LF: {_AJUSTAR}  |  Guardar escenario: {_GUARDAR_ESCEN}", flush=True)
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
