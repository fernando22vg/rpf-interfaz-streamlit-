"""
batch_carga_cond_pf.py — Carga condiciones iniciales en PowerFactory para todos los eventos

Uso:
  python batch_carga_cond_pf.py                           # Tab 3 (Proyecto 1, sufijo .1)
  python batch_carga_cond_pf.py --tab 1                   # Tab 3 — PF_PROYECTO  → datos_cargados_Ev*.1*.xlsx
  python batch_carga_cond_pf.py --tab 2                   # Tab 4 — PF_PROYECTO_2 → datos_cargados_Ev*.2*.xlsx
  python batch_carga_cond_pf.py --tab ambos               # Ambos proyectos secuencialmente
  python batch_carga_cond_pf.py --sem "2024 sem2"         # Solo un semestre
  python batch_carga_cond_pf.py --ev "Evento 1"           # Solo un evento específico
  python batch_carga_cond_pf.py --skip-done               # Omite eventos ya procesados (revisa el sufijo del tab)
  python batch_carga_cond_pf.py --dry-run                 # Ver lista sin ejecutar
  python batch_carga_cond_pf.py --no-ajuste               # Desactivar ajuste post-LF
  python batch_carga_cond_pf.py --config ruta/config.json

Sufijos de archivos de salida (alineados con la interfaz Bloque 1):
  Tab 3 (Proyecto 1): datos_cargados_Ev{N}.1[_ajustado].xlsx
  Tab 4 (Proyecto 2): datos_cargados_Ev{N}.2[_ajustado].xlsx

Sin límite de tiempo por evento (PowerFactory puede tardar lo que necesite).

Para excluir eventos específicos, usar "excluir_eventos" en batch_config_pf.json:
  "excluir_eventos": [{"semestre": "2024 sem2", "evento": "Evento 1"}]
"""

import os
import sys
import json
import subprocess
import argparse
import tempfile
from pathlib import Path
from datetime import datetime

# Forzar UTF-8 en stdout/stderr para evitar UnicodeEncodeError en consolas cp1252 (Windows)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ─── Rutas por defecto ────────────────────────────────────────────────────────
_HERE = Path(__file__).parent

_DEFAULTS = {
    "RAIZ":            r"C:\Datos del CNDC\01_INFO CNDC_RPF",
    "PF_BASE":         r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2",
    "LOC_GEN_PATH":    r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Designacion de loc_name\loc_names_gen.xlsx",
    "LOC_CAR_PATH":    r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Designacion de loc_name\loc_name_cargas.xlsx",
    "LOC_XFO_PATH":    r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Designacion de loc_name\loc_names_xfo.xlsx",
    "PF_PROYECTO":     "PMP_NOV25_OCT29_31102025(1)",   # Tab 3
    "PF_PROYECTO_2":   "PMP_NOV25_OCT29_31102025(2)",   # Tab 4
    "CASO_BASE":       "CNDC",
    "modo_disparo":    "1",
    "guardar_escenario": True,
    "excluir_slack":   [],
    "xfo_pf":          1.0,
    "excluir_eventos": [],   # [{"semestre": "2024 sem2", "evento": "Evento 1"}, ...]
}

_CONFIG_PATH = _HERE / "batch_config_pf.json"
_RUNNER      = _HERE / "runners" / "CargaCondIniciales_PF_run.py"
_LOG_DIR     = _HERE / "logs_batch_pf"

# Mapeo tab → (nombre_display, clave_en_cfg, sufijo_ev)
_TAB_INFO = {
    "1": ("Tab 3 — Proyecto (1)", "PF_PROYECTO",   ".1"),
    "2": ("Tab 4 — Proyecto (2)", "PF_PROYECTO_2", ".2"),
}


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _load_config(config_path: Path) -> dict:
    """Lee la config base. Si no existe la crea con los defaults y avisa."""
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        return {**_DEFAULTS, **cfg}
    config_path.write_text(
        json.dumps(_DEFAULTS, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"[INFO] Archivo de configuración creado: {config_path}")
    print("[INFO] Edítalo con las rutas correctas antes de continuar.")
    return dict(_DEFAULTS)


def _discover_eventos(raiz: str, sem_filter: str = None, ev_filter: str = None) -> list:
    """Retorna lista de dicts {semestre, evento, ev_path} ordenados."""
    eventos = []
    base = Path(raiz)
    if not base.exists():
        print(f"[ERROR] RAIZ no existe: {raiz}")
        sys.exit(1)
    for sem_dir in sorted(base.iterdir()):
        if not sem_dir.is_dir():
            continue
        if sem_filter and sem_dir.name != sem_filter:
            continue
        analisis_dir = sem_dir / "Análisis_todos_los_eventos"
        if not analisis_dir.exists():
            analisis_dir = sem_dir / "Analisis_todos_los_eventos"
        if not analisis_dir.exists():
            continue
        for ev_dir in sorted(analisis_dir.iterdir()):
            if not (ev_dir.is_dir() and ev_dir.name.startswith("Evento")):
                continue
            if ev_filter and ev_dir.name != ev_filter:
                continue
            eventos.append({
                "semestre": sem_dir.name,
                "evento":   ev_dir.name,
                "ev_path":  str(ev_dir),
            })
    return eventos


def _ya_procesado(ev_path: str, ev_suffix: str) -> bool:
    """True si ya existe datos_cargados_*{ev_suffix}*.xlsx en la carpeta del evento.

    ev_suffix debe ser '.1' o '.2' para alinearse con Tab 3 / Tab 4 de la interfaz.
    """
    import glob as _glob
    patron = os.path.join(ev_path, f"datos_cargados_*{ev_suffix}*.xlsx")
    return len(_glob.glob(patron)) > 0


def _run_evento(cfg: dict, sem: str, evento: str, log_path: Path, ajustar: bool,
                pf_proyecto: str, ev_suffix: str) -> tuple:
    """Ejecuta el runner para un evento. Retorna (rc, duracion_seg).

    pf_proyecto: nombre del proyecto PowerFactory a usar.
    ev_suffix:   sufijo de archivo de salida ('.1' para Tab 3, '.2' para Tab 4).
    """
    params = {
        "semestre":          sem,
        "evento":            evento,
        "RAIZ":              cfg["RAIZ"],
        "PF_BASE":           cfg["PF_BASE"],
        "LOC_GEN_PATH":      cfg["LOC_GEN_PATH"],
        "LOC_CAR_PATH":      cfg.get("LOC_CAR_PATH", ""),
        "LOC_XFO_PATH":      cfg["LOC_XFO_PATH"],
        "PF_PROYECTO":       pf_proyecto,
        "CASO_BASE":         cfg["CASO_BASE"],
        "modo_disparo":      str(cfg.get("modo_disparo", "1")),
        "pgini_manual":      cfg.get("pgini_manual", {}),
        "ajustar_post_lf":   ajustar,
        "guardar_escenario": cfg.get("guardar_escenario", True),
        "keep_pf_open":      False,   # batch: cerrar PF al terminar
        "excluir_slack":     cfg.get("excluir_slack", []),
        "xfo_pf":            cfg.get("xfo_pf", 1.0),
        "ev_suffix":         ev_suffix,
    }

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".json", encoding="utf-8", delete=False
    ) as tf:
        json.dump(params, tf, ensure_ascii=False, indent=2)
        params_path = tf.name

    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    log_path.parent.mkdir(parents=True, exist_ok=True)
    t_ini = datetime.now()

    # Matar procesos PF residuales antes de iniciar (libera licencia)
    subprocess.run(["taskkill", "/F", "/IM", "PowerFactory.exe", "/T"],
                   capture_output=True, check=False)
    import time as _time
    _time.sleep(5)

    try:
        result = subprocess.run(
            [sys.executable, str(_RUNNER), params_path],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            timeout=None,   # sin límite: PowerFactory tarda lo que necesite
        )
        rc = result.returncode
        output = result.stdout + ("\n[STDERR]\n" + result.stderr if result.stderr.strip() else "")
    except Exception as exc:
        rc = -1
        output = f"[ERROR] {exc}"
    finally:
        subprocess.run(["taskkill", "/F", "/IM", "PowerFactory.exe", "/T"],
                       capture_output=True, check=False)
        try:
            os.remove(params_path)
        except OSError:
            pass

    dur = (datetime.now() - t_ini).total_seconds()

    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write(f"=== {sem} / {evento} ===\n")
        lf.write(f"Proyecto PF:      {pf_proyecto}\n")
        lf.write(f"Sufijo archivo:   '{ev_suffix}'\n")
        lf.write(f"Inicio:           {t_ini.strftime('%Y-%m-%d %H:%M:%S')}\n")
        lf.write(f"Duracion:         {dur:.0f}s\n")
        lf.write(f"Codigo de salida: {rc}\n")
        lf.write(f"Ajuste post-LF:   {ajustar}\n")
        lf.write("\n" + "─" * 60 + "\n\n")
        lf.write(output)

    # PowerFactory a veces crashea al cerrar (0xC0000005 / ACCESS_VIOLATION).
    # Si el archivo de resultados existe, la carga fue exitosa.
    _RC_PF_CRASH = {3221225477, -1073741819}
    if rc in _RC_PF_CRASH:
        import glob as _glob
        _ev_path_real = os.path.join(cfg["RAIZ"], sem, "Análisis_todos_los_eventos", evento)
        if not os.path.isdir(_ev_path_real):
            _ev_path_real = os.path.join(cfg["RAIZ"], sem, "Analisis_todos_los_eventos", evento)
        if _glob.glob(os.path.join(_ev_path_real, f"datos_cargados_*{ev_suffix}*.xlsx")):
            with open(log_path, "a", encoding="utf-8") as lf:
                lf.write("\n[INFO] rc=0xC0000005: crash de PF al cerrar, pero datos_cargados existe → carga OK\n")
            rc = 0

    return rc, dur


def _procesar_tab(cfg, eventos, tab_id, ajustar, skip_done):
    """Ejecuta el batch para un tab (proyecto). Retorna lista de resultados."""
    import glob as _glob

    label, cfg_key, ev_suffix = _TAB_INFO[tab_id]
    pf_proyecto = cfg[cfg_key]
    log_dir     = _LOG_DIR / f"tab{tab_id}"

    # Filtrar ya procesados si --skip-done
    if skip_done:
        skip_list = [ev for ev in eventos if _ya_procesado(ev["ev_path"], ev_suffix)]
        eventos   = [ev for ev in eventos if not _ya_procesado(ev["ev_path"], ev_suffix)]
        if skip_list:
            print(f"  Omitidos ya procesados: {len(skip_list)}")
            for ev in skip_list:
                print(f"    ~  {ev['semestre']} / {ev['evento']}")

    print(f"  A procesar: {len(eventos)}\n")

    results = []
    for i, ev in enumerate(eventos, 1):
        sem    = ev["semestre"]
        evento = ev["evento"]

        sem_slug = sem.replace(" ", "_")
        ev_slug  = evento.replace(" ", "_")
        log_path = log_dir / sem_slug / f"{ev_slug}.log"

        print(f"[{i:02d}/{len(eventos)}] {sem} / {evento} ...", end="", flush=True)
        rc, dur = _run_evento(cfg, sem, evento, log_path, ajustar, pf_proyecto, ev_suffix)

        if rc == 0:
            estado = "OK"
        elif rc in (3221225477, -1073741819):
            estado = "CRASH-PF (0xC0000005, sin datos_cargados)"
        else:
            estado = f"ERROR (rc={rc})"
        print(f"  {estado}  ({dur:.0f}s)  log-> {log_path.name}", flush=True)
        results.append({
            "semestre": sem,  "evento": evento,
            "tab": label,     "proyecto": pf_proyecto,
            "rc": rc,         "dur": dur,
        })

    return results


# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Carga condiciones iniciales PF para todos los eventos.\n\n"
            "Seleccionar tab con --tab:\n"
            "  1     = Tab 3, PF_PROYECTO,   sufijo .1 (DEFAULT)\n"
            "  2     = Tab 4, PF_PROYECTO_2, sufijo .2\n"
            "  ambos = ejecuta Tab 3 y Tab 4 secuencialmente"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--config",    default=str(_CONFIG_PATH),
                        help="Ruta al JSON de configuración base")
    parser.add_argument("--tab",       default="1", choices=["1", "2", "ambos"],
                        help="Tab a cargar: 1 (Tab 3/Proyecto 1), 2 (Tab 4/Proyecto 2), ambos")
    parser.add_argument("--sem",       help="Filtrar por semestre (ej: '2024 sem2')")
    parser.add_argument("--ev",        help="Filtrar por evento (ej: 'Evento 1')")
    parser.add_argument("--modo", default=None, choices=["1", "2", "3"],
                        help=("Modo de asignación de potencia al disparo: "
                              "1=mantener actual (default config), "
                              "2=manual (requiere pgini_manual en config), "
                              "3=proporcional a pgini respetando Pmax"))
    parser.add_argument("--skip-done", action="store_true",
                        help="Omitir eventos con datos_cargados_*{sufijo}*.xlsx ya existentes")
    parser.add_argument("--no-ajuste", action="store_true",
                        help="Desactivar ajuste post-LF (por defecto siempre activo)")
    parser.add_argument("--dry-run",   action="store_true",
                        help="Solo listar eventos sin ejecutar")
    # Alias deprecado conservado por compatibilidad con scripts existentes
    parser.add_argument("--con-proyecto2", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()

    tab = "ambos" if args.con_proyecto2 else args.tab

    if not _RUNNER.exists():
        print(f"[ERROR] Runner no encontrado: {_RUNNER}")
        sys.exit(1)

    cfg     = _load_config(Path(args.config))
    ajustar = not args.no_ajuste
    if args.modo is not None:
        cfg["modo_disparo"] = args.modo

    tabs_a_ejecutar = ["1", "2"] if tab == "ambos" else [tab]

    # Validar que los proyectos estén configurados
    for t in tabs_a_ejecutar:
        label, cfg_key, _ = _TAB_INFO[t]
        if not cfg.get(cfg_key):
            print(f"[ERROR] {cfg_key} no está definido en la configuración para {label}.")
            sys.exit(1)

    # Descubrir eventos
    eventos_base = _discover_eventos(cfg["RAIZ"], sem_filter=args.sem, ev_filter=args.ev)
    if not eventos_base:
        print(f"[ERROR] No se encontraron eventos en: {cfg['RAIZ']}")
        sys.exit(1)

    # Filtrar excluidos explícitamente en config
    _excluir = {(e["semestre"], e["evento"]) for e in cfg.get("excluir_eventos", [])}
    excluidos_cfg = [ev for ev in eventos_base if (ev["semestre"], ev["evento"]) in _excluir]
    eventos_base  = [ev for ev in eventos_base if (ev["semestre"], ev["evento"]) not in _excluir]

    # ── Encabezado ────────────────────────────────────────────────────────────
    print(f"\n{'='*65}")
    print(f"  BATCH CARGA CONDICIONES PF")
    for t in tabs_a_ejecutar:
        label, cfg_key, ev_suffix = _TAB_INFO[t]
        print(f"  {label}: {cfg.get(cfg_key, '—')}  (sufijo '{ev_suffix}')")
    _modo_desc = {"1": "mantener actual", "2": "manual", "3": "proporcional p_desc/pgini"}
    print(f"  Caso base:       {cfg['CASO_BASE']}")
    print(f"  Modo disparo:    {cfg['modo_disparo']} — {_modo_desc.get(str(cfg['modo_disparo']), '?')}")
    print(f"  Ajuste post-LF:  {'SI (siempre)' if ajustar else 'NO (--no-ajuste)'}")
    print(f"  Eventos totales: {len(eventos_base) + len(excluidos_cfg)}")
    if excluidos_cfg:
        print(f"  Excluidos (config): {len(excluidos_cfg)}")
        for _ex in excluidos_cfg:
            print(f"    x  {_ex['semestre']} / {_ex['evento']}")
    print(f"{'='*65}\n")

    if args.dry_run:
        for t in tabs_a_ejecutar:
            label, cfg_key, ev_suffix = _TAB_INFO[t]
            print(f"  {label}:")
            for i, ev in enumerate(eventos_base, 1):
                done = " [YA HECHO]" if _ya_procesado(ev["ev_path"], ev_suffix) else ""
                print(f"    [{i:02d}] {ev['semestre']} / {ev['evento']}{done}")
            print()
        return

    if not eventos_base:
        print("  Nada que procesar.")
        return

    # ── Ejecución por tab ─────────────────────────────────────────────────────
    all_results = []
    t_batch_ini = datetime.now()

    for t in tabs_a_ejecutar:
        label, cfg_key, ev_suffix = _TAB_INFO[t]
        print(f"\n{'─'*65}")
        print(f"  {label}  →  {cfg[cfg_key]}")
        print(f"{'─'*65}\n")

        resultados = _procesar_tab(
            cfg=cfg,
            eventos=list(eventos_base),   # copia fresca por cada tab
            tab_id=t,
            ajustar=ajustar,
            skip_done=args.skip_done,
        )
        all_results.extend(resultados)

    dur_total = (datetime.now() - t_batch_ini).total_seconds()

    # ── Resumen ────────────────────────────────────────────────────────────────
    ok  = [r for r in all_results if r["rc"] == 0]
    err = [r for r in all_results if r["rc"] != 0]

    print(f"\n{'='*65}")
    print(f"  RESUMEN FINAL")
    print(f"  Procesados: {len(all_results)}  |  OK: {len(ok)}  |  Errores: {len(err)}")
    print(f"  Tiempo total: {dur_total/60:.1f} min  (~{dur_total/max(len(all_results),1):.0f}s/evento)")
    print(f"{'='*65}")

    if err:
        print(f"\n  Eventos con errores ({len(err)}):")
        for r in err:
            print(f"    x  {r['tab']} | {r['semestre']} / {r['evento']}  (rc={r['rc']}, {r['dur']:.0f}s)")
        print(f"\n  Logs en: {_LOG_DIR}")
    else:
        print("\n  Todos los eventos cargados correctamente.")

    # Guardar resumen JSON
    resumen_path = _LOG_DIR / f"resumen_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    _LOG_DIR.mkdir(parents=True, exist_ok=True)
    with open(resumen_path, "w", encoding="utf-8") as f:
        json.dump({
            "fecha":           datetime.now().isoformat(),
            "tabs":            tabs_a_ejecutar,
            "ajustar_post_lf": ajustar,
            "total":           len(all_results),
            "ok":              len(ok),
            "errores":         len(err),
            "resultados":      all_results,
        }, f, ensure_ascii=False, indent=2)
    print(f"  Resumen guardado: {resumen_path}\n")

    sys.exit(1 if err else 0)


if __name__ == "__main__":
    main()
