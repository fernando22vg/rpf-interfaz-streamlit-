"""
batch_carga_cond_pf.py — Carga condiciones iniciales en PowerFactory para todos los eventos

Uso:
  python batch_carga_cond_pf.py                           # todos los semestres/eventos
  python batch_carga_cond_pf.py --sem "2024 sem2"         # solo un semestre
  python batch_carga_cond_pf.py --ev "Evento 1"           # solo un evento específico
  python batch_carga_cond_pf.py --skip-done               # omite eventos ya procesados
  python batch_carga_cond_pf.py --dry-run                 # ver lista sin ejecutar
  python batch_carga_cond_pf.py --con-proyecto2           # carga también en PF_PROYECTO_2 (sufijo .2)
  python batch_carga_cond_pf.py --config ruta/config.json

El ajuste post-LF (AJUSTAR_POST_LF) se activa siempre para todos los eventos.
Para desactivarlo en casos específicos usar --no-ajuste.

Sin límite de tiempo por evento (PowerFactory puede tardar lo que necesite).

Para excluir eventos específicos, usar "excluir_eventos" en batch_config_pf.json:
  "excluir_eventos": [{"semestre": "2024 sem2", "evento": "Evento 1"}]

Configuración base: batch_config_pf.json en el mismo directorio (se crea si no existe).
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
    "PF_PROYECTO":     "PMP_NOV25_OCT29_31102025(1)",
    "PF_PROYECTO_2":   "PMP_NOV25_OCT29_31102025(2)",   # segundo proyecto (--con-proyecto2)
    "CASO_BASE":       "CNDC",
    "modo_disparo":    "1",
    "guardar_escenario": True,
    "excluir_slack":   [],
    "xfo_pf":          1.0,
    "excluir_eventos": [],   # [{"semestre": "2024 sem2", "evento": "Evento 1"}, ...]
}

_CONFIG_PATH  = _HERE / "batch_config_pf.json"
_RUNNER       = _HERE / "runners" / "CargaCondIniciales_PF_run.py"
_LOG_DIR      = _HERE / "logs_batch_pf"


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _load_config(config_path: Path) -> dict:
    """Lee la config base. Si no existe la crea con los defaults y avisa."""
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        return {**_DEFAULTS, **cfg}
    # Primera vez: crear archivo de referencia
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


def _ya_procesado(ev_path: str) -> bool:
    """True si ya existe datos_cargados_*.xlsx en la carpeta del evento."""
    import glob as _glob
    patron = os.path.join(ev_path, "datos_cargados_*.xlsx")
    return len(_glob.glob(patron)) > 0


def _run_evento(cfg: dict, sem: str, evento: str, log_path: Path, ajustar: bool,
                pf_proyecto: str = None, ev_suffix: str = "") -> tuple:
    """Ejecuta el runner para un evento. Retorna (rc, duracion_seg).

    pf_proyecto: sobreescribe cfg['PF_PROYECTO'] (para el segundo proyecto).
    ev_suffix:   sufijo que el runner añade al nombre de archivo (ej. '.2').
    """
    params = {
        "semestre":          sem,
        "evento":            evento,
        "RAIZ":              cfg["RAIZ"],
        "PF_BASE":           cfg["PF_BASE"],
        "LOC_GEN_PATH":      cfg["LOC_GEN_PATH"],
        "LOC_CAR_PATH":      cfg.get("LOC_CAR_PATH", ""),
        "LOC_XFO_PATH":      cfg["LOC_XFO_PATH"],
        "PF_PROYECTO":       pf_proyecto if pf_proyecto else cfg["PF_PROYECTO"],
        "CASO_BASE":         cfg["CASO_BASE"],
        "modo_disparo":      str(cfg.get("modo_disparo", "1")),
        "pgini_manual":      cfg.get("pgini_manual", {}),
        "ajustar_post_lf":   ajustar,
        "guardar_escenario": cfg.get("guardar_escenario", True),
        "keep_pf_open":      False,   # batch: siempre cerrar PF al terminar
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

    # Matar procesos PF residuales ANTES de iniciar (libera licencia de runs anteriores)
    subprocess.run(["taskkill", "/F", "/IM", "PowerFactory.exe", "/T"],
                   capture_output=True, check=False)
    import time as _time
    _time.sleep(5)   # espera a que el servidor de licencias registre la liberación

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
        # Forzar cierre de PF al terminar (crash o éxito) para liberar licencia
        subprocess.run(["taskkill", "/F", "/IM", "PowerFactory.exe", "/T"],
                       capture_output=True, check=False)
        try:
            os.remove(params_path)
        except OSError:
            pass

    dur = (datetime.now() - t_ini).total_seconds()

    _proyecto_log = pf_proyecto if pf_proyecto else cfg["PF_PROYECTO"]
    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write(f"=== {sem} / {evento} ===\n")
        lf.write(f"Proyecto PF:     {_proyecto_log}\n")
        lf.write(f"Sufijo archivo:  '{ev_suffix}'\n")
        lf.write(f"Inicio:          {t_ini.strftime('%Y-%m-%d %H:%M:%S')}\n")
        lf.write(f"Duración:        {dur:.0f}s\n")
        lf.write(f"Código de salida: {rc}\n")
        lf.write(f"Ajuste post-LF:  {ajustar}\n")
        lf.write("\n" + "─" * 60 + "\n\n")
        lf.write(output)

    # PowerFactory a veces crashea al cerrar (0xC0000005 / ACCESS_VIOLATION).
    # Si datos_cargados_*.xlsx existe en la carpeta del evento, la carga fue exitosa.
    _RC_PF_CRASH = {3221225477, -1073741819}   # 0xC0000005 con/sin signo
    if rc in _RC_PF_CRASH:
        import glob as _glob
        _ev_path_real = os.path.join(cfg["RAIZ"], sem, "Análisis_todos_los_eventos", evento)
        if not os.path.isdir(_ev_path_real):
            _ev_path_real = os.path.join(cfg["RAIZ"], sem, "Analisis_todos_los_eventos", evento)
        if _glob.glob(os.path.join(_ev_path_real, "datos_cargados_*.xlsx")):
            with open(log_path, "a", encoding="utf-8") as lf:
                lf.write("\n[INFO] rc=0xC0000005: crash de PF al cerrar, pero datos_cargados_*.xlsx existe → carga OK\n")
            rc = 0   # tratar como éxito

    return rc, dur


# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Carga condiciones iniciales PF para todos los eventos"
    )
    parser.add_argument("--config",     default=str(_CONFIG_PATH),
                        help="Ruta al JSON de configuración base")
    parser.add_argument("--sem",        help="Filtrar por semestre (ej: '2024 sem2')")
    parser.add_argument("--ev",         help="Filtrar por evento (ej: 'Evento 1')")
    parser.add_argument("--skip-done",  action="store_true",
                        help="Omitir eventos que ya tienen datos_cargados_*.xlsx")
    parser.add_argument("--no-ajuste",  action="store_true",
                        help="Desactivar ajuste post-LF (por defecto siempre activo)")
    parser.add_argument("--dry-run",    action="store_true",
                        help="Solo listar eventos sin ejecutar")
    parser.add_argument("--con-proyecto2", action="store_true",
                        help="Ejecutar también la carga en PF_PROYECTO_2 (sufijo .2) por cada evento")
    args = parser.parse_args()

    if not _RUNNER.exists():
        print(f"[ERROR] Runner no encontrado: {_RUNNER}")
        sys.exit(1)

    cfg = _load_config(Path(args.config))
    ajustar = not args.no_ajuste   # True por defecto

    eventos = _discover_eventos(cfg["RAIZ"], sem_filter=args.sem, ev_filter=args.ev)
    if not eventos:
        print(f"[ERROR] No se encontraron eventos en: {cfg['RAIZ']}")
        sys.exit(1)

    # Filtrar eventos excluidos explícitamente en la config
    _excluir = {(e["semestre"], e["evento"]) for e in cfg.get("excluir_eventos", [])}
    excluidos_cfg = [ev for ev in eventos if (ev["semestre"], ev["evento"]) in _excluir]
    eventos       = [ev for ev in eventos if (ev["semestre"], ev["evento"]) not in _excluir]

    # Filtrar ya procesados
    skip_list = []
    if args.skip_done:
        skip_list = [ev for ev in eventos if _ya_procesado(ev["ev_path"])]
        eventos   = [ev for ev in eventos if not _ya_procesado(ev["ev_path"])]

    usar_proyecto2 = args.con_proyecto2
    pf2            = cfg.get("PF_PROYECTO_2", "")

    if usar_proyecto2 and not pf2:
        print("[ERROR] --con-proyecto2 activado pero PF_PROYECTO_2 no está definido en la config.")
        sys.exit(1)

    print(f"\n{'='*65}")
    print(f"  BATCH CARGA CONDICIONES PF")
    print(f"  Proyecto PF:    {cfg['PF_PROYECTO']}")
    if usar_proyecto2:
        print(f"  Proyecto PF 2:  {pf2}  (sufijo .2 — --con-proyecto2 activo)")
    print(f"  Caso base:      {cfg['CASO_BASE']}")
    print(f"  Ajuste post-LF: {'SÍ (siempre)' if ajustar else 'NO (--no-ajuste)'}")
    print(f"  Timeout/evento: sin límite")
    print(f"  Eventos totales: {len(eventos) + len(skip_list) + len(excluidos_cfg)}")
    if excluidos_cfg:
        print(f"  Excluidos (config): {len(excluidos_cfg)}")
        for _ex in excluidos_cfg:
            print(f"    ✗  {_ex['semestre']} / {_ex['evento']}  (excluir_eventos en config)")
    if skip_list:
        print(f"  Omitidos (ya procesados): {len(skip_list)}")
    print(f"  A procesar: {len(eventos)}")
    print(f"{'='*65}\n")

    if args.dry_run:
        for i, ev in enumerate(eventos, 1):
            done = " [YA HECHO]" if _ya_procesado(ev["ev_path"]) else ""
            print(f"  [{i:02d}] {ev['semestre']} / {ev['evento']}{done}")
        if skip_list:
            print(f"\n  Omitidos:")
            for ev in skip_list:
                print(f"       {ev['semestre']} / {ev['evento']}")
        return

    if not eventos:
        print("  Nada que procesar.")
        return

    results = []
    t_batch_ini = datetime.now()

    for i, ev in enumerate(eventos, 1):
        sem    = ev["semestre"]
        evento = ev["evento"]
        label  = f"{sem} / {evento}"

        sem_slug = sem.replace(" ", "_")
        ev_slug  = evento.replace(" ", "_")

        # ── Proyecto 1 ──────────────────────────────────────────────────────────
        print(f"[{i:02d}/{len(eventos)}] {label}  [P1] ...", end="", flush=True)
        log_path = _LOG_DIR / sem_slug / f"{ev_slug}.log"
        rc, dur = _run_evento(cfg, sem, evento, log_path, ajustar)

        if rc == 0:
            estado = "OK"
        elif rc in (3221225477, -1073741819):
            estado = "CRASH-PF (0xC0000005, sin datos_cargados)"
        else:
            estado = f"ERROR (rc={rc})"
        print(f"  {estado}  ({dur:.0f}s)  log→ {log_path.name}", flush=True)
        results.append({"semestre": sem, "evento": evento, "proyecto": cfg["PF_PROYECTO"],
                        "rc": rc, "dur": dur})

        # ── Proyecto 2 (opcional) ───────────────────────────────────────────────
        if usar_proyecto2:
            print(f"[{i:02d}/{len(eventos)}] {label}  [P2] ...", end="", flush=True)
            log_path2 = _LOG_DIR / sem_slug / f"{ev_slug}_p2.log"
            rc2, dur2 = _run_evento(cfg, sem, evento, log_path2, ajustar,
                                    pf_proyecto=pf2, ev_suffix=".2")
            if rc2 == 0:
                estado2 = "OK"
            elif rc2 in (3221225477, -1073741819):
                estado2 = "CRASH-PF (0xC0000005, sin datos_cargados)"
            else:
                estado2 = f"ERROR (rc={rc2})"
            print(f"  {estado2}  ({dur2:.0f}s)  log→ {log_path2.name}", flush=True)
            results.append({"semestre": sem, "evento": evento, "proyecto": pf2,
                            "rc": rc2, "dur": dur2})

    dur_total = (datetime.now() - t_batch_ini).total_seconds()

    # ── Resumen ────────────────────────────────────────────────────────────────
    ok  = [r for r in results if r["rc"] == 0]
    err = [r for r in results if r["rc"] != 0]

    print(f"\n{'='*65}")
    print(f"  RESUMEN FINAL")
    print(f"  Procesados: {len(results)}  |  OK: {len(ok)}  |  Errores: {len(err)}")
    print(f"  Tiempo total: {dur_total/60:.1f} min  (~{dur_total/max(len(results),1):.0f}s/evento)")
    print(f"{'='*65}")

    if err:
        print(f"\n  Eventos con errores ({len(err)}):")
        for r in err:
            print(f"    ✗  {r['semestre']} / {r['evento']}  (rc={r['rc']}, {r['dur']:.0f}s)")
        print(f"\n  Logs en: {_LOG_DIR}")
    else:
        print("\n  Todos los eventos cargados correctamente.")

    # Guardar resumen JSON
    resumen_path = _LOG_DIR / f"resumen_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    _LOG_DIR.mkdir(parents=True, exist_ok=True)
    with open(resumen_path, "w", encoding="utf-8") as f:
        json.dump({
            "fecha":           datetime.now().isoformat(),
            "ajustar_post_lf": ajustar,
            "proyecto_pf":     cfg["PF_PROYECTO"],
            "proyecto_pf_2":   pf2 if usar_proyecto2 else None,
            "con_proyecto2":   usar_proyecto2,
            "total":           len(results),
            "ok":              len(ok),
            "errores":         len(err),
            "resultados":      results,
        }, f, ensure_ascii=False, indent=2)
    print(f"  Resumen guardado: {resumen_path}\n")

    sys.exit(1 if err else 0)


if __name__ == "__main__":
    main()
