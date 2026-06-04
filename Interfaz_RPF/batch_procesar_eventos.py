"""
batch_procesar_eventos.py — Ejecuta SCADA y EMF para todos los eventos

Uso:
  python batch_procesar_eventos.py                    # todos los semestres/eventos
  python batch_procesar_eventos.py --sem "2025 sem1"  # solo un semestre
  python batch_procesar_eventos.py --skip-scada       # solo EMF
  python batch_procesar_eventos.py --skip-emf         # solo SCADA
  python batch_procesar_eventos.py --dry-run          # ver lista sin ejecutar
"""

import os
import sys
import json
import subprocess
import argparse
import tempfile
from pathlib import Path
from datetime import datetime

# ─── Rutas ───────────────────────────────────────────────────────────────────
RAIZ_RPF   = r"C:\Datos del CNDC\01_INFO CNDC_RPF"
RAIZ_DATOS = r"C:\Datos del CNDC\02_DATOS CNDC_RPF"
RUNNERS_DIR = Path(__file__).parent / "runners"

SCADA_RUNNER = RUNNERS_DIR / "OrdenadorDatosEvento_run.py"
EMF_RUNNER   = RUNNERS_DIR / "ExtractorResultadosCNDC_run.py"
CARPETA_COBEE = "Resultados_COBEE"

LOG_DIR = Path(__file__).parent / "logs_batch"


# ─── Descubrimiento de eventos ────────────────────────────────────────────────

def discover_eventos(raiz: str, sem_filter: str = None) -> list[dict]:
    """Retorna lista de {semestre, evento, ev_path} ordenados."""
    eventos = []
    base = Path(raiz)
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
            if ev_dir.is_dir() and ev_dir.name.startswith("Evento"):
                eventos.append({
                    "semestre": sem_dir.name,
                    "evento":   ev_dir.name,
                    "ev_path":  str(ev_dir),
                })
    return eventos


# ─── Ejecución de un runner ───────────────────────────────────────────────────

def run_script(runner: Path, params: dict, label: str, log_path: Path) -> int:
    """Escribe params.json, lanza el runner y retorna el código de salida."""
    with tempfile.NamedTemporaryFile(mode="w", suffix=".json",
                                     encoding="utf-8", delete=False) as tf:
        json.dump(params, tf, ensure_ascii=False, indent=2)
        params_path = tf.name

    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    log_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        result = subprocess.run(
            [sys.executable, str(runner), params_path],
            capture_output=True, text=True,
            encoding="utf-8", errors="replace",
            env=env, timeout=300,
        )
        rc = result.returncode
        output = result.stdout + result.stderr
    except subprocess.TimeoutExpired:
        rc = -1
        output = "[ERROR] Timeout después de 5 minutos"
    except Exception as e:
        rc = -1
        output = f"[ERROR] {e}"
    finally:
        try:
            os.remove(params_path)
        except OSError:
            pass

    with open(log_path, "w", encoding="utf-8") as lf:
        lf.write(f"=== {label} ===\n")
        lf.write(f"Fecha: {datetime.now().isoformat()}\n")
        lf.write(f"Código de salida: {rc}\n\n")
        lf.write(output)

    return rc


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Batch SCADA + EMF para todos los eventos")
    parser.add_argument("--sem",        help="Filtrar por semestre (ej: '2025 sem1')")
    parser.add_argument("--skip-scada", action="store_true", help="Omitir extractor SCADA")
    parser.add_argument("--skip-emf",   action="store_true", help="Omitir extractor EMF")
    parser.add_argument("--dry-run",    action="store_true", help="Solo listar, sin ejecutar")
    args = parser.parse_args()

    if not SCADA_RUNNER.exists():
        print(f"[ERROR] No se encontró: {SCADA_RUNNER}")
        sys.exit(1)
    if not EMF_RUNNER.exists():
        print(f"[ERROR] No se encontró: {EMF_RUNNER}")
        sys.exit(1)

    eventos = discover_eventos(RAIZ_RPF, sem_filter=args.sem)
    if not eventos:
        print(f"[ERROR] No se encontraron eventos en: {RAIZ_RPF}")
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"  BATCH RPF — {len(eventos)} eventos encontrados")
    print(f"  SCADA: {'SKIP' if args.skip_scada else 'ON'} | EMF: {'SKIP' if args.skip_emf else 'ON'}")
    print(f"{'='*60}\n")

    results = []
    for i, ev in enumerate(eventos, 1):
        sem = ev["semestre"]
        evento = ev["evento"]
        label = f"{sem} / {evento}"
        print(f"[{i:02d}/{len(eventos)}] {label}", end="", flush=True)

        if args.dry_run:
            print("  → DRY RUN")
            continue

        ev_results = {"semestre": sem, "evento": evento, "scada": None, "emf": None}

        # SCADA
        if not args.skip_scada:
            scada_params = {
                "semestre":   sem,
                "evento":     evento,
                "RAIZ_RPF":   RAIZ_RPF,
                "RAIZ_DATOS": RAIZ_DATOS,
            }
            log_path = LOG_DIR / sem.replace(" ", "_") / f"{evento.replace(' ', '_')}_scada.log"
            rc = run_script(SCADA_RUNNER, scada_params, f"SCADA {label}", log_path)
            ev_results["scada"] = rc
            status = "✓" if rc == 0 else f"✗ (rc={rc})"
            print(f"  SCADA:{status}", end="", flush=True)

        # EMF
        if not args.skip_emf:
            emf_params = {
                "semestre":     sem,
                "evento":       evento,
                "RAIZ":         RAIZ_RPF,
                "CARPETA_COBEE": CARPETA_COBEE,
            }
            log_path = LOG_DIR / sem.replace(" ", "_") / f"{evento.replace(' ', '_')}_emf.log"
            rc = run_script(EMF_RUNNER, emf_params, f"EMF {label}", log_path)
            ev_results["emf"] = rc
            status = "✓" if rc == 0 else f"✗ (rc={rc})"
            print(f"  EMF:{status}", end="", flush=True)

        print()
        results.append(ev_results)

    if args.dry_run:
        return

    # Resumen
    print(f"\n{'='*60}")
    print("  RESUMEN")
    print(f"{'='*60}")
    ok_scada = sum(1 for r in results if r["scada"] == 0)
    ok_emf   = sum(1 for r in results if r["emf"] == 0)
    err = [r for r in results if r["scada"] not in (0, None) or r["emf"] not in (0, None)]

    if not args.skip_scada:
        print(f"  SCADA: {ok_scada}/{len(results)} OK")
    if not args.skip_emf:
        print(f"  EMF:   {ok_emf}/{len(results)} OK")

    if err:
        print(f"\n  Eventos con errores ({len(err)}):")
        for r in err:
            print(f"    • {r['semestre']} / {r['evento']} — SCADA:{r['scada']} EMF:{r['emf']}")
        print(f"\n  Logs en: {LOG_DIR}")
    else:
        print("\n  Todos los eventos procesados correctamente.")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
