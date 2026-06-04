#!/usr/bin/env python3
"""
generate_context.py — Capa 3: Genera documento de contexto RPF para Ollama/Open WebUI

Lee rpf_kpi_cobee en PostgreSQL y produce un Markdown estructurado con:
- Resumen ejecutivo del parque generador COBEE
- Estadísticas históricas por unidad
- Detalle por evento (frecuencia, reservas, cumplimiento)

Salida: /home/joselozano/rpf-ejecucion/context/rpf_context.md

Uso:
  python3 generate_context.py
  python3 generate_context.py --output /ruta/custom.md
"""

import os
import sys
import argparse
import logging
from datetime import datetime, timezone
from pathlib import Path

import psycopg2
import psycopg2.extras
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

DEFAULT_OUTPUT = Path(__file__).parent / 'context' / 'rpf_context.md'


def get_conn():
    return psycopg2.connect(
        host=os.getenv('POSTGRES_HOST', 'localhost'),
        port=int(os.getenv('POSTGRES_PORT', '5432')),
        dbname=os.getenv('POSTGRES_DB', 'rpf_intelligence'),
        user=os.getenv('POSTGRES_USER', 'n8n'),
        password=os.getenv('POSTGRES_PASSWORD', ''),
    )


def query(conn, sql, params=None):
    with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute(sql, params)
        return cur.fetchall()


# ─── Secciones del documento ─────────────────────────────────────────────────

def sec_header(now: datetime) -> str:
    return f"""# Contexto RPF COBEE — Sistema de Inteligencia Energética
> Generado automáticamente el {now.strftime('%Y-%m-%d %H:%M')} UTC
> Fuente: base de datos `rpf_intelligence` (tabla `rpf_kpi_cobee`)
> Este documento es el contexto de referencia para análisis de Regulación Primaria de Frecuencia (RPF)
> de las unidades generadoras de COBEE en el Sistema Interconectado Nacional (SIN) de Bolivia.

---
"""


def sec_glosario() -> str:
    return """## Glosario de términos

| Término | Significado |
|---------|-------------|
| RPF | Regulación Primaria de Frecuencia — respuesta automática de gobernadores ante caídas de frecuencia |
| P_max | Potencia máxima instalada de la unidad [MW] |
| P_0 | Potencia activa antes del evento (pre-falla) [MW] |
| P_35 | Potencia activa a los ~35 segundos del evento (ventana RPF) [MW] |
| R. inicial | Reserva RPF disponible al inicio del evento = P_max − P_0 [MW o %] |
| P. entregada | Potencia adicional real entregada durante el evento [MW o %] |
| Aporta RPF | Si la unidad respondió correctamente: "Sí", "No", o "Pot. máx" (estaba a plena carga) |
| droop inf. | Estatismo declarado por la empresa ante el CNDC [%] |
| droop calc. | Estatismo real calculado a partir de la respuesta medida [%] |
| f_0 | Frecuencia del sistema antes del evento [Hz] |
| f_min | Frecuencia mínima alcanzada durante el nadir [Hz] |
| f_35 | Frecuencia a los ~35 segundos del evento [Hz] |
| t_0 | Hora del inicio del evento (disparo de la falla) |
| t_min | Hora del nadir de frecuencia |
| Semestre | Período de evaluación RPF: formato YYYY_semN (ej: 2025_sem1 = primer semestre 2025) |
| Evento | Perturbación eléctrica analizada dentro de un semestre |

---
"""


def sec_resumen(conn) -> str:
    rows = query(conn, """
        SELECT
            COUNT(DISTINCT semestre || evento) AS total_eventos,
            COUNT(DISTINCT semestre)           AS total_semestres,
            COUNT(DISTINCT unidad)             AS total_unidades,
            MIN(fecha_evento)                  AS fecha_inicio,
            MAX(fecha_evento)                  AS fecha_fin,
            ROUND(AVG(f_min_hz)::numeric, 3)   AS f_min_promedio,
            ROUND(MIN(f_min_hz)::numeric, 3)    AS f_min_historico,
            ROUND(AVG(r_inicial_mw)::numeric, 2) AS reserva_media_mw
        FROM rpf_kpi_cobee
    """)
    r = rows[0]

    aporta = query(conn, """
        SELECT
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'Sí' THEN 1 ELSE 0 END) / COUNT(*), 1) AS pct_si,
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'No' THEN 1 ELSE 0 END) / COUNT(*), 1) AS pct_no,
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'Pot. máx' THEN 1 ELSE 0 END) / COUNT(*), 1) AS pct_potmax
        FROM rpf_kpi_cobee
        WHERE aporta_rpf IS NOT NULL
    """)
    a = aporta[0]

    return f"""## Resumen ejecutivo

| Métrica | Valor |
|---------|-------|
| Período analizado | {r['fecha_inicio']} → {r['fecha_fin']} |
| Semestres evaluados | {r['total_semestres']} |
| Eventos analizados | {r['total_eventos']} |
| Unidades generadoras COBEE | {r['total_unidades']} |
| f_min promedio histórico | {r['f_min_promedio']} Hz |
| f_min más bajo registrado | {r['f_min_historico']} Hz |
| Reserva RPF media por unidad | {r['reserva_media_mw']} MW |

### Cumplimiento RPF histórico (todas las unidades, todos los eventos)

- **Sí aportaron RPF:** {a['pct_si']}%
- **No aportaron RPF:** {a['pct_no']}%
- **Estaban a potencia máxima (sin reserva disponible):** {a['pct_potmax']}%

---
"""


def sec_por_semestre(conn) -> str:
    rows = query(conn, """
        SELECT
            semestre,
            COUNT(DISTINCT evento)  AS n_eventos,
            ROUND(AVG(f_min_hz)::numeric, 3) AS f_min_avg,
            ROUND(MIN(f_min_hz)::numeric, 3) AS f_min_worst,
            ROUND(AVG(r_inicial_mw)::numeric, 2) AS reserva_avg,
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'Sí' THEN 1 ELSE 0 END)
                  / NULLIF(SUM(CASE WHEN aporta_rpf IS NOT NULL THEN 1 ELSE 0 END), 0), 1) AS pct_cumple
        FROM rpf_kpi_cobee
        GROUP BY semestre
        ORDER BY semestre
    """)

    lines = ["## Estadísticas por semestre\n"]
    lines.append("| Semestre | Eventos | f_min prom [Hz] | f_min peor [Hz] | Reserva prom [MW] | % Cumple RPF |")
    lines.append("|----------|---------|-----------------|-----------------|-------------------|-------------|")
    for r in rows:
        lines.append(
            f"| {r['semestre']} | {r['n_eventos']} | {r['f_min_avg']} | {r['f_min_worst']} "
            f"| {r['reserva_avg']} | {r['pct_cumple']}% |"
        )
    lines.append("\n---\n")
    return "\n".join(lines)


def sec_por_unidad(conn) -> str:
    rows = query(conn, """
        SELECT
            unidad,
            COUNT(DISTINCT semestre || evento)  AS n_eventos,
            ROUND(AVG(p_max_mw)::numeric, 2)    AS p_max_avg,
            ROUND(AVG(r_inicial_mw)::numeric, 2) AS reserva_avg,
            ROUND(AVG(droop_calc_pct)::numeric, 2) AS droop_avg,
            ROUND(AVG(droop_inf_pct)::numeric, 2)  AS droop_inf_avg,
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'Sí' THEN 1 ELSE 0 END)
                  / NULLIF(SUM(CASE WHEN aporta_rpf IS NOT NULL THEN 1 ELSE 0 END), 0), 1) AS pct_si,
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'No' THEN 1 ELSE 0 END)
                  / NULLIF(SUM(CASE WHEN aporta_rpf IS NOT NULL THEN 1 ELSE 0 END), 0), 1) AS pct_no,
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'Pot. máx' THEN 1 ELSE 0 END)
                  / NULLIF(SUM(CASE WHEN aporta_rpf IS NOT NULL THEN 1 ELSE 0 END), 0), 1) AS pct_potmax
        FROM rpf_kpi_cobee
        GROUP BY unidad
        ORDER BY pct_no DESC NULLS LAST, unidad
    """)

    lines = ["## Desempeño histórico por unidad generadora\n"]
    lines.append("Ordenado por mayor porcentaje de incumplimiento (No aporta RPF).\n")
    lines.append("| Unidad | Eventos | P_max prom [MW] | Reserva prom [MW] | Droop calc [%] | Droop inf [%] | % Sí | % No | % Pot.máx |")
    lines.append("|--------|---------|-----------------|-------------------|----------------|---------------|------|------|-----------|")
    for r in rows:
        lines.append(
            f"| {r['unidad']} | {r['n_eventos']} | {r['p_max_avg']} | {r['reserva_avg']} "
            f"| {r['droop_avg'] or '—'} | {r['droop_inf_avg'] or '—'} "
            f"| {r['pct_si'] or 0}% | {r['pct_no'] or 0}% | {r['pct_potmax'] or 0}% |"
        )
    lines.append("\n---\n")
    return "\n".join(lines)


def sec_por_evento(conn) -> str:
    eventos = query(conn, """
        SELECT DISTINCT semestre, evento, fecha_evento,
            ROUND(AVG(f_0_hz)::numeric, 3)  AS f_0,
            ROUND(AVG(f_min_hz)::numeric, 3) AS f_min,
            ROUND(AVG(f_35_hz)::numeric, 3)  AS f_35,
            COUNT(DISTINCT unidad)           AS n_unidades,
            ROUND(SUM(r_inicial_mw)::numeric, 2) AS reserva_total,
            ROUND(SUM(p_entregada_mw)::numeric, 2) AS entregada_total
        FROM rpf_kpi_cobee
        GROUP BY semestre, evento, fecha_evento
        ORDER BY semestre, evento
    """)

    lines = ["## Detalle por evento\n"]

    for ev in eventos:
        sem = ev['semestre']
        evn = ev['evento']
        fecha = ev['fecha_evento']

        # Unidades que no aportaron
        no_aportan = query(conn, """
            SELECT unidad, p_max_mw, p_0_mw, r_inicial_mw, aporta_rpf
            FROM rpf_kpi_cobee
            WHERE semestre=%s AND evento=%s AND aporta_rpf='No'
            ORDER BY unidad
        """, (sem, evn))

        # Unidades que sí aportaron
        si_aportan = query(conn, """
            SELECT unidad, r_inicial_mw, p_entregada_mw, droop_calc_pct
            FROM rpf_kpi_cobee
            WHERE semestre=%s AND evento=%s AND aporta_rpf='Sí'
            ORDER BY unidad
        """, (sem, evn))

        lines.append(f"### {sem} — {evn} ({fecha})\n")
        lines.append(f"- **Frecuencia:** f_0={ev['f_0']} Hz | f_min={ev['f_min']} Hz | f_35={ev['f_35']} Hz")
        lines.append(f"- **Unidades analizadas:** {ev['n_unidades']}")
        lines.append(f"- **Reserva total disponible:** {ev['reserva_total']} MW")
        lines.append(f"- **Potencia total entregada:** {ev['entregada_total']} MW\n")

        if si_aportan:
            lines.append(f"**Unidades que SÍ aportaron RPF ({len(si_aportan)}):**")
            for u in si_aportan:
                droop = f", droop calc={u['droop_calc_pct']}%" if u['droop_calc_pct'] else ""
                lines.append(
                    f"- {u['unidad']}: reserva={u['r_inicial_mw']} MW, "
                    f"entregada={u['p_entregada_mw']} MW{droop}"
                )
            lines.append("")

        if no_aportan:
            lines.append(f"**Unidades que NO aportaron RPF ({len(no_aportan)}):**")
            for u in no_aportan:
                lines.append(
                    f"- {u['unidad']}: P_max={u['p_max_mw']} MW, "
                    f"P_0={u['p_0_mw']} MW, reserva={u['r_inicial_mw']} MW"
                )
            lines.append("")

        lines.append("")

    lines.append("---\n")
    return "\n".join(lines)


def sec_unidades_problematicas(conn) -> str:
    rows = query(conn, """
        SELECT
            unidad,
            COUNT(DISTINCT semestre || evento) AS total_eventos,
            SUM(CASE WHEN aporta_rpf = 'No' THEN 1 ELSE 0 END) AS veces_no,
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'No' THEN 1 ELSE 0 END)
                  / NULLIF(COUNT(*), 0), 1) AS pct_no,
            ROUND(AVG(droop_calc_pct)::numeric, 2)  AS droop_calc,
            ROUND(AVG(droop_inf_pct)::numeric, 2)   AS droop_inf
        FROM rpf_kpi_cobee
        WHERE aporta_rpf IS NOT NULL
        GROUP BY unidad
        HAVING SUM(CASE WHEN aporta_rpf = 'No' THEN 1 ELSE 0 END) > 0
        ORDER BY pct_no DESC
    """)

    lines = ["## Unidades con incumplimiento RPF frecuente\n"]
    lines.append("Unidades que no aportaron RPF en al menos un evento, ordenadas por frecuencia de incumplimiento.\n")
    lines.append("| Unidad | Veces no aportó | Total eventos | % Incumplimiento | Droop calc [%] | Droop inf [%] |")
    lines.append("|--------|----------------|---------------|-----------------|----------------|---------------|")
    for r in rows:
        lines.append(
            f"| {r['unidad']} | {r['veces_no']} | {r['total_eventos']} | {r['pct_no']}% "
            f"| {r['droop_calc'] or '—'} | {r['droop_inf'] or '—'} |"
        )
    lines.append("\n---\n")
    return "\n".join(lines)


def sec_patrones_diagnostico(conn) -> str:
    """Análisis automático de patrones para guiar al LLM en diagnósticos."""

    # Unidades con droop calculado muy diferente al declarado
    droop_rows = query(conn, """
        SELECT unidad,
            ROUND(AVG(droop_inf_pct)::numeric, 2)  AS droop_inf,
            ROUND(AVG(droop_calc_pct)::numeric, 2) AS droop_calc,
            ROUND(AVG(droop_calc_pct - droop_inf_pct)::numeric, 2) AS desviacion,
            COUNT(*) AS n
        FROM rpf_kpi_cobee
        WHERE droop_inf_pct IS NOT NULL AND droop_calc_pct IS NOT NULL
        GROUP BY unidad
        HAVING ABS(AVG(droop_calc_pct - droop_inf_pct)) > 15
        ORDER BY ABS(AVG(droop_calc_pct - droop_inf_pct)) DESC
    """)

    # Eventos con peores nadires de frecuencia
    worst_freq = query(conn, """
        SELECT semestre, evento, fecha_evento,
            ROUND(MIN(f_min_hz)::numeric, 3) AS f_min,
            COUNT(CASE WHEN aporta_rpf = 'No' THEN 1 END) AS unidades_no_aportaron,
            COUNT(DISTINCT unidad) AS total_unidades
        FROM rpf_kpi_cobee
        GROUP BY semestre, evento, fecha_evento
        ORDER BY MIN(f_min_hz)
        LIMIT 5
    """)

    # Unidades con 100% cumplimiento
    perfectas = query(conn, """
        SELECT unidad, COUNT(*) AS eventos
        FROM rpf_kpi_cobee
        WHERE aporta_rpf IS NOT NULL
        GROUP BY unidad
        HAVING SUM(CASE WHEN aporta_rpf != 'Sí' THEN 1 ELSE 0 END) = 0
        ORDER BY COUNT(*) DESC
    """)

    # Correlación: eventos con más no-cumplimiento
    eventos_criticos = query(conn, """
        SELECT semestre, evento,
            COUNT(CASE WHEN aporta_rpf = 'No' THEN 1 END) AS n_no,
            COUNT(DISTINCT unidad) AS total,
            ROUND(100.0 * COUNT(CASE WHEN aporta_rpf = 'No' THEN 1 END)
                / NULLIF(COUNT(DISTINCT unidad), 0), 1) AS pct_no
        FROM rpf_kpi_cobee
        WHERE aporta_rpf IS NOT NULL
        GROUP BY semestre, evento
        ORDER BY pct_no DESC
        LIMIT 5
    """)

    lines = ["## Patrones y Diagnósticos Automáticos\n"]
    lines.append("Esta sección identifica automáticamente patrones anómalos para facilitar el análisis.\n")

    # Droop desviado
    if droop_rows:
        lines.append("### Unidades con Estatismo (Droop) Fuera de Rango\n")
        lines.append("Unidades donde el droop calculado difiere >15% respecto al declarado — indica posible desajuste del gobernador.\n")
        lines.append("| Unidad | Droop declarado [%] | Droop calculado [%] | Desviación [pp] | Eventos |")
        lines.append("|--------|--------------------|--------------------|-----------------|---------|")
        for r in droop_rows:
            signo = "↑" if r['desviacion'] > 0 else "↓"
            lines.append(f"| {r['unidad']} | {r['droop_inf']} | {r['droop_calc']} | {signo}{abs(r['desviacion'])} | {r['n']} |")
        lines.append("")

    # Eventos críticos
    if eventos_criticos:
        lines.append("### Eventos con Mayor Incumplimiento\n")
        lines.append("| Semestre | Evento | Unidades sin RPF | Total | % Incumplimiento |")
        lines.append("|----------|--------|-----------------|-------|-----------------|")
        for r in eventos_criticos:
            lines.append(f"| {r['semestre']} | {r['evento']} | {r['n_no']} | {r['total']} | {r['pct_no']}% |")
        lines.append("")

    # Peores nadires
    if worst_freq:
        lines.append("### Eventos con Frecuencia Mínima Más Baja (Mayor Riesgo)\n")
        lines.append("| Semestre | Evento | Fecha | f_min [Hz] | Unidades sin RPF |")
        lines.append("|----------|--------|-------|-----------|-----------------|")
        for r in worst_freq:
            lines.append(
                f"| {r['semestre']} | {r['evento']} | {r['fecha_evento']} "
                f"| **{r['f_min']}** | {r['unidades_no_aportaron']}/{r['total_unidades']} |"
            )
        lines.append("")

    # Unidades perfectas
    if perfectas:
        names = ", ".join(f"**{r['unidad']}** ({r['eventos']} eventos)" for r in perfectas)
        lines.append(f"### Unidades con 100% de Cumplimiento RPF\n")
        lines.append(f"Estas unidades aportaron RPF en TODOS sus eventos analizados: {names}\n")

    lines.append("\n---\n")
    return "\n".join(lines)


# ─── Entry point ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Genera contexto RPF para Ollama')
    parser.add_argument('--output', type=Path, default=DEFAULT_OUTPUT,
                        help='Ruta de salida del archivo Markdown')
    args = parser.parse_args()

    now = datetime.now(timezone.utc)
    log.info("Generando contexto RPF...")

    conn = get_conn()
    try:
        sections = [
            sec_header(now),
            sec_glosario(),
            sec_resumen(conn),
            sec_por_semestre(conn),
            sec_por_unidad(conn),
            sec_unidades_problematicas(conn),
            sec_patrones_diagnostico(conn),
            sec_por_evento(conn),
        ]
        content = "\n".join(sections)
    finally:
        conn.close()

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(content, encoding='utf-8')

    size_kb = args.output.stat().st_size / 1024
    log.info(f"Contexto generado: {args.output} ({size_kb:.1f} KB)")

    # Exportar CSV para acceso remoto vía SharePoint
    _export_csv(args.output.parent / 'rpf_kpi_cobee.csv')


def _export_csv(csv_path: Path):
    """Exporta toda la tabla rpf_kpi_cobee como CSV para uso remoto."""
    try:
        import csv as csv_mod
        conn = get_conn()
        rows = query(conn, "SELECT * FROM rpf_kpi_cobee ORDER BY semestre, evento, unidad")
        conn.close()
        if not rows:
            return
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv_mod.DictWriter(f, fieldnames=rows[0].keys())
            writer.writeheader()
            writer.writerows(rows)
        log.info(f"CSV exportado: {csv_path} ({len(rows)} registros)")
    except Exception as e:
        log.warning(f"No se pudo exportar CSV: {e}")


if __name__ == '__main__':
    main()
