#!/usr/bin/env python3
"""
create_qa_dataset.py — Genera dataset Q&A para fine-tuning del modelo COBEE-AI

Lee todos los datos disponibles (PostgreSQL + Excel + SCADA) y genera pares
pregunta-respuesta en formato JSONL para fine-tuning con QLoRA.

Salida: /home/joselozano/rpf-ejecucion/training/qa_dataset.jsonl
        /home/joselozano/rpf-ejecucion/training/qa_dataset_stats.json

Uso:
  python3 create_qa_dataset.py
  python3 create_qa_dataset.py --min-pairs 500  # objetivo mínimo de pares
"""

import os
import json
import random
import argparse
import logging
from pathlib import Path
from datetime import datetime

import psycopg2
import psycopg2.extras
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

OUTPUT_DIR = Path(__file__).parent / 'training'

SYSTEM_PROMPT = """Eres COBEE-AI, asistente técnico experto en Regulación Primaria de Frecuencia (RPF) del Sistema Interconectado Nacional (SIN) de Bolivia, especializado en el análisis de las unidades generadoras de COBEE."""


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


def qa_pair(question: str, answer: str) -> dict:
    return {
        "messages": [
            {"role": "system",    "content": SYSTEM_PROMPT},
            {"role": "user",      "content": question},
            {"role": "assistant", "content": answer},
        ]
    }


# ─── Generadores de pares Q&A ─────────────────────────────────────────────────

def gen_resumen_general(conn) -> list[dict]:
    pairs = []
    rows = query(conn, """
        SELECT
            COUNT(DISTINCT semestre || evento) AS eventos,
            COUNT(DISTINCT semestre) AS semestres,
            COUNT(DISTINCT unidad) AS unidades,
            MIN(fecha_evento) AS inicio, MAX(fecha_evento) AS fin,
            ROUND(AVG(f_min_hz)::numeric,3) AS f_min_avg,
            ROUND(MIN(f_min_hz)::numeric,3) AS f_min_peor,
            ROUND(100.0*SUM(CASE WHEN aporta_rpf='Sí' THEN 1 ELSE 0 END)/COUNT(*),1) AS pct_si
        FROM rpf_kpi_cobee
    """)
    r = rows[0]
    answer = (
        f"En la base de datos de COBEE tenemos {r['eventos']} eventos RPF analizados "
        f"en {r['semestres']} semestres (de {r['inicio']} a {r['fin']}), "
        f"con {r['unidades']} unidades generadoras evaluadas.\n\n"
        f"**Estadísticas clave:**\n"
        f"- Cumplimiento RPF global: **{r['pct_si']}%** de los casos\n"
        f"- Frecuencia mínima promedio: **{r['f_min_avg']} Hz**\n"
        f"- Peor nadir registrado: **{r['f_min_peor']} Hz**"
    )
    pairs.append(qa_pair("¿Cuántos eventos RPF tiene COBEE registrados?", answer))
    pairs.append(qa_pair("Dame un resumen del historial RPF de COBEE", answer))
    pairs.append(qa_pair("¿Cuál es el porcentaje de cumplimiento RPF histórico de COBEE?",
                         f"El cumplimiento RPF histórico de COBEE es del **{r['pct_si']}%**. "
                         f"Esto se calcula sobre {r['eventos']} eventos en {r['semestres']} semestres, "
                         f"evaluando {r['unidades']} unidades generadoras distintas."))
    return pairs


def gen_por_unidad(conn) -> list[dict]:
    pairs = []
    rows = query(conn, """
        SELECT unidad,
            COUNT(DISTINCT semestre||evento) AS n_ev,
            ROUND(AVG(p_max_mw)::numeric,2)    AS p_max,
            ROUND(AVG(r_inicial_mw)::numeric,2) AS reserva,
            ROUND(AVG(droop_inf_pct)::numeric,2) AS droop_inf,
            ROUND(AVG(droop_calc_pct)::numeric,2) AS droop_calc,
            ROUND(100.0*SUM(CASE WHEN aporta_rpf='Sí' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN aporta_rpf IS NOT NULL THEN 1 ELSE 0 END),0),1) AS pct_si,
            ROUND(100.0*SUM(CASE WHEN aporta_rpf='No' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN aporta_rpf IS NOT NULL THEN 1 ELSE 0 END),0),1) AS pct_no
        FROM rpf_kpi_cobee
        GROUP BY unidad
        ORDER BY unidad
    """)
    for r in rows:
        u = r['unidad']
        # Pregunta por unidad específica
        droop_status = ""
        if r['droop_calc'] and r['droop_inf']:
            diff = float(r['droop_calc']) - float(r['droop_inf'])
            if abs(diff) > 15:
                droop_status = f"\n⚠️ **Alerta droop**: el estatismo calculado ({r['droop_calc']}%) difiere significativamente del declarado ({r['droop_inf']}%), lo que indica un posible desajuste del gobernador."
        answer = (
            f"**Unidad {u}** — datos históricos COBEE:\n\n"
            f"- Potencia máxima: **{r['p_max']} MW**\n"
            f"- Reserva RPF promedio disponible: **{r['reserva']} MW**\n"
            f"- Droop declarado al CNDC: **{r['droop_inf']}%**\n"
            f"- Droop calculado real: **{r['droop_calc'] or '—'}%**\n"
            f"- Cumplimiento RPF: **{r['pct_si']}%** (aportó en {r['pct_si']}% de {r['n_ev']} eventos)\n"
            f"- Incumplimiento: **{r['pct_no']}%** (no respondió en {r['pct_no']}% de eventos)"
            f"{droop_status}"
        )
        pairs.append(qa_pair(f"¿Cómo es el desempeño RPF de la unidad {u}?", answer))
        pairs.append(qa_pair(f"¿Cuánto aporta {u} en eventos RPF?", answer))
        if r['pct_no'] and float(r['pct_no']) > 40:
            pairs.append(qa_pair(
                f"¿Por qué {u} tiene problemas con la RPF?",
                f"La unidad **{u}** tiene un historial de incumplimiento del **{r['pct_no']}%** en "
                f"{r['n_ev']} eventos analizados. Esto significa que en {r['pct_no']}% de los eventos "
                f"de frecuencia del SIN, esta unidad NO incrementó su potencia como requiere la RPF.\n\n"
                f"Posibles causas:\n"
                f"1. Gobernador no configurado correctamente (droop: declarado={r['droop_inf']}%, calc={r['droop_calc'] or 'N/D'}%)\n"
                f"2. Operando cerca de su límite máximo ({r['p_max']} MW)\n"
                f"3. Restricciones operativas o mantenimiento en esos eventos\n\n"
                f"Recomendación: revisar la configuración del gobernador y comparar con eventos donde sí aportó."
            ))
    return pairs


def gen_por_evento(conn) -> list[dict]:
    pairs = []
    eventos = query(conn, """
        SELECT DISTINCT semestre, evento, fecha_evento,
            ROUND(MIN(f_min_hz)::numeric,3) AS f_min,
            ROUND(AVG(f_0_hz)::numeric,3) AS f_0,
            COUNT(DISTINCT unidad) AS n_unidades,
            SUM(CASE WHEN aporta_rpf='No' THEN 1 ELSE 0 END) AS n_no,
            SUM(CASE WHEN aporta_rpf='Sí' THEN 1 ELSE 0 END) AS n_si
        FROM rpf_kpi_cobee
        GROUP BY semestre, evento, fecha_evento
        ORDER BY semestre, evento
    """)
    for ev in eventos:
        sem = ev['semestre']
        evn = ev['evento']
        fecha = ev['fecha_evento']
        no_units = query(conn, """
            SELECT unidad FROM rpf_kpi_cobee
            WHERE semestre=%s AND evento=%s AND aporta_rpf='No'
            ORDER BY unidad
        """, (sem, evn))
        si_units = query(conn, """
            SELECT unidad FROM rpf_kpi_cobee
            WHERE semestre=%s AND evento=%s AND aporta_rpf='Sí'
            ORDER BY unidad
        """, (sem, evn))

        no_list = ', '.join(r['unidad'] for r in no_units) or 'ninguna'
        si_list = ', '.join(r['unidad'] for r in si_units) or 'ninguna'

        answer = (
            f"**{sem} — {evn}** ({fecha}):\n\n"
            f"- Frecuencia previa al evento: {ev['f_0']} Hz\n"
            f"- Frecuencia mínima (nadir): **{ev['f_min']} Hz**\n"
            f"- Unidades analizadas: {ev['n_unidades']}\n"
            f"- Unidades que SÍ aportaron RPF ({ev['n_si']}): {si_list}\n"
            f"- Unidades que NO aportaron RPF ({ev['n_no']}): {no_list}"
        )
        pairs.append(qa_pair(f"¿Qué ocurrió en el {evn} del {sem}?", answer))
        pairs.append(qa_pair(
            f"¿Qué unidades no aportaron RPF en el {evn} del {sem}?",
            f"En el **{evn} del {sem}** ({fecha}), las siguientes unidades NO aportaron RPF: **{no_list}**.\n\n"
            f"En total {ev['n_no']} de {ev['n_unidades']} unidades incumplieron. "
            f"La frecuencia cayó hasta **{ev['f_min']} Hz** durante este evento."
        ))
    return pairs


def gen_comparativas(conn) -> list[dict]:
    pairs = []
    # Peores unidades por semestre
    sems = query(conn, "SELECT DISTINCT semestre FROM rpf_kpi_cobee ORDER BY semestre")
    for sem_row in sems:
        sem = sem_row['semestre']
        rows = query(conn, """
            SELECT unidad,
                ROUND(100.0*SUM(CASE WHEN aporta_rpf='No' THEN 1 ELSE 0 END)
                    /NULLIF(COUNT(*),0),1) AS pct_no
            FROM rpf_kpi_cobee
            WHERE semestre=%s AND aporta_rpf IS NOT NULL
            GROUP BY unidad ORDER BY pct_no DESC LIMIT 5
        """, (sem,))
        if rows:
            lista = '\n'.join(f"  {i+1}. **{r['unidad']}**: {r['pct_no']}% incumplimiento"
                              for i, r in enumerate(rows))
            pairs.append(qa_pair(
                f"¿Cuáles son las peores unidades RPF del {sem}?",
                f"En el **{sem}**, las unidades con mayor incumplimiento RPF fueron:\n{lista}"
            ))

    # Tendencia de mejora/empeoramiento
    trend = query(conn, """
        SELECT semestre,
            ROUND(100.0*SUM(CASE WHEN aporta_rpf='Sí' THEN 1 ELSE 0 END)
                /NULLIF(SUM(CASE WHEN aporta_rpf IS NOT NULL THEN 1 ELSE 0 END),0),1) AS pct_si,
            ROUND(MIN(f_min_hz)::numeric,3) AS f_min_peor
        FROM rpf_kpi_cobee
        GROUP BY semestre ORDER BY semestre
    """)
    if len(trend) > 1:
        tabla = '\n'.join(
            f"| {r['semestre']} | {r['pct_si']}% | {r['f_min_peor']} Hz |"
            for r in trend
        )
        pairs.append(qa_pair(
            "¿Ha mejorado el cumplimiento RPF de COBEE con el tiempo?",
            f"Evolución del cumplimiento RPF de COBEE por semestre:\n\n"
            f"| Semestre | % Cumplimiento | f_min peor [Hz] |\n"
            f"|----------|---------------|----------------|\n"
            f"{tabla}\n\n"
            f"{'El cumplimiento mejoró significativamente' if float(trend[-1]['pct_si']) > float(trend[0]['pct_si']) else 'El cumplimiento se ha mantenido variable'} "
            f"desde {trend[0]['semestre']} hasta {trend[-1]['semestre']}."
        ))
    return pairs


def gen_conceptos_tecnicos() -> list[dict]:
    """Genera pares Q&A sobre conceptos técnicos RPF/SIN Bolivia."""
    return [
        qa_pair(
            "¿Qué es el droop o estatismo de una turbina?",
            "El **droop** (o estatismo) es el porcentaje de caída de frecuencia necesario para que una unidad generadora incremente su potencia hasta el máximo disponible. "
            "Por ejemplo, con droop=6% en un sistema de 50 Hz, si la frecuencia cae 3 Hz (6% de 50 Hz), la unidad debería pasar de potencia mínima a máxima.\n\n"
            "En el SIN Bolivia, el CDM (Contrato de Despacho de Mínimo Costo) establece que el droop debe estar entre **6% y 12%**.\n\n"
            "**Droop declarado**: el valor que la empresa informa al CNDC.\n"
            "**Droop calculado**: el valor real medido durante eventos RPF reales.\n\n"
            "Si el droop calculado >> declarado, el gobernador responde más lento de lo esperado y la unidad puede ser penalizada."
        ),
        qa_pair(
            "¿Qué es la RPF?",
            "La **Regulación Primaria de Frecuencia (RPF)** es la respuesta automática e inmediata de los gobernadores de las turbinas ante variaciones de frecuencia en la red eléctrica.\n\n"
            "Cuando una generadora se desconecta bruscamente, la frecuencia del SIN Bolivia cae por debajo de 50 Hz. Las demás unidades deben detectar esta caída y automáticamente aumentar su potencia para compensar el déficit.\n\n"
            "El CNDC evalúa semestralmente si cada unidad respondió correctamente (dentro de los primeros 35 segundos). Una unidad 'aporta RPF' si incrementó su potencia al menos en el porcentaje que indica su droop declarado."
        ),
        qa_pair(
            "¿Qué significa f_min en un evento RPF?",
            "**f_min** es la frecuencia mínima (nadir) alcanzada durante un evento de frecuencia en el SIN Bolivia.\n\n"
            "Cuando una generadora grande se desconecta, la frecuencia cae desde ~50 Hz hasta un mínimo (el nadir) antes de recuperarse gracias a la RPF.\n\n"
            "Valores típicos en el SIN Bolivia:\n"
            "- f_min > 49.5 Hz: evento leve, RPF eficiente\n"
            "- f_min entre 49.3-49.5 Hz: evento moderado\n"
            "- f_min < 49.3 Hz: evento severo, posible riesgo de disparos de carga\n\n"
            "El CNDC usa f_min como indicador de la gravedad del evento y la calidad de la respuesta del sistema."
        ),
        qa_pair(
            "¿Cuál es la diferencia entre P_0 y P_35 en el análisis RPF?",
            "En el análisis RPF de COBEE:\n\n"
            "- **P_0**: Potencia activa de la unidad ANTES del evento (pre-falla), en MW\n"
            "- **P_35**: Potencia activa de la unidad a los ~35 segundos del evento (ventana de evaluación RPF)\n\n"
            "La diferencia **P_35 - P_0** representa la potencia adicional que aportó la unidad durante el evento RPF.\n\n"
            "Si P_35 > P_0: la unidad incrementó su generación → contribuyó a la RPF\n"
            "Si P_35 ≈ P_0: la unidad no respondió → incumplimiento RPF\n"
            "Si P_35 < P_0: la unidad redujo generación → caso anómalo"
        ),
        qa_pair(
            "¿Qué unidades COBEE están en el SIN Bolivia?",
            "COBEE opera las siguientes unidades generadoras en el SIN Bolivia:\n\n"
            "**Zona Zongo (hidroeléctrica):**\n"
            "- ZON (Zongo), BOT01, BOT02, BOT03 (Botijlaca)\n"
            "- CUT01-CUT05 (Cuticucho), SRO01, SRO02 (Sainani)\n\n"
            "**Zona Sur (hidroeléctrica):**\n"
            "- SAI (Sainani), CHU01, CHU02 (Chururaqui)\n"
            "- HAR01, HAR02 (Harca), CAH01, CAH02 (Cahua)\n"
            "- HUA01, HUA02 (Huaji)\n\n"
            "**Otras:**\n"
            "- TIQ (Tiquina)\n\n"
            "Todas son centrales hidroeléctricas con gobernadores de turbina que participan en la RPF del SIN."
        ),
    ]


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Genera dataset Q&A para fine-tuning')
    parser.add_argument('--min-pairs', type=int, default=300)
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_file = OUTPUT_DIR / 'qa_dataset.jsonl'

    conn = get_conn()
    all_pairs = []
    try:
        log.info("Generando pares Q&A desde PostgreSQL...")
        generators = [
            ('Resumen general',   gen_resumen_general),
            ('Por unidad',        gen_por_unidad),
            ('Por evento',        gen_por_evento),
            ('Comparativas',      gen_comparativas),
        ]
        for name, gen_fn in generators:
            if name == 'Conceptos técnicos':
                pairs = gen_fn()
            else:
                pairs = gen_fn(conn)
            log.info(f"  {name}: {len(pairs)} pares")
            all_pairs.extend(pairs)
    finally:
        conn.close()

    # Agregar conceptos técnicos estáticos
    tech = gen_conceptos_tecnicos()
    log.info(f"  Conceptos técnicos: {len(tech)} pares")
    all_pairs.extend(tech)

    # Mezclar para entrenamiento más robusto
    random.shuffle(all_pairs)

    # Guardar JSONL
    with open(output_file, 'w', encoding='utf-8') as f:
        for pair in all_pairs:
            f.write(json.dumps(pair, ensure_ascii=False) + '\n')

    # Estadísticas
    stats = {
        'total_pairs':   len(all_pairs),
        'generated_at':  datetime.now().isoformat(),
        'output_file':   str(output_file),
        'ready_for_finetuning': len(all_pairs) >= args.min_pairs,
    }
    stats_file = OUTPUT_DIR / 'qa_dataset_stats.json'
    stats_file.write_text(json.dumps(stats, indent=2, ensure_ascii=False))

    log.info(f"\nDataset generado: {len(all_pairs)} pares → {output_file}")
    log.info(f"{'✓ Suficiente para fine-tuning' if stats['ready_for_finetuning'] else '⚠ Pocos pares — considera agregar más fuentes'}")


if __name__ == '__main__':
    main()
