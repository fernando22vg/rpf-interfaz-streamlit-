#!/usr/bin/env python3
"""
expand_dataset.py — Amplía el dataset Q&A leyendo los markdown del corpus CNDC

Lee los archivos .md generados por build_full_corpus.py y extrae pares Q&A
adicionales para enriquecer el dataset de fine-tuning.

Salida: agrega pares al archivo training/qa_dataset.jsonl existente
        (o crea uno nuevo si no existe)

Uso:
  python3 expand_dataset.py
  python3 expand_dataset.py --corpus-dir /ruta/al/corpus
"""

import os
import re
import json
import random
import logging
import argparse
from pathlib import Path

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

OUTPUT_DIR   = Path(__file__).parent / 'training'
CORPUS_DIR   = Path(__file__).parent / 'corpus_cache'
DATASET_FILE = OUTPUT_DIR / 'qa_dataset.jsonl'

SYSTEM_PROMPT = """Eres COBEE-AI, asistente técnico experto en Regulación Primaria de Frecuencia (RPF) del Sistema Interconectado Nacional (SIN) de Bolivia, especializado en el análisis de las unidades generadoras de COBEE."""


def qa_pair(q: str, a: str) -> dict:
    return {"messages": [
        {"role": "system",    "content": SYSTEM_PROMPT},
        {"role": "user",      "content": q},
        {"role": "assistant", "content": a},
    ]}


# ─── Parser de archivos markdown del corpus ───────────────────────────────────

def parse_resultado_md(path: Path) -> dict | None:
    """Extrae campos clave de un archivo resultado_*_UNIDAD.md"""
    try:
        txt = path.read_text(encoding='utf-8')
    except Exception:
        return None

    def extract(label, text):
        m = re.search(rf'{re.escape(label)}[:\s|]*([^\n|]+)', text, re.IGNORECASE)
        return m.group(1).strip() if m else None

    return {
        'file': path.name,
        'evento': extract('Evento', txt),
        'semestre': extract('Semestre', txt),
        'unidad': extract('Unidad', txt),
        'f_min': extract('f_min', txt) or extract('Frecuencia mínima', txt),
        'droop_inf': extract('Droop inf', txt) or extract('droop_inf', txt),
        'droop_calc': extract('Droop calc', txt) or extract('droop_calc', txt),
        'aporta_rpf': extract('Aporta RPF', txt) or extract('aporta_rpf', txt),
        'reserva': extract('Reserva', txt) or extract('r_inicial', txt),
        'clasificacion': extract('Clasificación', txt) or extract('clasificacion', txt),
        'raw': txt,
    }


def parse_tabla_resultados(path: Path) -> list[dict]:
    """Extrae filas de tabla_resultados_COBEE.md"""
    try:
        txt = path.read_text(encoding='utf-8')
    except Exception:
        return []

    rows = []
    for line in txt.splitlines():
        if '|' not in line or line.strip().startswith('|--'):
            continue
        cols = [c.strip() for c in line.strip('|').split('|')]
        if len(cols) >= 4 and cols[0] not in ('Unidad', 'unidad', ''):
            rows.append(cols)
    return rows


# ─── Generadores de pares Q&A desde markdown ──────────────────────────────────

def gen_from_resultado(r: dict) -> list[dict]:
    """Genera pares Q&A desde un archivo de resultado de unidad."""
    pairs = []
    if not r or not r.get('unidad'):
        return pairs

    u = r['unidad']
    ev = r.get('evento', '?')
    sem = r.get('semestre', '?')

    # Par 1 — desempeño general de la unidad en el evento
    if r.get('aporta_rpf') and r.get('f_min'):
        ans = (
            f"En el **Evento {ev}** del semestre {sem}, la unidad **{u}** "
            f"{'aportó' if 'Sí' in str(r['aporta_rpf']) else 'NO aportó'} RPF. "
        )
        if r.get('f_min'):
            ans += f"La frecuencia mínima registrada fue **{r['f_min']} Hz**. "
        if r.get('droop_calc'):
            ans += f"El droop calculado fue **{r['droop_calc']}%**"
            if r.get('droop_inf'):
                ans += f" (informado: {r['droop_inf']}%)."
        pairs.append(qa_pair(
            f"¿Cómo fue el desempeño de la unidad {u} en el Evento {ev}?", ans))
        pairs.append(qa_pair(
            f"¿La unidad {u} aportó RPF en el evento {ev} del semestre {sem}?",
            f"{'Sí' if 'Sí' in str(r['aporta_rpf']) else 'No'}, la unidad {u} "
            f"{'aportó' if 'Sí' in str(r['aporta_rpf']) else 'no aportó'} RPF. " + ans))

    # Par 2 — droop
    if r.get('droop_calc') and r.get('droop_inf'):
        diff_text = ""
        try:
            dc = float(str(r['droop_calc']).replace(',', '.').replace('%', ''))
            di = float(str(r['droop_inf']).replace(',', '.').replace('%', ''))
            diff = abs(dc - di)
            if diff > 1.0:
                diff_text = (f" Existe una **desviación de {diff:.1f}%** entre el droop "
                             f"calculado y el informado, lo que puede indicar un problema "
                             f"de configuración del regulador.")
        except Exception:
            pass
        pairs.append(qa_pair(
            f"¿Cuál es el droop de la unidad {u} en el evento {ev}?",
            f"La unidad **{u}** tiene droop informado de **{r['droop_inf']}%** "
            f"y droop calculado de **{r['droop_calc']}%** en el Evento {ev}.{diff_text}"))

    # Par 3 — reserva
    if r.get('reserva'):
        pairs.append(qa_pair(
            f"¿Cuál fue la reserva inicial de {u} en el evento {ev}?",
            f"La unidad **{u}** tenía una reserva inicial de **{r['reserva']} MW** "
            f"disponible para RPF en el Evento {ev} del semestre {sem}."))

    # Par 4 — clasificación
    if r.get('clasificacion'):
        pairs.append(qa_pair(
            f"¿Cuál es la clasificación de {u} en el evento {ev}?",
            f"La unidad **{u}** fue clasificada como **{r['clasificacion']}** "
            f"en el Evento {ev} del semestre {sem}."))

    return pairs


def gen_from_tabla(path: Path, rows: list) -> list[dict]:
    """Genera pares Q&A desde tabla_resultados de un evento."""
    pairs = []
    # Extraer evento y semestre del nombre del archivo
    m = re.search(r'resultado_(\d{4}_sem\d+)_Evento_(\d+)', path.name)
    if not m:
        return pairs
    sem, ev = m.group(1).replace('_', ' '), m.group(2)

    if not rows:
        return pairs

    # Par — resumen del evento
    aportan = [r[0] for r in rows if len(r) > 2 and 'Sí' in str(r[2])]
    no_aportan = [r[0] for r in rows if len(r) > 2 and 'No' in str(r[2])]

    if aportan or no_aportan:
        ans = f"En el **Evento {ev}** ({sem}):\n"
        if aportan:
            ans += f"- ✅ Aportaron RPF: **{', '.join(aportan)}**\n"
        if no_aportan:
            ans += f"- ❌ No aportaron RPF: **{', '.join(no_aportan)}**\n"
        ans += f"\nTotal: {len(aportan)} de {len(rows)} unidades cumplieron."

        pairs.append(qa_pair(
            f"¿Qué unidades de COBEE aportaron RPF en el Evento {ev}?", ans))
        pairs.append(qa_pair(
            f"Dame el resumen del Evento {ev} para COBEE", ans))

        if no_aportan:
            pairs.append(qa_pair(
                f"¿Qué unidades NO aportaron RPF en el Evento {ev}?",
                f"En el Evento {ev} ({sem}), las unidades que **no aportaron RPF** fueron: "
                f"**{', '.join(no_aportan)}**. Esto puede deberse a droop fuera de norma, "
                f"regulador deshabilitado o unidad fuera de servicio."))

    return pairs


def gen_conceptos_tecnicos() -> list[dict]:
    """Pares Q&A sobre conceptos técnicos RPF — no dependen de BD."""
    return [
        qa_pair(
            "¿Qué es el droop en una unidad generadora?",
            "El **droop** (o estatismo) es el parámetro del regulador de velocidad que define "
            "cuánta potencia activa entrega una unidad por cada desviación de frecuencia. "
            "Se expresa en porcentaje: un droop de **4%** significa que la unidad varía su "
            "potencia en un 100% de su capacidad cuando la frecuencia se desvía 4% de 50 Hz. "
            "En Bolivia, el CDM exige un droop entre **4% y 6%** para participar en RPF."
        ),
        qa_pair(
            "¿Qué es la Regulación Primaria de Frecuencia (RPF)?",
            "La **RPF** es la respuesta automática e inmediata de las unidades generadoras "
            "ante desviaciones de frecuencia en el sistema eléctrico. Ocurre en los primeros "
            "**30 segundos** tras un evento (pérdida de generación o carga). "
            "Las unidades con regulador habilitado y droop correcto aumentan o reducen "
            "automáticamente su potencia para estabilizar la frecuencia. En Bolivia, el "
            "CNDC evalúa el cumplimiento RPF semestral y puede aplicar multas por incumplimiento."
        ),
        qa_pair(
            "¿Cuál es la frecuencia nominal del SIN boliviano?",
            "La frecuencia nominal del Sistema Interconectado Nacional (SIN) de Bolivia es "
            "**50 Hz**. Las tolerancias operativas normales son ±0.2 Hz (49.8 – 50.2 Hz). "
            "Durante un evento RPF, la frecuencia puede caer a valores de nadir entre "
            "49.0 y 49.7 Hz dependiendo de la magnitud del disturbio y la respuesta de las unidades."
        ),
        qa_pair(
            "¿Qué unidades generadoras tiene COBEE en el SIN?",
            "COBEE opera principalmente unidades **hidroeléctricas** en la cascada del río Zongo "
            "y otras plantas. Las unidades analizadas en RPF incluyen: BOT01, BOT02, BOT03 "
            "(Botijlaca), HUA01, HUA02 (Huayna Potosí), ZON (Zongo), CHU01, CHU02 "
            "(Chururaqui), CAH01, CAH02 (Cahua), SAI (Sainani), CUT01, CUT02, CUT03, CUT05 "
            "(Cuticucho), HAR01, HAR02 (Harkapampa), SRO01, SRO02 (Santa Rosa), TIQ (Tiquimani)."
        ),
        qa_pair(
            "¿Qué es el nadir de frecuencia?",
            "El **nadir** es el valor mínimo de frecuencia alcanzado durante un evento RPF, "
            "antes de que la regulación primaria detenga la caída. Un nadir más bajo indica "
            "un evento más severo o una respuesta insuficiente del sistema. "
            "El CNDC registra el nadir para evaluar la gravedad de cada evento y el "
            "desempeño de las unidades. Valores por debajo de **49.0 Hz** son considerados "
            "críticos en el SIN boliviano."
        ),
        qa_pair(
            "¿Cómo se calcula el droop de una unidad en base a mediciones SCADA?",
            "El droop calculado se obtiene de las señales SCADA del evento RPF usando la fórmula:\n\n"
            "```\ndroop_calc = (Δf / f_nominal) / (ΔP / P_nominal) × 100%\n```\n\n"
            "Donde:\n"
            "- **Δf** = variación de frecuencia entre t_inicio y t_nadir\n"
            "- **ΔP** = variación de potencia activa de la unidad en el mismo período\n"
            "- **f_nominal** = 50 Hz\n"
            "- **P_nominal** = potencia máxima de la unidad\n\n"
            "Si el droop calculado difiere significativamente del droop informado al CNDC, "
            "puede indicar un problema de configuración del regulador de velocidad."
        ),
        qa_pair(
            "¿Qué pasa si una unidad no aporta RPF?",
            "Si una unidad generadora **no aporta RPF** cuando debería, las consecuencias son:\n"
            "1. **Técnica**: contribuye a un nadir más bajo y mayor riesgo de desconexión en cadena\n"
            "2. **Regulatoria**: el CNDC puede aplicar **cargos económicos** según el CDM\n"
            "3. **Operativa**: la unidad puede ser evaluada para verificar si el regulador "
            "de velocidad está habilitado y correctamente configurado\n\n"
            "Las causas típicas son: regulador deshabilitado, droop fuera del rango 4-6%, "
            "unidad en modo manual, o limitaciones mecánicas del gobernador."
        ),
        qa_pair(
            "¿Qué es el CDM en Bolivia?",
            "El **CDM** (Contrato de Abastecimiento con Despacho de Mínimo Costo) es el "
            "reglamento técnico y económico del mercado eléctrico boliviano administrado "
            "por el CNDC. Define las obligaciones de las unidades generadoras, incluyendo:\n"
            "- Participación obligatoria en RPF\n"
            "- Rangos de droop permitidos (4-6%)\n"
            "- Metodología de evaluación semestral\n"
            "- Cargos por incumplimiento de servicios complementarios"
        ),
        qa_pair(
            "¿Cómo corregir el droop de una unidad en PowerFactory?",
            "Para corregir el droop en un modelo DIgSILENT PowerFactory:\n"
            "1. Abre el elemento de la unidad generadora (ElmSym)\n"
            "2. Ve al controlador de velocidad asociado (ElmGovm o frame del gobernador)\n"
            "3. Localiza el parámetro **R** (estatismo permanente) o **droop**\n"
            "4. El valor se ingresa en por unidad: droop 5% = R = 0.05\n"
            "5. Verifica que el regulador esté **habilitado** (flag de control activo)\n"
            "6. Ejecuta un Load Flow para confirmar el punto de operación\n\n"
            "Para condiciones iniciales RPF, también verifica que la reserva girante "
            "esté correctamente configurada en los límites del gobernador."
        ),
        qa_pair(
            "¿Qué es una simulación RMS en PowerFactory?",
            "Una simulación **RMS** (Root Mean Square) en DIgSILENT PowerFactory es un "
            "análisis de estabilidad transitoria en el dominio del tiempo que resuelve "
            "las ecuaciones diferenciales del sistema eléctrico con paso de tiempo típico "
            "de 10-50 ms. Se usa para estudiar:\n"
            "- Respuesta de frecuencia ante pérdida de generación (RPF)\n"
            "- Estabilidad de voltaje\n"
            "- Comportamiento de reguladores (gobernadores, AVR)\n\n"
            "Para estudios RPF de COBEE, el evento se simula como una perturbación de "
            "potencia activa y se verifica que el nadir de frecuencia quede dentro de "
            "los límites del CDM."
        ),
    ]


def gen_diagnosticos_patrones() -> list[dict]:
    """Pares Q&A sobre diagnóstico de problemas comunes."""
    return [
        qa_pair(
            "¿Cómo identificar si el regulador de una unidad estaba deshabilitado durante un evento?",
            "Para identificar si el regulador estaba deshabilitado:\n"
            "1. **Señal SCADA**: la potencia activa de la unidad no varía durante los "
            "primeros 30 segundos del evento (curva plana)\n"
            "2. **Droop calculado → ∞**: ΔP ≈ 0 mientras Δf es significativa\n"
            "3. **Indicador RPF = No** en el análisis del CNDC\n"
            "4. **Comparación con evento anterior**: si en eventos previos sí aportó, "
            "sugiere cambio de configuración o falla del regulador\n\n"
            "Acción correctiva: verificar parámetro de habilitación del gobernador "
            "y coordinar con operación para prueba de respuesta en frecuencia."
        ),
        qa_pair(
            "¿Cuáles son los eventos RPF más críticos registrados en COBEE?",
            "Los eventos más críticos se identifican por:\n"
            "- **Nadir más bajo** (frecuencia mínima más alejada de 50 Hz)\n"
            "- **Mayor número de unidades que no aportaron**\n"
            "- **Mayor desviación entre droop informado y calculado**\n\n"
            "Para consultar los eventos críticos específicos, usa la base de datos RPF "
            "con: `SELECT evento, semestre, f_min_hz, COUNT(*) as no_aportan FROM "
            "rpf_kpi_cobee WHERE aporta_rpf='No' GROUP BY evento, semestre, f_min_hz "
            "ORDER BY f_min_hz ASC LIMIT 10`"
        ),
        qa_pair(
            "¿Qué diferencia hay entre droop informado y droop calculado?",
            "- **Droop informado**: valor que la empresa generadora declara al CNDC en "
            "sus registros de coordinación técnica. Es el parámetro configurado en el "
            "regulador de velocidad.\n\n"
            "- **Droop calculado**: valor obtenido de las mediciones reales durante el "
            "evento RPF, usando señales SCADA de potencia y frecuencia.\n\n"
            "**Una diferencia > 1%** entre ambos es señal de alarma e indica que:\n"
            "1. El regulador no está configurado con el valor declarado\n"
            "2. Hay saturación o limitaciones mecánicas\n"
            "3. El modelo de simulación necesita actualización\n\n"
            "En COBEE, esta desviación se monitorea evento a evento para detectar "
            "degradación en los gobernadores hidráulicos."
        ),
    ]


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--corpus-dir', default=str(CORPUS_DIR),
                        help='Directorio con archivos .md del corpus CNDC')
    parser.add_argument('--output', default=str(DATASET_FILE),
                        help='Archivo JSONL de salida')
    parser.add_argument('--min-pairs', type=int, default=500,
                        help='Objetivo mínimo de pares Q&A')
    args = parser.parse_args()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    corpus_dir = Path(args.corpus_dir)
    output_path = Path(args.output)

    # Cargar pares existentes (del create_qa_dataset.py anterior)
    existing = []
    if output_path.exists():
        with open(output_path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        existing.append(json.loads(line))
                    except Exception:
                        pass
    log.info(f"Pares existentes: {len(existing)}")

    new_pairs = []

    # 1. Pares desde archivos markdown del corpus
    if corpus_dir.exists():
        md_files = sorted(corpus_dir.glob('*.md'))
        log.info(f"Archivos markdown encontrados: {len(md_files)}")
        for f in md_files:
            if 'tabla_resultados' in f.name:
                rows = parse_tabla_resultados(f)
                new_pairs.extend(gen_from_tabla(f, rows))
            elif 'resultado_' in f.name:
                r = parse_resultado_md(f)
                if r:
                    new_pairs.extend(gen_from_resultado(r))
    else:
        log.warning(f"Directorio corpus no encontrado: {corpus_dir}")
        log.info("Generando solo pares conceptuales y de diagnóstico")

    # 2. Pares de conceptos técnicos
    conceptos = gen_conceptos_tecnicos()
    new_pairs.extend(conceptos)
    log.info(f"  Conceptos técnicos: {len(conceptos)} pares")

    # 3. Patrones de diagnóstico
    diagnosticos = gen_diagnosticos_patrones()
    new_pairs.extend(diagnosticos)
    log.info(f"  Patrones diagnóstico: {len(diagnosticos)} pares")

    # Combinar y deduplicar
    all_pairs = existing.copy()
    seen_questions = {
        p['messages'][1]['content'] for p in existing
        if len(p.get('messages', [])) > 1
    }
    added = 0
    for p in new_pairs:
        q = p['messages'][1]['content'] if len(p.get('messages', [])) > 1 else ''
        if q and q not in seen_questions:
            all_pairs.append(p)
            seen_questions.add(q)
            added += 1

    # Mezclar aleatoriamente
    random.seed(42)
    random.shuffle(all_pairs)

    # Guardar
    with open(output_path, 'w', encoding='utf-8') as f:
        for p in all_pairs:
            f.write(json.dumps(p, ensure_ascii=False) + '\n')

    log.info(f"\nDataset final: {len(all_pairs)} pares ({added} nuevos agregados)")
    log.info(f"Guardado en: {output_path}")

    if len(all_pairs) < args.min_pairs:
        log.warning(f"⚠ Objetivo {args.min_pairs} no alcanzado — "
                    f"considera agregar documentos al corpus_dir o más fuentes")
    else:
        log.info(f"✓ Objetivo de {args.min_pairs} pares alcanzado")


if __name__ == '__main__':
    main()
