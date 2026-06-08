#!/usr/bin/env python3
"""
generate_obsidian_vault.py — Genera vault Obsidian para COBEE-AI

Crea una estructura de notas interconectadas con [[wikilinks]] a partir de
los datos de PostgreSQL. El vault puede editarse en Obsidian y sincronizarse
a Open WebUI como Knowledge Base enriquecida.

Estructura generada:
  vault/
  ├── Unidades/       ← una nota por unidad generadora (historial + KPIs)
  ├── Eventos/        ← una nota por evento (todas las unidades + contexto)
  ├── Conceptos/      ← droop, nadir, reserva girante, etc.
  ├── Normativa/      ← CDM, CNDC, metodología RPF
  └── Diagnóstico/   ← patrones de incumplimiento y corrección

Uso:
  python3 generate_obsidian_vault.py
  python3 generate_obsidian_vault.py --vault-dir /ruta/al/vault
"""

import os
import json
import logging
import argparse
from pathlib import Path
from datetime import datetime

import psycopg2
import psycopg2.extras
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
log = logging.getLogger(__name__)

DEFAULT_VAULT = Path(__file__).parent / 'obsidian_vault'


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


def write_note(path: Path, content: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding='utf-8')


# ─── Generadores de notas ─────────────────────────────────────────────────────

def gen_unidades(conn, vault: Path):
    """Una nota por unidad con historial completo de eventos."""
    rows = query(conn, """
        SELECT
            unidad,
            COUNT(DISTINCT semestre || evento)                          AS n_eventos,
            ROUND(AVG(p_max_mw)::numeric, 2)                           AS p_max_avg,
            ROUND(AVG(r_inicial_mw)::numeric, 2)                       AS reserva_avg,
            ROUND(AVG(droop_inf_pct)::numeric, 2)                      AS droop_inf_avg,
            ROUND(AVG(droop_calc_pct)::numeric, 2)                     AS droop_calc_avg,
            ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'Sí' THEN 1 ELSE 0 END)
                / NULLIF(COUNT(*), 0), 1)                              AS pct_cumple,
            MIN(fecha_evento)                                           AS primer_evento,
            MAX(fecha_evento)                                           AS ultimo_evento
        FROM rpf_kpi_cobee
        GROUP BY unidad
        ORDER BY unidad
    """)

    for r in rows:
        u = r['unidad']
        # Historial por evento
        detalle = query(conn, """
            SELECT semestre, evento, fecha_evento,
                   p_max_mw, r_inicial_mw,
                   droop_inf_pct, droop_calc_pct,
                   f_min_hz, aporta_rpf, clasificacion,
                   observaciones
            FROM rpf_kpi_cobee
            WHERE unidad = %s
            ORDER BY semestre, evento::int
        """, (u,))

        # Calcular desviación droop promedio
        desviacion = ""
        try:
            dc = float(r['droop_calc_avg'] or 0)
            di = float(r['droop_inf_avg'] or 0)
            diff = abs(dc - di)
            if diff > 1.5:
                desviacion = f"\n> ⚠️ **Alerta**: Desviación droop promedio {diff:.1f}% — revisar configuración del regulador\n"
            elif diff > 0.5:
                desviacion = f"\n> 📊 Desviación droop promedio {diff:.1f}% — dentro de rango aceptable\n"
        except Exception:
            pass

        # Estado semáforo
        pct = float(r['pct_cumple'] or 0)
        estado = "🟢 Bueno" if pct >= 80 else ("🟡 Regular" if pct >= 50 else "🔴 Crítico")

        content = f"""---
tipo: unidad
unidad: {u}
cumplimiento_pct: {r['pct_cumple']}
droop_inf_avg: {r['droop_inf_avg']}
droop_calc_avg: {r['droop_calc_avg']}
p_max_avg_mw: {r['p_max_avg']}
n_eventos: {r['n_eventos']}
actualizado: {datetime.now().strftime('%Y-%m-%d')}
---

# Unidad {u}

**Estado general**: {estado} ({r['pct_cumple']}% cumplimiento RPF)
**Eventos analizados**: {r['n_eventos']} | **Período**: {r['primer_evento']} → {r['ultimo_evento']}
**Potencia máxima promedio**: {r['p_max_avg']} MW | **Reserva promedio**: {r['reserva_avg']} MW
**Droop informado promedio**: {r['droop_inf_avg']}% | **Droop calculado promedio**: {r['droop_calc_avg']}%
{desviacion}
## Historial de eventos

| Semestre | Evento | Fecha | P_max (MW) | Droop inf | Droop calc | f_min (Hz) | RPF | Clasificación |
|----------|--------|-------|-----------|-----------|-----------|-----------|-----|---------------|
"""
        for d in detalle:
            rpf_icon = "✅" if d['aporta_rpf'] == 'Sí' else ("❌" if d['aporta_rpf'] == 'No' else "—")
            content += (f"| {d['semestre']} | [[Evento {d['evento']} {d['semestre']}|E{d['evento']}]] "
                        f"| {d['fecha_evento'] or '—'} "
                        f"| {d['p_max_mw'] or '—'} "
                        f"| {d['droop_inf_pct'] or '—'}% "
                        f"| {d['droop_calc_pct'] or '—'}% "
                        f"| {d['f_min_hz'] or '—'} "
                        f"| {rpf_icon} "
                        f"| {d['clasificacion'] or '—'} |\n")

        content += f"""
## Diagnóstico

### Droop
- Droop informado promedio: **{r['droop_inf_avg']}%** (debe estar entre [[droop|4% y 6%]])
- Droop calculado promedio: **{r['droop_calc_avg']}%**
- Ver: [[Diagnóstico de droop desviado]]

### Cumplimiento RPF
- Tasa de cumplimiento: **{r['pct_cumple']}%**
- Ver: [[Patrones de incumplimiento RPF]]

## Referencias
- [[Normativa CDM RPF]] — rangos permitidos de droop
- [[Metodología CNDC de evaluación RPF]]
- [[Reserva girante]]
"""
        write_note(vault / 'Unidades' / f'{u}.md', content)

    log.info(f"  Unidades: {len(rows)} notas generadas")
    return [r['unidad'] for r in rows]


def gen_eventos(conn, vault: Path):
    """Una nota por evento con resumen de todas las unidades participantes."""
    eventos = query(conn, """
        SELECT DISTINCT semestre, evento,
               MIN(fecha_evento) AS fecha,
               MIN(f_min_hz) AS f_min,
               COUNT(DISTINCT unidad) AS n_unidades,
               SUM(CASE WHEN aporta_rpf = 'Sí' THEN 1 ELSE 0 END) AS aportan,
               SUM(CASE WHEN aporta_rpf = 'No' THEN 1 ELSE 0 END) AS no_aportan
        FROM rpf_kpi_cobee
        GROUP BY semestre, evento
        ORDER BY semestre, evento::int
    """)

    for ev in eventos:
        sem = ev['semestre']
        num = ev['evento']
        titulo = f"Evento {num} {sem}"

        unidades = query(conn, """
            SELECT unidad, p_max_mw, r_inicial_mw,
                   droop_inf_pct, droop_calc_pct,
                   f_min_hz, aporta_rpf, clasificacion, observaciones
            FROM rpf_kpi_cobee
            WHERE semestre = %s AND evento = %s
            ORDER BY unidad
        """, (sem, num))

        # Severidad del evento
        f_min = float(ev['f_min'] or 50)
        if f_min < 49.0:
            severidad = "🔴 Crítico"
        elif f_min < 49.5:
            severidad = "🟡 Moderado"
        else:
            severidad = "🟢 Leve"

        pct_cumple = round(100 * ev['aportan'] / ev['n_unidades'], 1) if ev['n_unidades'] else 0

        content = f"""---
tipo: evento
semestre: {sem}
evento: {num}
fecha: {ev['fecha'] or 'desconocida'}
f_min_hz: {ev['f_min']}
severidad: {severidad}
n_unidades: {ev['n_unidades']}
pct_cumplimiento: {pct_cumple}
actualizado: {datetime.now().strftime('%Y-%m-%d')}
---

# {titulo}

**Severidad**: {severidad} | **Fecha**: {ev['fecha'] or 'desconocida'}
**Frecuencia mínima ([[nadir]])**: **{ev['f_min']} Hz**
**Cumplimiento COBEE**: {ev['aportan']}/{ev['n_unidades']} unidades ({pct_cumple}%)

## Resultados por unidad

| Unidad | P_max (MW) | Reserva (MW) | Droop inf | Droop calc | f_min (Hz) | RPF | Clasificación |
|--------|-----------|-------------|-----------|-----------|-----------|-----|---------------|
"""
        for u in unidades:
            rpf_icon = "✅" if u['aporta_rpf'] == 'Sí' else ("❌" if u['aporta_rpf'] == 'No' else "—")
            content += (f"| [[{u['unidad']}]] "
                        f"| {u['p_max_mw'] or '—'} "
                        f"| {u['r_inicial_mw'] or '—'} "
                        f"| {u['droop_inf_pct'] or '—'}% "
                        f"| {u['droop_calc_pct'] or '—'}% "
                        f"| {u['f_min_hz'] or '—'} "
                        f"| {rpf_icon} "
                        f"| {u['clasificacion'] or '—'} |\n")

        no_aportaron = [u['unidad'] for u in unidades if u['aporta_rpf'] == 'No']
        aportaron    = [u['unidad'] for u in unidades if u['aporta_rpf'] == 'Sí']

        content += f"""
## Análisis

### Unidades que aportaron RPF
{', '.join(f'[[{u}]]' for u in aportaron) or '— ninguna —'}

### Unidades que NO aportaron RPF
{', '.join(f'[[{u}]]' for u in no_aportaron) or '— todas aportaron ✅ —'}

### Contexto
- Nadir de frecuencia: **{ev['f_min']} Hz** (nominal 50 Hz, mínimo aceptable ~49.0 Hz)
- Ver [[Normativa CDM RPF]] para criterios de evaluación
- Ver [[Metodología CNDC de evaluación RPF]] para cálculo de cargos

## Referencias
- Semestre: [[{sem}]]
"""
        write_note(vault / 'Eventos' / f'{titulo}.md', content)

    log.info(f"  Eventos: {len(eventos)} notas generadas")


def gen_conceptos(vault: Path):
    """Notas de conceptos técnicos con links entre sí."""
    conceptos = {
        'droop': """---
tipo: concepto
tags: [regulacion, gobernador, parametro]
---

# Droop (Estatismo del Regulador)

El **droop** es el parámetro del regulador de velocidad que define la respuesta de potencia activa ante desviaciones de frecuencia.

## Fórmula

$$droop = \\frac{\\Delta f / f_{nominal}}{\\Delta P / P_{nominal}} \\times 100\\%$$

## Valores normativos en Bolivia ([[Normativa CDM RPF]])

| Parámetro | Valor |
|-----------|-------|
| Droop mínimo | 4% |
| Droop máximo | 6% |
| Fuera de rango | Incumplimiento RPF |

## Tipos
- **Droop informado**: valor declarado al [[Metodología CNDC de evaluación RPF|CNDC]]
- **Droop calculado**: valor obtenido de mediciones SCADA durante el evento

## Desviación entre droop informado y calculado
Una diferencia > 1% indica posible problema de configuración del regulador.
Ver [[Diagnóstico de droop desviado]].

## Unidades COBEE con droop histórico
Ver notas en [[Unidades/]] — cada unidad tiene su historial de droop por evento.
""",
        'nadir': """---
tipo: concepto
tags: [frecuencia, evento, minimo]
---

# Nadir de Frecuencia

El **nadir** es el valor mínimo de frecuencia alcanzado durante un evento RPF, antes de que la regulación primaria detenga la caída.

## Valores de referencia (SIN Bolivia)

| Rango | Clasificación |
|-------|--------------|
| > 49.5 Hz | 🟢 Leve |
| 49.0 – 49.5 Hz | 🟡 Moderado |
| < 49.0 Hz | 🔴 Crítico |

## Relación con [[droop]]
Un droop correcto (4-6%) maximiza la respuesta de potencia y eleva el nadir.
Un droop incorrecto o regulador deshabilitado resulta en nadir más bajo.

## Ver también
- [[Reserva girante]] — potencia disponible para responder ante el evento
- [[Normativa CDM RPF]] — consecuencias económicas según severidad
""",
        'reserva-girante': """---
tipo: concepto
tags: [potencia, regulacion, operacion]
---

# Reserva Girante

La **reserva girante** (o reserva primaria) es la potencia activa disponible en una unidad generadora sincronizada para responder automáticamente ante una caída de frecuencia.

## Cálculo
$$R_{inicial} = P_{nominal} - P_{actual}$$

Se mide al momento del evento (antes del [[nadir]]).

## Importancia para RPF
- Una reserva mayor permite mayor aporte de potencia durante el evento
- Unidades con reserva = 0 MW (al límite de capacidad) no pueden aportar RPF
- Ver [[Normativa CDM RPF]] para requisitos mínimos

## En COBEE
Las unidades hidroeléctricas (cascada Zongo) tienen reserva variable según el caudal disponible y el despacho del [[Metodología CNDC de evaluación RPF|CNDC]].
""",
    }

    for nombre, contenido in conceptos.items():
        write_note(vault / 'Conceptos' / f'{nombre}.md', contenido)

    log.info(f"  Conceptos: {len(conceptos)} notas generadas")


def gen_normativa(vault: Path):
    """Notas de normativa y metodología."""
    notas = {
        'Normativa CDM RPF': """---
tipo: normativa
tags: [CDM, CNDC, regulacion, Bolivia]
---

# Normativa CDM — Regulación Primaria de Frecuencia

El **CDM** (Contrato de Abastecimiento con Despacho de Mínimo Costo) es el reglamento técnico-económico del mercado eléctrico boliviano administrado por el CNDC.

## Requisitos RPF para generadores

| Parámetro | Requisito |
|-----------|-----------|
| [[droop\|Droop]] permitido | **4% – 6%** |
| Tiempo de respuesta | Primeros **30 segundos** |
| Habilitación del regulador | **Obligatoria** durante operación |
| Evaluación | **Semestral** por el CNDC |

## Consecuencias del incumplimiento
1. Cargo económico proporcional a la energía no aportada
2. Requerimiento de prueba de regulador
3. Posible restricción de despacho

## Metodología de evaluación
Ver [[Metodología CNDC de evaluación RPF]]

## Unidades COBEE evaluadas
Ver [[Unidades/]] — historial de cumplimiento por semestre.
""",
        'Metodología CNDC de evaluación RPF': """---
tipo: normativa
tags: [CNDC, metodologia, evaluacion]
---

# Metodología CNDC de Evaluación RPF

## Proceso de evaluación semestral

1. **Selección de eventos**: el CNDC selecciona eventos con Δf significativa
2. **Extracción de señales SCADA**: frecuencia del sistema + potencia de cada unidad
3. **Cálculo de [[droop]] real**: usando Δf y ΔP medidos
4. **Comparación con droop informado**: diferencia > umbral → incumplimiento
5. **Informe semestral**: publicado a todos los agentes del mercado

## Criterios de incumplimiento
- [[droop\|Droop]] calculado fuera del rango 4-6%
- Potencia no varía durante el evento (regulador deshabilitado)
- Variación de potencia en sentido contrario al requerido

## Cálculo del cargo por incumplimiento
$$Cargo = E_{no\_aportada} \\times CMg_{marginal}$$

Donde $E_{no\_aportada}$ es la energía que debió haber aportado la unidad según su [[droop]] declarado.

## Referencias
- [[Normativa CDM RPF]]
- [[droop]]
- [[Diagnóstico de droop desviado]]
""",
    }

    for nombre, contenido in notas.items():
        write_note(vault / 'Normativa' / f'{nombre}.md', contenido)

    log.info(f"  Normativa: {len(notas)} notas generadas")


def gen_diagnostico(conn, vault: Path):
    """Notas de diagnóstico con patrones identificados desde los datos."""
    # Unidades con mayor desviación droop
    problemas = query(conn, """
        SELECT unidad,
               ROUND(AVG(ABS(droop_calc_pct - droop_inf_pct))::numeric, 2) AS desv_avg,
               ROUND(100.0 * SUM(CASE WHEN aporta_rpf = 'No' THEN 1 ELSE 0 END)
                   / NULLIF(COUNT(*), 0), 1) AS pct_incumple,
               COUNT(*) AS n_registros
        FROM rpf_kpi_cobee
        WHERE droop_calc_pct IS NOT NULL AND droop_inf_pct IS NOT NULL
        GROUP BY unidad
        HAVING AVG(ABS(droop_calc_pct - droop_inf_pct)) > 0.5
        ORDER BY desv_avg DESC
        LIMIT 10
    """)

    tabla = "| Unidad | Desv. droop promedio | % Incumplimiento |\n|--------|---------------------|------------------|\n"
    for p in problemas:
        tabla += f"| [[{p['unidad']}]] | {p['desv_avg']}% | {p['pct_incumple']}% |\n"

    content = f"""---
tipo: diagnostico
tags: [droop, incumplimiento, alerta]
actualizado: {datetime.now().strftime('%Y-%m-%d')}
---

# Diagnóstico de Droop Desviado

Unidades con mayor desviación histórica entre [[droop]] informado y calculado.

## Unidades con mayor desviación (datos históricos)

{tabla}

## Causas típicas

1. **Regulador mal configurado**: el parámetro R del gobernador no coincide con lo declarado al CNDC
2. **Saturación del regulador**: límites de potencia activos durante el evento
3. **Modo manual**: operador deshabilitó el regulador automático
4. **Falla mecánica**: desgaste en válvulas o servo del gobernador hidráulico

## Proceso de diagnóstico en PowerFactory

1. Abrir modelo de la unidad sospechosa (ElmSym)
2. Verificar parámetro **R** (estatismo) en el gobernador (ElmGovm)
3. Comparar con valor declarado al CNDC
4. Ejecutar simulación RMS del evento y comparar curva de potencia con SCADA
5. Si divergen: ajustar R hasta reproducir el comportamiento medido

## Protocolo de corrección

```
1. Identificar evento con mayor desviación
2. Extraer señales SCADA (frecuencia + potencia) del evento
3. Calcular droop real = Δf/Δp normalizado
4. Actualizar parámetro en el modelo PowerFactory
5. Notificar al CNDC si el droop real queda fuera del rango 4-6%
```

## Referencias
- [[droop]] — definición y fórmula
- [[Normativa CDM RPF]] — consecuencias del incumplimiento
- [[Metodología CNDC de evaluación RPF]]
"""
    write_note(vault / 'Diagnóstico' / 'Diagnóstico de droop desviado.md', content)

    # Nota de patrones generales
    content2 = f"""---
tipo: diagnostico
tags: [patrones, incumplimiento, rpf]
actualizado: {datetime.now().strftime('%Y-%m-%d')}
---

# Patrones de Incumplimiento RPF

Patrones identificados en el historial de eventos de COBEE.

## Patrón 1 — Regulador deshabilitado
**Señal**: potencia activa constante durante el evento (curva plana en SCADA)
**Causa**: operador en modo manual o falla del regulador automático
**Acción**: verificar flag de habilitación del gobernador + prueba de respuesta

## Patrón 2 — Droop fuera de rango
**Señal**: droop calculado < 4% o > 6%
**Causa**: configuración incorrecta del parámetro R en el gobernador
**Acción**: ver [[Diagnóstico de droop desviado]]

## Patrón 3 — Reserva insuficiente
**Señal**: la unidad aporta pero mucho menos de lo esperado
**Causa**: unidad operando cerca del límite de capacidad ([[reserva-girante|reserva girante]] ≈ 0)
**Acción**: coordinar con despacho para mantener margen de reserva mínimo

## Patrón 4 — Respuesta tardía
**Señal**: la potencia varía pero con retraso > 10 segundos tras el evento
**Causa**: constante de tiempo del gobernador demasiado alta
**Acción**: revisar parámetro Tw (tiempo de agua) y Td (tiempo derivativo) del gobernador

## Estadísticas por patrón (datos históricos COBEE)
Ver tablas en cada nota de [[Unidades/]] y [[Eventos/]].
"""
    write_note(vault / 'Diagnóstico' / 'Patrones de incumplimiento RPF.md', content2)

    log.info("  Diagnóstico: 2 notas generadas")


def gen_indice(vault: Path, unidades: list):
    """Nota índice principal del vault."""
    content = f"""---
tipo: indice
actualizado: {datetime.now().strftime('%Y-%m-%d')}
---

# COBEE-AI — Base de Conocimiento RPF

Vault de conocimiento para análisis de Regulación Primaria de Frecuencia (RPF)
del Sistema Interconectado Nacional de Bolivia — COBEE S.A.

## Navegación rápida

### Por unidad generadora
{chr(10).join(f'- [[{u}]]' for u in unidades)}

### Por tema
- [[droop]] — estatismo del regulador de velocidad
- [[nadir]] — frecuencia mínima durante un evento
- [[reserva-girante]] — potencia disponible para RPF
- [[Normativa CDM RPF]] — requisitos y consecuencias
- [[Metodología CNDC de evaluación RPF]] — cómo evalúa el CNDC
- [[Diagnóstico de droop desviado]] — unidades con problemas
- [[Patrones de incumplimiento RPF]] — causas comunes

### Por carpeta
- [[Unidades/]] — historial KPI por generador
- [[Eventos/]] — resumen por evento semestral
- [[Conceptos/]] — glosario técnico
- [[Normativa/]] — CDM y metodología CNDC
- [[Diagnóstico/]] — patrones y correcciones

## Uso con COBEE-AI
Este vault se sincroniza automáticamente con Open WebUI.
Para actualizar la Knowledge Base ejecutar en el servidor:
```bash
python3 sync_vault_to_kb.py
```
"""
    write_note(vault / 'INDICE.md', content)


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Genera vault Obsidian para COBEE-AI')
    parser.add_argument('--vault-dir', default=str(DEFAULT_VAULT),
                        help=f'Directorio del vault (default: {DEFAULT_VAULT})')
    args = parser.parse_args()

    vault = Path(args.vault_dir)
    vault.mkdir(parents=True, exist_ok=True)
    log.info(f"Generando vault en: {vault}")

    conn = get_conn()
    try:
        log.info("Generando notas...")
        unidades = gen_unidades(conn, vault)
        gen_eventos(conn, vault)
        gen_conceptos(vault)
        gen_normativa(vault)
        gen_diagnostico(conn, vault)
        gen_indice(vault, unidades)
    finally:
        conn.close()

    # Contar notas generadas
    total = sum(1 for _ in vault.rglob('*.md'))
    log.info(f"\n✓ Vault generado: {total} notas en {vault}")
    log.info(f"\nPara sincronizar con Open WebUI:")
    log.info(f"  python3 sync_vault_to_kb.py --vault-dir {vault}")


if __name__ == '__main__':
    main()
