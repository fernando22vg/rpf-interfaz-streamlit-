# CargaCondIniciales_PF.py — Documentación técnica completa

**Script:** `Interfaz_RPF/CargaCondIniciales_PF.py`
**Propósito:** Reproducir en DIgSILENT PowerFactory el estado operativo real del SIN boliviano en el instante previo a un evento de frecuencia, dejando el modelo listo para la simulación RMS del disparo.

---

## 1. Posición en el flujo de trabajo

```
ExtFLujos2daO.py          → extrae despacho/demanda del CNDC (datos_simulacion_*.xlsx)
CondInicialesPF.py        → genera condiciones_iniciales_{fecha}_Ev{N}.xlsx
CargaCondIniciales_PF.py  → ESTE SCRIPT: carga el estado en PF + Load Flow + ComInc
        ↓
PowerFactory (manual)     → ComEvt (disparo) + ComSim (simulación RMS)
        ↓
Comparación curva simulada vs registro SCADA real
```

**Entradas:**
- `condiciones_iniciales_{fecha}_Ev{N}.xlsx` — despacho por unidad (pgini) y demanda por carga (plini) en el instante del evento
- `loc_names_xfo.xlsx` — capacidad de transformadores (restricción de cargas por barra)
- `loc_names_gen.xlsx` — Pmax/Pmin/tecnología por generador
- `postot{YYYYMMDD}.xlsx` o `td_{DDMMYY}.xlsx` (opcional) — retiros nodales oficiales del CNDC

**Salidas:**
- `datos_cargados_Ev{N}.xlsx` — estado cargado pre-Load-Flow
- `datos_cargados_Ev{N}_ajustado.xlsx` — estado post-ajuste (incluye pérdidas reales)
- Escenario de operación guardado en PF, con LF convergido y ComInc ejecutado

---

## 2. Descripción detallada por etapa

### [1] Selección de semestre y evento
Navegación interactiva por `C:\Datos del CNDC\01_INFO CNDC_RPF\{semestre}\Análisis_todos_los_eventos\{Evento N}`. En modo batch (runner) las respuestas se inyectan automáticamente desde un JSON.

### [2] Lectura de condiciones iniciales
Lee las hojas `pgini_GEN` (despacho por generador con su "Fuente": P0_medido, CNDC_proporcional, sin_despacho, mantenimiento, disparo) y `plini_CAR` (demanda por carga) del Excel generado por `CondInicialesPF.py`.

### [2b] Corrección de Pdem_evento con retiros CNDC nodal
Si existe el reporte "MWh y Costo Marginal en Nodos" del CNDC, la demanda objetivo del evento se reemplaza por los **retiros oficiales nodales** interpolados al minuto del evento.

**Justificación:** la demanda estimada en `datos_simulacion` proviene de mediciones de generación menos pérdidas estimadas; los retiros nodales son la medición oficial de consumo del CNDC. Usar el valor oficial elimina un sesgo de hasta ±4% que impedía la convergencia del Load Flow o forzaba al slack a valores no físicos.

### [3]–[4] Conexión a PowerFactory y escenario de operación
Activa el proyecto (`PF_PROYECTO`), el caso de estudio (`CASO_BASE`) y crea/activa un **escenario de operación** (IntScenario) propio del evento. Esto aísla las modificaciones: el caso base nunca se altera.

### [4b] Restauración de conectividad
Cierra acopladores/seccionadores que quedaron abiertos en el caso base y verifica que no haya islas eléctricas. Una isla sin slack hace fallar el LF con códigos de error crípticos; detectarlo antes ahorra horas de diagnóstico.

### [5]–[6] Asignación de pgini a generadores
Para cada generador del Excel:
- **P0_medido:** unidades con telemetría → pgini = potencia real medida (fijas, no se tocan).
- **CNDC_proporcional:** unidades sin telemetría individual → el bloque CNDC se redistribuye proporcional a `Pmax × ngnum` del modelo PF.
- **sin_despacho / mantenimiento:** `outserv = 1` (fuera de servicio).
- **disparo:** la unidad que disparó queda EN SERVICIO generando su potencia pre-evento (el trip se simula después con ComEvt).

**Detalle crítico ElmGenstat (renovables):** PowerFactory almacena `pgini` de los generadores estáticos **por máquina** (total/ngnum), mientras que `ElmSym` lo almacena como total. El script escribe y lee con la conversión `× ngnum` correspondiente. Este fue origen de un bug histórico (generación aparente de 845 MW en lugar de 1041 MW).

### [6_DISPARO] Verificación de potencia desconectada
Compara la potencia de las unidades del disparo contra la `p_desc` registrada en la tabla de eventos del CNDC. Cuatro modos: automático desde Excel, manual por unidad, desde p_desc registrada, o sin disparo. La exactitud de `p_desc` define el ΔP de la perturbación — el parámetro más influyente del nadir.

### [6b]–[6c] Verificación Pmax y balance generación-demanda
Recorta unidades que excedan su Pmax del modelo PF y escala los proporcionales para que:

```
Pgen_total = Pdem_evento + P_hierro
```

respetando Pmin y Pmax de cada unidad (saturación iterativa: las unidades que llegan a límite se congelan y el resto absorbe el resto del delta).

### [7] Selección de máquina slack
Prioridad por planta marginal (GCH → CAR → WAR → ERI...) según la lista `PREFIJOS_MARGINALES`. Reglas:
- Renovables y la lista `EXCLUIR_SLACK` nunca son slack.
- En ciclos combinados, solo unidades de **gas** (loc_name terminado en "1") son aptas; las de **vapor** (terminado en "0") no pueden regular frecuencia de forma independiente.
- El slack conserva su `pgini = P0_medido` — el script NO le asigna el residuo del balance a priori.

**Justificación:** el slack del Load Flow es un artefacto matemático (cierra el balance de potencia), pero en la realidad esa unidad operaba en un punto medido. Elegir la unidad marginal real del despacho CNDC como slack minimiza la distorsión.

### [8]–[9] Asignación de plini a cargas
Las cargas se escalan desde el Excel para igualar `Pdem_evento`, con dos restricciones:
1. **Capacidad de transformadores** (`loc_names_xfo.xlsx` × factor `XFO_PF=0.90`): ninguna carga puede exceder la capacidad del trafo que la alimenta. El déficit se redistribuye entre las demás cargas.
2. **Factores por distribuidor** (si hay retiros nodales): cada distribuidora se escala con su propio factor, preservando la distribución geográfica real del consumo.

### [9d] Rebalance final de generación
Tras conocer la demanda realmente cargada en PF (post-restricciones), se recalcula el objetivo:

```
Pgen_objetivo = Pdem_PF_real + P_hierro
```

y se reajustan los proporcionales. El balance pre-LF queda en `+P_hierro` (~8.3 MW para Ev1 2024-II).

### [9b] Exportación de datos cargados
Genera `datos_cargados_Ev{N}.xlsx` con el estado exacto cargado (148 generadores, 226 cargas en el caso típico) para trazabilidad y comparación con el CNDC.

### [10b] Configuración dinámica de cargas (ZIP + kpf)
Vía los tipos de carga (`TypLod`) configura:
- **Modelo de voltaje:** `MODO_ZIP_P = 0` (potencia constante). Con corriente constante, `P_real = plini × V/Vnom`: si V < 1 pu la demanda efectiva baja respecto al valor CNDC, creando un falso superávit. Potencia constante garantiza que las cargas consuman exactamente su plini.
- **Sensibilidad a frecuencia:** `kpf = 1.0`. La dependencia de la carga con la frecuencia (∂P/∂f) amortigua la caída: con kpf alto la frecuencia cae más lento de lo real. Se redujo de 2.0 a 1.0 para acercar la velocidad de caída al registro.

*Nota PF 2025:* `ElmLod.typ_id` retorna proxies de solo lectura; el script verifica primero si los valores ya son correctos (frecuentemente lo son) y solo reporta diferencia real.

### Control de voltaje (av_mode=0 + Vset)
El CNDC solo reporta MW; la potencia reactiva no se asigna. Sin AVR activo los generadores quedan en modo PQ con Q congelada del caso base → red sin soporte reactivo → shunts al límite → voltajes bajos → LF no converge. El script fuerza `av_mode=0` (control de tensión) en todos los ElmSym con setpoints por nivel de tensión (1.00 pu en 500/230/115 kV, 0.98 en 69 kV).

### [10] Load Flow AC
Ejecuta `ComLdf` con regulación automática de taps y límites de reactivos activos. El LF resuelve las ecuaciones nodales completas: aquí aparecen por primera vez las **pérdidas en el cobre** (I²R en líneas y devanados), que no se pueden conocer a priori porque dependen de los flujos resultantes.

### Diagnóstico post-LF
Compara la potencia real (`m:P`) de cada generador contra su pgini asignado. Resultado esperado: todos los no-slack con |m:P − pgini| ≤ 1 MW y el slack absorbiendo exactamente las pérdidas no pre-asignadas.

### [10-corrección] Ajuste post-LF del slack (AJUSTAR_POST_LF=True)
El núcleo del realismo físico. Tras el LF:

```
delta = P_slack_real_LF − P0_medido_slack     (ej. 60.64 − 21.85 = +38.79 MW)
```

Ciclo iterativo (máx. 15 iteraciones, tolerancia 0.1 MW):
1. Reparte `delta` entre los proporcionales (proporcional a su pgini, respetando Pmax/Pmin con saturación).
2. Re-ejecuta el Load Flow.
3. Lee el nuevo P_slack y recalcula delta.
4. Repite hasta |delta| < 0.1 MW.

**Por qué iterativo:** al subir los proporcionales suben los flujos → suben las pérdidas I²R → el LF necesita aún más generación. Las pérdidas de cobre **emergen del propio Load Flow**; no hay fórmula cerrada. El ciclo converge cuando la generación total cubre demanda + pérdidas reales con el slack exactamente en su punto medido.

Exporta `datos_cargados_Ev{N}_ajustado.xlsx`. El balance final (ej. +50.8 MW en Ev1, +63.6 MW en Ev2) son las **pérdidas reales totales de la red** (hierro + cobre), 4.9–5.2% de la demanda — consistente con las pérdidas típicas del SIN (líneas largas 115/230 kV).

### [11] Inicialización RMS (ComInc)
Ejecuta `ComInc` (cálculo de condiciones iniciales dinámicas). PF inicializa todos los modelos DSL (gobernadores, AVR, PSS) en equilibrio con el LF convergido. Si ComInc reporta derivadas no nulas, los modelos arrancarían fuera de equilibrio y la simulación tendría transitorios espurios en t=0.

### [13] Diagnóstico ROCOF
Estima la tasa inicial de caída de frecuencia con la ecuación de oscilación agregada:

```
ROCOF [Hz/s] = ΔP × fn / (2 × H_total)
H_total = Σ (H_i × Snom_i)   sobre generadores en servicio
fn = 50 Hz (SIN boliviano)
```

Lee la constante de inercia H de cada máquina desde PF (fallback H=5 s). Reporta caída esperada a 5 s/10 s y nadir sin RPF, y advierte si ROCOF < 0.15 Hz/s (inercia probablemente sobreestimada en el modelo).

### [12] Guardado del escenario
Persiste el escenario de operación. Opcionalmente crea el ComEvt de disparo automático (`CONFIGURAR_COMEVT=False` por defecto — se recomienda configurarlo manualmente cuando el evento tiene múltiples desconexiones).

---

## 3. Justificación de ingeniería eléctrica del proceso completo

### 3.1 El problema físico
Una simulación RMS de un evento de frecuencia solo es válida si parte del **mismo punto de operación** que el sistema real:

| Variable de estado | Por qué importa para la dinámica |
|---|---|
| Despacho por unidad (P) | Define el margen de reserva rodante de cada gobernador y qué unidades pueden aportar RPF |
| Demanda por nodo | Define los flujos, las pérdidas y la sensibilidad de la carga a f y V |
| Inercia en línea (ΣH·S) | Define el ROCOF: solo cuentan las máquinas EN SERVICIO en ese instante |
| Tensiones/reactivos | Definen los límites de los AVR y el comportamiento de cargas dependientes de V |
| Unidad disparada y su p_desc | Es el ΔP de la perturbación — entrada directa de la ecuación de oscilación |

Un error de 5% en el despacho cambia qué unidades tienen margen de regulación; un error en las unidades fuera de servicio cambia la inercia total y por tanto el ROCOF.

### 3.2 El balance de potencia y el rol del slack
En régimen permanente:

```
Σ Pgen = Σ Pdem + Pérdidas(hierro) + Pérdidas(cobre)
```

- Las pérdidas de hierro (magnetización de trafos) son ~constantes (≈0.8% Pdem) y se pre-asignan.
- Las pérdidas de cobre (~3–4.5% Pdem) dependen cuadráticamente de los flujos → solo se conocen después del LF.

Sin el ajuste post-LF, el slack absorbe todo el cobre y queda en un punto de operación ficticio (ej. 60.6 MW cuando realmente generaba 21.85 MW). Esto distorsiona la simulación RMS de dos formas: (a) el gobernador del slack parte con un margen de reserva equivocado, y (b) la distribución de flujos —y por tanto las pérdidas dinámicas— no corresponde a la real. El ajuste iterativo elimina ambas distorsiones llevando cada unidad, incluida la slack, a su valor medido.

### 3.3 Modelo de cargas (ZIP + kpf)
La respuesta de frecuencia del sistema tiene tres componentes en serie temporal:
1. **Respuesta inercial** (0–2 s): solo H.
2. **Autorregulación de la carga** (continua): D = kpf, las cargas bajan consumo cuando f baja.
3. **Respuesta primaria de gobernadores** (2–30 s): determina el nadir y la recuperación.

kpf actúa como amortiguamiento D en la ecuación de oscilación: `2H·(df/dt) = ΔP − D·Δf`. Un kpf sobreestimado produce un nadir menos profundo y más lento que el real. El valor 1.0 (%P/%f) es representativo de la carga mixta boliviana.

### 3.4 Control de tensión
El evento de frecuencia es un fenómeno P-f, pero la convergencia del caso y la respuesta de las cargas dependen del plano Q-V. Forzar AVR activo con setpoints realistas reproduce el soporte reactivo que las máquinas realmente prestaban, evita el colapso numérico del LF y deja a los AVR inicializados dentro de sus límites para la RMS.

---

## 4. Resultados esperados de una ejecución correcta

| Indicador | Valor esperado | Ejemplo Ev1 (2024-II) |
|---|---|---|
| Load Flow | Convergido (código 0) | ✓ |
| Balance pre-LF (Pgen − Pdem) | +P_hierro (~8.3 MW) | +8.30 MW |
| Slack post-ajuste | = P0_medido ± 0.1 MW | 21.85 MW ✓ |
| Balance post-ajuste | Pérdidas reales ~4–5.5% Pdem | +50.76 MW (4.9%) |
| No-slack: m:P vs pgini | dif ≤ 1 MW en todas | ✓ |
| Cargas no encontradas | 0 | 0 |
| ComInc | Sin derivadas residuales | ✓ |
| ROCOF estimado | Comparable al registro SCADA inicial | ~0.12 Hz/s |

Que el balance del archivo `_ajustado` crezca con la demanda del evento (50.8 MW @ 1033 MW vs 63.6 MW @ 1220 MW) **es correcto**: las pérdidas de cobre escalan con I² ≈ (Pdem)².

---

## 5. ¿Es esta la manera más cercana de recrear eventos reales del SIN?

**Para el punto de partida (condición inicial): sí.** El método implementa el estado del arte para estudios post-evento:
- Usa mediciones reales del CNDC (P0_medido por unidad, retiros nodales oficiales).
- Resuelve las pérdidas reales con el propio Load Flow en lugar de estimarlas.
- Preserva el punto de operación medido de TODAS las unidades, incluida la slack.
- Reproduce la inercia en línea real (solo unidades en servicio en el instante del evento).

No hay información adicional disponible que permita un punto de partida más fiel: el CNDC no publica reactivos ni tensiones nodales, y el script ya usa todo lo que sí se publica.

**Para la trayectoria dinámica (la curva f(t)): el punto inicial es condición necesaria pero NO suficiente.** Aquí está la causa del problema que observas.

### 5.1 Por qué la simulación no alcanza el nadir en el mismo tiempo que el registro

Los datos del Evento 1: el registro SCADA muestra el nadir (49.68 Hz) ~9 s después del disparo; la simulación lo alcanza a ~19–21 s. La magnitud del nadir coincide, el **tiempo no**.

La física que gobierna el instante del nadir es:

```
nadir ocurre cuando: ΔP_gobernadores(t) + D·Δf(t) = ΔP_perturbación
```

Es decir, el nadir llega cuando la potencia inyectada por la respuesta primaria iguala al déficit. Si la simulación tarda 2× más, los gobernadores del modelo entregan potencia 2× más lento que los reales. Las condiciones iniciales **no pueden causar este desfase**: con el mismo ΔP y la misma H, el ROCOF inicial coincide (y coincide: ~0.12 Hz/s en ambos); lo que difiere es la respuesta de los reguladores.

Causas concretas, en orden de probabilidad:

1. **Constantes de tiempo de los gobernadores sobredimensionadas.** El SCADA muestra que WAR41 (turbina de gas, droop real 4.2%) aportaba potencia ya entre t=3–7 s. Una turbina de gas real responde en 2–5 s; si su modelo DSL en PF usa constantes de turbina/válvula de 10–15 s, el nadir se retrasa exactamente como se observa.

2. **Bandas muertas (deadband) distintas a las reales.** La tabla SCADA del evento muestra que la MAYORÍA de las unidades NO aportó RPF (ej. MIS03: −0.41 MW). Si en el modelo PF muchas unidades sí responden (deadband pequeña), la respuesta agregada tiene una forma temporal diferente: más unidades lentas respondiendo en lugar de pocas rápidas.

3. **Unidades sin modelo de gobernador activo** (o en modo "constante") donde la real sí regulaba, y viceversa.

4. **kpf residual:** menor influencia, pero un D mayor al real retrasa y eleva el nadir.

### 5.2 Qué hacer (y qué no)

**No recomendado:** forzar artificialmente la curva (escalar el tiempo, inyectar potencia ficticia, ajustar H sin evidencia). Se obtiene una coincidencia cosmética que no predice nada en el siguiente evento.

**Camino correcto — calibración de gobernadores con los datos que ya tienes:**

1. **Validar el ROCOF inicial** (primeros 1–2 s). Si coincide con el registro → H y ΔP del modelo son correctos y el problema está 100% en los gobernadores. (Ya verificado: coincide.)
2. **Clasificar las unidades con la tabla SCADA del evento** (columna "APORTA RPF"): en el modelo PF, las unidades que NO aportaron deben tener su regulación bloqueada o deadband grande; las que SÍ aportaron deben tener el droop calculado del registro (ej. WAR41: R=4.2%).
3. **Ajustar las constantes de tiempo** de los gobernadores de las unidades que sí respondieron, hasta que el tiempo de nadir simulado coincida. Para una turbina de gas: T_válvula ≈ 0.4–0.5 s, T_turbina ≈ 2–5 s.
4. **Validar contra un segundo evento** del mismo semestre sin re-ajustar nada. Si la curva coincide, la calibración es física y no un sobreajuste.

Este procedimiento (system identification con registros de eventos) es la práctica estándar de los operadores de red (NERC MOD-027 exige validar los modelos de gobernador contra eventos reales cada 5 años). Requiere acceso de escritura a los modelos DSL del proyecto PF — si no se dispone de él, la conclusión honesta es: **el script entrega la mejor condición inicial posible, y la discrepancia temporal restante es un problema de calibración de los modelos dinámicos del proyecto, no del proceso de carga.**

### 5.3 Resumen ejecutivo

| Aspecto | Estado |
|---|---|
| Condición inicial (despacho, demanda, pérdidas, slack) | La más fiel posible con los datos públicos del CNDC ✓ |
| Magnitud del nadir | Reproducida ✓ |
| ROCOF inicial | Reproducido ✓ |
| Tiempo al nadir / velocidad de recuperación | NO reproducido — requiere calibrar gobernadores DSL, fuera del alcance de este script |
