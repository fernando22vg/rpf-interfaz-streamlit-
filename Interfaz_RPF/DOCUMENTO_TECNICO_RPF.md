# Documento Técnico — Interfaz de Análisis RPF
**Sistema Interconectado Nacional de Bolivia (SIN)**
**COBEE S.A. — Área de Estudios y Operación**
**Versión 2.0 — Mayo 2026**

---

## 1. Resumen Ejecutivo

La **Interfaz de Análisis RPF** es una aplicación web desarrollada en Python/Streamlit que centraliza y automatiza el flujo de trabajo completo para la evaluación de la **Respuesta Primaria de Frecuencia (RPF)** de unidades generadoras del SIN boliviano.

El proceso de evaluación RPF exige comparar datos reales de eventos (SCADA/EMF) con simulaciones RMS en DIgSILENT PowerFactory, calcular indicadores de desempeño según la metodología oficial del CNDC, y generar reportes técnicos documentados. Antes de esta herramienta, ese flujo se ejecutaba con scripts Python independientes, hojas Excel manuales y sin trazabilidad. La app unifica todo en un panel único, accesible tanto localmente como desde Streamlit Cloud.

---

## 2. Contexto Normativo

| Organismo | Rol |
|---|---|
| **CNDC** | Comité Nacional de Despacho de Carga — define metodología RPF y evalúa cumplimiento |
| **COBEE S.A.** | Empresa generadora — debe demostrar aporte de RPF en cada evento de frecuencia |
| **Regulación** | Las unidades deben aportar ΔP% ≥ 1.5% de P_max en Δt segundos tras la falla |

### Indicadores KPI que define la metodología CNDC

| KPI | Descripción |
|---|---|
| **f₀** | Frecuencia en el instante de la falla |
| **P₀** | Potencia activa en el instante de la falla |
| **f_min** | Frecuencia mínima (nadir) tras la falla |
| **t_min** | Tiempo al nadir [s] |
| **Δf** | f₀ − f_min |
| **f_Δt** | Frecuencia en t₀+Δt |
| **P_Δt** | Potencia activa en t₀+Δt |
| **ΔP** | P_Δt − P₀ [MW] — aporte real de la unidad |
| **ΔP%** | ΔP / P_max × 100 — criterio de cumplimiento |
| **Droop** | Estatismo calculado [%] |
| **ROCOF** | Tasa de cambio de frecuencia [Hz/s] — por regresión lineal en [0, 3s] |

---

## 3. Arquitectura del Sistema

```
┌─────────────────────────────────────────────────────────────────┐
│                    interfaz_analisis_RPF.py                     │
│                    (Streamlit — app principal)                   │
├───────────┬──────────────┬──────────────┬───────────────────────┤
│graph_     │graph_        │sharepoint_   │  Scripts Python        │
│builders.py│config.py     │client.py     │  (subprocesos)         │
│           │              │              │                        │
│Funciones  │Constantes    │Acceso        │ExtFLujos2daO.py        │
│de gráficas│y estilos     │SharePoint    │CondInicialesPF.py      │
│Plotly     │              │vía cookies   │CargaCondIniciales_PF   │
│           │              │              │OrdenadorDatosEvento    │
│           │              │              │ExtractorResultadosCNDC │
│           │              │              │DatosCurvas_v3.py (PF)  │
└───────────┴──────────────┴──────────────┴───────────────────────┘
         │                              │
    Streamlit Cloud              Instalación local
    (nube, solo lectura)         DIgSILENT PowerFactory
         │                              │
    SharePoint OneDrive          Disco local Windows
    (datos vía REST API)         C:\Datos del CNDC\
```

### Módulos principales

| Archivo | Función |
|---|---|
| `interfaz_analisis_RPF.py` | App principal — lógica de UI, KPIs, coordinación |
| `graph_builders.py` | Funciones de gráficas Plotly reutilizables |
| `graph_config.py` | Paletas, tamaños, configuraciones de ejes |
| `sharepoint_client.py` | Cliente SharePoint vía sesión/cookies sin Azure App Registration |
| `ExtFLujos2daO.py` | Extrae flujos de segunda orden desde archivos CNDC |
| `CondInicialesPF.py` | Calcula condiciones iniciales para PowerFactory |
| `CargaCondIniciales_PF.py` | Carga condiciones iniciales en el modelo PowerFactory |
| `OrdenadorDatosEvento.py` | Organiza estructura de carpetas por semestre/evento |
| `ExtractorResultadosCNDC.py` | Extrae resultados de simulaciones CNDC |
| `DatosCurvas_v3.py` | Extrae curvas RMS desde PowerFactory (se ejecuta dentro de PF) |

---

## 4. Flujo de Trabajo — 8 Bloques

```
[0] Modelo Base  →  [1] Carga Datos  →  [2] Config. Unidades
                                                  ↓
[7] Config Gráf]  [6] Reporte]  ←  [5] Real vs Simu]  ←  [3] Análisis SCADA/EMF
                                          ↑                        ↓
                                   [4] Análisis Simu]  ←  DIgSILENT PowerFactory
```

### Bloque 0 — Datos del Modelo
- Carga inventario de generadores, líneas, transformadores, shunts desde PowerFactory
- Ejecuta scripts `loc_namesGEN`, `loc_names_xfo`, `loc_namesLineas`, `InventarioShunts_PF`
- Genera archivos de mapeo `loc_names_*.xlsx` usados en bloques posteriores

### Bloque 1 — Carga de Datos
**Tab 1: Extracción CNDC**
- Ejecuta `ExtFLujos2daO.py` sobre los datos del evento
- Procesa archivos del CNDC (postot/td_ de la STI)
- Genera tabla de flujos de potencia y topología del SIN en el momento del evento

**Tab 2: Condiciones Iniciales**
- Ejecuta `CondInicialesPF.py`
- Calcula P₀, Q₀, V₀, θ₀ de cada unidad y barra en el instante previo a la falla
- Formato listo para carga en PowerFactory

**Tab 3: Carga en PowerFactory**
- Ejecuta `CargaCondIniciales_PF_run.py` vía subprocess
- Inicializa el modelo PowerFactory con las condiciones del evento real
- Lanza simulaciones E{N}.0 (modelo CNDC) y E{N}.1 (modelo COBEE)

### Bloque 2 — Configuración de Unidades
- Define P_max, droop nominal (Rp%) y tecnología por unidad generadora
- Fuentes: `datos_cargados.json` (manual) o `loc_names_gen.xlsx` (desde PF)
- Persiste configuración por evento en `event_config.json`

### Bloque 3 — Análisis de Datos Registrados (SCADA / EMF)
**Pestaña SCADA:**
- Lee archivos Excel de registros SCADA (frecuencia Hz + potencia MW + tiempo HH:MM:SS)
- Auto-detecta instante de falla por df/dt sostenido (umbral configurable)
- Calcula todos los KPI CNDC con metodología oficial
- Gráfica doble eje Y (frecuencia + potencia) con marcadores f₀, f_min, f_Δt, P₀, P_Δt
- Soporta visualización en segundos relativos o HH:MM:SS (hora del día real)

**Pestaña EMF:**
- Mismo flujo que SCADA pero para registros digitalizados de curvas EMF
- Los datos EMF tienen origen en la digitalización de registros analógicos del CNDC

**Pestaña Comparativa:**
- Superpone curvas SCADA + EMF en mismos ejes para verificar consistencia

### Bloque 4 — Análisis de Simulación (PowerFactory)
- Lee resultados de simulación `E{N}.0` (CNDC) y `E{N}.1` (COBEE) en Excel
- Auto-detecta t₀ desde señal de frecuencia simulada
- Calcula KPI CNDC sobre los resultados de simulación
- Pestaña comparativa E0 vs E1 para ver diferencias entre modelos

### Bloque 5 — Comparativa Real vs. Simulación
- Alinea curvas reales (SCADA/EMF) con curvas simuladas en ejes comunes
- Permite evaluar calidad del modelo: ¿reproduce la respuesta real?
- Escala de ejes sincronizada entre bloques para comparación visual justa

### Bloque 6 — Reporte Técnico
- Consolida KPIs de todas las fuentes (SCADA, EMF, Sim E0, Sim E1)
- Genera tabla resumen de cumplimiento RPF por unidad
- Exporta a Excel con formato corporativo (colores, bordes, semaforización)

### Bloque 7 — Configuración de Gráficas
- Paleta de colores, grosor de línea, tamaño de marcadores, altura de gráfica
- Plantilla Plotly (plotly_white, ggplot2, seaborn, etc.)
- Opciones de visualización de banda muerta (±25 mHz) y líneas de referencia

---

## 5. Problemas que Resuelve

### 5.1 Workflow fragmentado
**Antes:** 8+ scripts Python independientes ejecutados manualmente en orden específico, con rutas hardcodeadas, sin interfaz gráfica, propensos a errores de secuencia.
**Después:** Un solo panel con flujo guiado, botones de ejecución con logs en vivo, bloqueo de navegación durante procesos activos.

### 5.2 Cálculo manual de KPIs
**Antes:** Los indicadores f₀, P_Δt, ΔP%, droop se calculaban en Excel caso por caso, sin metodología consistente.
**Después:** Función `_cndc_kpis()` implementa la metodología oficial CNDC paso a paso. Resultado reproducible, trazable, exportable.

### 5.3 Detección manual de t₀
**Antes:** El analista inspeccionaba visualmente la curva de frecuencia para estimar cuándo comenzó la falla.
**Después:** `_detectar_inicio_falla()` usa análisis de df/dt sobre señal suavizada. El analista solo confirma o ajusta con precisión de ±1 segundo via `number_input`.

### 5.4 Inconsistencia entre datos reales y simulados
**Antes:** Sin herramienta para comparar directamente curvas SCADA con simulaciones RMS.
**Después:** Bloque 5 alinea automáticamente curvas reales y simuladas con el mismo t₀ y ejes, facilitando validación del modelo.

### 5.5 Acceso a datos en nube
**Antes:** La app solo corría en la PC del analista con acceso local a `C:\Datos del CNDC\`.
**Después:** Desplegada en Streamlit Cloud, lee datos desde SharePoint vía sesión HTTP/cookies. Sin Azure App Registration. Los botones de PowerFactory se deshabilitan en nube (solo disponibles con instalación local).

### 5.6 Formato de tiempo incorrecto en gráficas
**Antes:** Los datos SCADA en HH:MM:SS se mostraban siempre como segundos relativos (perdiendo la hora del día real).
**Después:** Checkbox "Mostrar tiempo en HH:MM:SS" muestra la hora exacta del evento en el eje X, con conversión correcta a milisegundos de época para Plotly.

---

## 6. Limitaciones Actuales

| Limitación | Descripción |
|---|---|
| **PowerFactory solo local** | Los scripts que invocan PF requieren la instalación en la PC del analista. En nube los botones están deshabilitados. |
| **Estructura de carpetas fija** | La app asume jerarquía `Semestre/Análisis_todos_los_eventos/EventoN/E{N}.0/...`. Cambios en estructura rompen la app. |
| **Un evento a la vez** | No hay modo batch para procesar múltiples eventos simultáneamente. |
| **Sin autenticación** | Cualquier persona con la URL de Streamlit Cloud puede ver los datos. |
| **Almacenamiento en archivos** | Toda la configuración y resultados están en archivos Excel/JSON locales o SharePoint. Sin base de datos. |
| **Reporte no automatizado** | El Bloque 6 genera tabla resumen pero no produce un PDF/DOCX listo para entregar al CNDC. |
| **Sin versionado de modelos** | No hay trazabilidad de qué versión del modelo PowerFactory generó qué simulación. |

---

## 7. Mejoras Propuestas

### 7.1 Prioridad Alta

**A. Generación automática de informe PDF/Word**
- Consolidar KPIs + gráficas en un informe con plantilla corporativa
- Tecnologías: `reportlab` o `python-docx`
- Impacto: elimina trabajo manual de armado del informe CNDC

**B. Procesamiento batch de eventos**
- Seleccionar múltiples eventos y ejecutar análisis completo en todos
- Generar tabla resumen comparativa de cumplimiento RPF del semestre
- Impacto: reduce tiempo de análisis semestral de días a minutos

**C. Autenticación de usuarios**
- Integrar Streamlit Cloud secrets + OAuth o contraseña simple
- Separar roles: visualizador (solo lectura) vs. analista (puede ejecutar scripts)

### 7.2 Prioridad Media

**D. Agente local para PowerFactory remoto**
- Servidor FastAPI (~50 líneas) en la PC con PowerFactory
- Túnel ngrok para exponer URL pública a Streamlit Cloud
- El botón de la nube hace `requests.post(ngrok_url, ...)` en vez de `subprocess`
- Documentado en `DIGSILENT_REMOTO.md`

**E. Base de datos SQLite/PostgreSQL**
- Almacenar KPIs calculados, configuraciones y metadatos de eventos
- Permite consultas históricas: "¿cómo evolucionó el droop de UNI_01 en los últimos 6 meses?"
- Reemplaza los `event_config.json` dispersos

**F. Validación automática de archivos de entrada**
- Verificar columnas, unidades y rangos de datos SCADA/EMF al cargar
- Alertas tempranas antes de calcular KPIs incorrectos
- Actualmente los errores de datos llegan tarde (en el cálculo)

**G. Umbral de calidad de señal SCADA**
- Detectar gaps, valores anómalos (spikes) y saturaciones en la señal
- Marcar datos sospechosos antes del cálculo de KPIs

### 7.3 Prioridad Baja / Largo Plazo

**H. Análisis estadístico multievento**
- Histogramas de ΔP%, droop, ROCOF por unidad y por semestre
- Tendencias: ¿está deteriorándose la respuesta de alguna unidad?

**I. Integración directa con API del CNDC**
- Si el CNDC publica APIs de datos históricos, reemplazar la carga manual de archivos
- Sincronización automática de nuevos eventos

**J. Machine Learning para clasificación de eventos**
- Clasificar tipo de perturbación (salida de generación, salida de línea, etc.) desde la firma de frecuencia
- Estimar unidad causante de la perturbación

**K. Módulo de sensibilidad de parámetros**
- Variar parámetros del modelo PF (droop, inercia, ganancia) y ver efecto en KPIs
- Útil para optimizar configuración de gobernadores

---

## 8. Dependencias y Entorno

### Requisitos de ejecución local
```
Python 3.12
streamlit==1.40.2
pandas==2.2.3
numpy==1.26.4
plotly==5.24.0
openpyxl==3.1.5
python-calamine>=0.2.0   # lectura rápida de Excel
xlrd==2.0.1
scikit-learn==1.5.2
scipy==1.14.1
pyarrow==17.0.0
requests>=2.32.0
DIgSILENT PowerFactory (para Bloques 1/4)
```

### Ejecución local
```bash
streamlit run ProgramasLimpio\interfaz_analisis_RPF.py
```

### Ejecución en nube
- URL: `https://interfazrpf.streamlit.app`
- Datos: SharePoint OneDrive (COBEE) vía URL compartida
- PowerFactory: deshabilitado (solo análisis y visualización)

### Estructura de carpetas esperada
```
C:\Datos del CNDC\
├── 01_INFO CNDC_RPF\
│   ├── 2024-1\
│   │   ├── Tabla_Eventos_2024-1.xlsx
│   │   └── Análisis_todos_los_eventos\
│   │       ├── Evento_001\
│   │       │   ├── E1.0\Datos Curvas\UNI_01.xlsx
│   │       │   ├── E1.1\Datos Curvas\UNI_01.xlsx
│   │       │   └── Resultados_COBEE\UNI_01.xlsx
│   │       └── Evento_002\...
└── DATOS EXTRAIDOS DE DIGSILENT\
    └── Designacion de loc_name\
        ├── loc_names_gen.xlsx
        ├── loc_name_cargas.xlsx
        └── loc_names_xfo.xlsx
```

---

## 9. Decisiones de Diseño Relevantes

| Decisión | Razón |
|---|---|
| Streamlit (no Flask/Django) | Prototipado rápido, sin HTML/CSS, adecuado para ingenieros no web |
| `python-calamine` para Excel | 5-10× más rápido que openpyxl para solo lectura de archivos grandes |
| SharePoint vía cookies (no Azure) | No requiere permisos de Azure AD Admin — viable en organización sin acceso administrativo |
| `IS_CLOUD = not os.path.isdir(r"C:\Datos del CNDC")` | Detección simple y robusta: False en Windows con datos, True en Linux (Streamlit Cloud) |
| KPIs en tiempo alineado (`t_al = t_raw - t_falla`) | Permite comparar curvas de distintos eventos en la misma escala temporal relativa |
| `_to_plotly_time` con ms de época para escalares | Evita TypeError de Pandas 2.x en `add_vline`/`add_hline` que usa `sum()` internamente |
| `number_input` en lugar de slider para t₀ | Slider por índice imposibilita precisión de ±1s; `number_input` permite teclear el segundo exacto |
| Auto-reset de t₀ al cambiar evento | Evita arrastrar un t₀ incorrecto de un evento anterior a otro |

---

## 10. Historial de Cambios Recientes

| Fecha | Cambio |
|---|---|
| May 2026 | Despliegue en Streamlit Cloud + cliente SharePoint sin Azure App Registration |
| May 2026 | Reemplazo de sliders de t₀ por `number_input` en segundos + botón auto-detección |
| May 2026 | Fix modo HH:MM:SS: corrección de escala ×1000 en marcadores KPI y líneas de referencia |
| May 2026 | Bloque 4: auto-detección de t₀ desde señal de frecuencia simulada |
| May 2026 | Bloque 4: botones para guardar y resetear t₀ y Δt por evento |
| May 2026 | Fix bloque 1: tabla de estado de archivos ya no se duplica en bucle |
| May 2026 | Creación de `DIGSILENT_REMOTO.md` con arquitectura de agente local vía ngrok |

---

*Documento generado automáticamente — COBEE S.A. / Área de Estudios*
