# Plan Técnico — Rediseño UI v4 en Streamlit
**Referencia:** `Interfaz RPF Rediseño v4.html` (handoff bundle) + `interfaz_analisis_RPF_v4_redesign.py`  
**App actual:** `Interfaz_RPF/interfaz_analisis_RPF.py` (~5 300 líneas)  
**Estado:** Plan de implementación — NO implementado aún  

---

## 1. Resumen del diseño objetivo

El diseño v4 (estilo Grafana/Datadog) introduce cuatro cambios estructurales sobre la UI actual:

| Componente | Estado actual | Estado objetivo |
|---|---|---|
| Barra superior | Ausente / nativo Streamlit | **Top bar fija** con logo, semestre + evento, chips de unidades desconectadas, toggles Local/Nube y tema claro/oscuro |
| Navegación de bloques | `st.segmented_control` horizontal | **Stepper horizontal** (8 pasos con numeración, indicador de completado) |
| Sidebar | Ausente | **Side nav colapsable** agrupado en Setup / Análisis / Salida |
| Barra de unidad | Badge HTML grande | **Unit context bar sticky** con selector inline y stats (Pmax, Tecnología, Droop, Estado) |
| Área de contenido | Bloques planos | Breadcrumb + número de bloque grande + controles compactos en fila + KPI strip 10 celdas + gráficas con header/toolbar |
| Tema | Solo modo oscuro | Modo claro / modo oscuro con design tokens |

---

## 2. Evaluación de factibilidad en Streamlit

### 2.1 ✅ Factible con CSS injection

Streamlit permite inyectar HTML/CSS arbitrario con `st.markdown(..., unsafe_allow_html=True)`.  
Todos los componentes de layout **decorativo** son factibles así.

| Componente | Técnica | Complejidad |
|---|---|---|
| Top bar HTML fija | `st.markdown` con `position:sticky; top:0; z-index:100` | ★★☆ |
| Stepper HTML | `st.markdown` completo del stepper | ★★☆ |
| Unit context bar | `st.markdown` + CSS sticky en key de container | ★★★ |
| KPI strip (10 celdas) | `st.markdown` con grid CSS | ★★☆ |
| Breadcrumb y block title | `st.markdown` simple | ★☆☆ |
| Tema claro/oscuro | TOKENS dict + f-strings en todo el CSS | ★★★ |

### 2.2 ⚠️ Factible con workarounds

| Componente | Limitación Streamlit | Workaround |
|---|---|---|
| Side nav colapsable | Streamlit sidebar no colapsa programáticamente; `initial_sidebar_state` no persiste | Usar `st.sidebar` con `initial_sidebar_state="expanded"` + CSS que oculta el hamburger nativo y estiliza los botones como navItems |
| Stepper interactivo | HTML puro no ejecuta callbacks | Cada paso es un `st.button` dentro de columnas ocultas; el HTML del stepper es solo visual, los botones invisibles detrás activan el bloque |
| Unit selector dropdown | El HTML dropdown no puede llamar a `st.session_state` | `st.selectbox` nativo estilizado con CSS override, o `st.popover` (Streamlit ≥ 1.31) |
| Sticky top bar | `position:sticky` en el HTML funciona, pero el scroll de Streamlit es del `stApp`-wrapper, no `window` | Posicionar dentro del `main` container con `position:sticky; top:0` (funciona en layout=wide) |
| Toggle Local/Nube + tema | Botones HTML no pueden cambiar `st.session_state` directamente | `st.toggle` nativo posicionado en la top bar con CSS override, o usar JavaScript via `st.components.v1.html` |

### 2.3 ❌ No factible / requiere componente custom

| Componente | Motivo |
|---|---|
| Stepper con connectors animados | CSS animado en un elemento que reacciona al `active_block` actual requiere JS — usar versión estática que se regenera en cada rerun |
| Side nav hover tooltip (collapsed mode) | CSS `:hover` sí funciona, pero el tooltip posicionado absolutamente dentro del sidebar de Streamlit puede salirse del clip |
| `position:sticky` del unit bar dentro de la columna principal | Funciona solo si la columna tiene `overflow:auto`; en Streamlit el scroll es del body global, lo que limita el sticky en columnas anidadas |

---

## 3. Arquitectura de implementación

### 3.1 Estrategia de migración

**No reescribir — envolver:**  
Toda la lógica de negocio (loaders, cálculos de KPIs, builders de gráficas) **permanece en `interfaz_analisis_RPF.py`**.  
Se crea un nuevo archivo `interfaz_analisis_RPF_v4.py` que:
1. Configura la página con el nuevo CSS
2. Renderiza los componentes de shell (topbar, stepper, sidenav, unit ctx)
3. Delega el contenido de cada bloque a funciones internas que contienen **el mismo código** que hoy viven en los bloques del archivo original

### 3.2 Estructura de archivos

```
Interfaz_RPF/
├── interfaz_analisis_RPF.py          ← INTACTO (app actual en producción)
├── interfaz_analisis_RPF_v4.py       ← NUEVO (shell v4 + misma lógica)
├── ui_v4/
│   ├── __init__.py
│   ├── tokens.py                     ← TOKENS dict (light + dark)
│   ├── css.py                        ← inject_css(t) → str
│   ├── components.py                 ← render_topbar(), render_stepper(), render_sidenav(), render_unit_ctx()
│   └── helpers.py                    ← icon(), badge(), kpi_strip(), block_header()
```

> **Opción alternativa más simple (recomendada para primera iteración):**  
> Todo en un único `interfaz_analisis_RPF_v4.py`, sin el sub-paquete `ui_v4/`, con secciones claramente delimitadas.

### 3.3 Session state adicional requerido

```python
# Nuevos keys de estado para la UI v4:
{
    "theme":        "light",          # "light" | "dark"
    "mode":         "cloud" | "local",  # ya existe IS_CLOUD — migrar a session_state
    "active_block": "scada",          # reemplaza bloque_trabajo
    "semestre":     "2024-1",
    "evento_id":    "E1",
    "sidenav_collapsed": False,
    # keys existentes se conservan tal cual
}
```

---

## 4. Componentes — Implementación detallada

### 4.1 Top bar

```python
def render_topbar(t, semestres, eventos_sem):
    """
    Inyecta la top bar HTML fija con:
    - Logo + título
    - Selectores semestre + evento (st.selectbox nativos dentro de columnas ocultas)
    - Chips de unidades desconectadas
    - Toggle tema y modo local/nube
    """
    # HTML decorativo (marca, layout)
    st.markdown(f"""
    <div class="topbar">
      <div class="topbar-brand">
        <div class="topbar-mark">{icon("bolt", 18, "#FFF")}</div>
        <div>
          <div class="topbar-title">RPF Analysis</div>
          <div class="topbar-sub">Análisis de Respuesta Primaria de Frecuencia</div>
        </div>
      </div>
      <!-- El centro y la derecha se renderizan como st.columns abajo -->
    </div>
    """, unsafe_allow_html=True)
    
    # Selectores nativos en columns ocultos con CSS
    col_sem, col_ev, col_disc, col_mode, col_theme = st.columns([1,2,3,1,1])
    with col_sem:
        st.selectbox("semestre", semestres, key="semestre", label_visibility="collapsed")
    with col_ev:
        st.selectbox("evento", [e["id"] for e in eventos_sem], key="evento_id", label_visibility="collapsed")
    # ...
```

> **Nota técnica:** Los `st.selectbox` con `label_visibility="collapsed"` se posicionan dentro del grid de la topbar usando `position:absolute` y CSS que mapea su `[data-testid]` al slot correcto del grid.

### 4.2 Stepper

El stepper es HTML puro + botones invisibles (técnica de "overlay"):

```python
def render_stepper(t, active_block, is_cloud, bloques):
    # 1. Calcular estado de cada paso
    active_idx = next(i for i, b in enumerate(bloques) if b["id"] == active_block)
    
    # 2. Renderizar HTML visual del stepper
    items_html = ""
    for i, b in enumerate(bloques):
        is_active = b["id"] == active_block
        is_past   = i < active_idx
        disabled  = is_cloud and b["requierePF"]
        # ...construir HTML de cada step...
    
    st.markdown(f'<div class="stepper-wrap"><div class="stepper-inner">{items_html}</div></div>',
                unsafe_allow_html=True)
    
    # 3. Botones reales invisibles (mismo order) para capturar clicks
    cols = st.columns(len(bloques))
    for i, (b, col) in enumerate(zip(bloques, cols)):
        with col:
            disabled = is_cloud and b["requierePF"]
            if st.button(b["short"], key=f"step_{b['id']}", disabled=disabled,
                         label_visibility="collapsed"):
                st.session_state.active_block = b["id"]
                st.rerun()
```

CSS que hace los botones transparentes superpuestos al HTML del stepper:
```css
.stepper-wrap + div[data-testid="stHorizontalBlock"] {
    position: absolute;  /* superpuesto */
    top: [offset calculado];
    opacity: 0;          /* invisible */
    pointer-events: auto;
    height: 56px;
}
```

> **Alternativa más simple:** Usar `st.segmented_control` (el que ya existe) pero estilizarlo con CSS para que luzca como el stepper. Mucho menos complejidad.

### 4.3 Side nav

```python
def render_sidenav(t, active_block, is_cloud, bloques):
    """Usa st.sidebar nativo + CSS overrides"""
    with st.sidebar:
        # Header
        col_title, col_collapse = st.columns([4, 1])
        # Grupos
        for grupo in ["Setup", "Análisis", "Salida"]:
            bloques_grupo = [b for b in bloques if b["grupo"] == grupo]
            st.markdown(f'<div class="nav-group-label">{grupo}</div>', unsafe_allow_html=True)
            for b in bloques_grupo:
                disabled = is_cloud and b["requierePF"]
                is_active = b["id"] == active_block
                if st.button(
                    f"{icon(b['icon'], 15)}  {b['label']}",
                    key=f"nav_{b['id']}",
                    disabled=disabled,
                    type="primary" if is_active else "secondary",
                    use_container_width=True,
                ):
                    st.session_state.active_block = b["id"]
                    st.rerun()
```

### 4.4 Unit context bar

```python
def render_unit_ctx(t, unidades, active_unit_code):
    u = next((x for x in unidades if x["codigo"] == active_unit_code), unidades[0])
    
    with st.container(key="unit_ctx"):
        col_label, col_sel, col_pmax, col_tech, col_droop, col_estado = st.columns([1, 2, 1, 1, 1, 1])
        
        with col_label:
            st.markdown(f'<span class="unit-ctx-label">Unidad activa</span>', unsafe_allow_html=True)
        
        with col_sel:
            st.selectbox(
                "unidad",
                [x["codigo"] for x in unidades],
                index=[x["codigo"] for x in unidades].index(active_unit_code),
                key="global_selected_unit",
                label_visibility="collapsed",
                format_func=lambda c: f"{c} — {next(x['nombre'] for x in unidades if x['codigo']==c)}",
            )
        
        with col_pmax:
            st.markdown(stat_html("P_max", f"{u['pmax']}", "MW"), unsafe_allow_html=True)
        with col_tech:
            st.markdown(stat_html("Tecnología", u["tech"], ""), unsafe_allow_html=True)
        with col_droop:
            st.markdown(stat_html("Droop", f"{u['droop']}", "%"), unsafe_allow_html=True)
        with col_estado:
            st.markdown(stat_html("Estado", "Conectada", "", color=t["success"]), unsafe_allow_html=True)
```

> **Nota:** El key `"unit_ctx"` es el que el CSS usa para aplicar `position:sticky`.  
> El `key="global_selected_unit"` es el **mismo** que ya usa `interfaz_analisis_RPF.py`, entonces toda la lógica existente que lee `st.session_state.global_selected_unit` sigue funcionando sin cambios.

### 4.5 KPI strip

```python
def render_kpi_strip(t, kpis: list[dict]):
    """
    kpis: [{"label": "f₀", "value": "59.98", "unit": "Hz", "tone": "neutral"}, ...]
    tone: "neutral" | "success" | "warning" | "danger"
    """
    cells = "".join([
        f'''<div class="kpi-cell kpi-{k["tone"]}">
              <div class="kpi-label">{k["label"]}</div>
              <div class="kpi-value">{k["value"]}<span class="kpi-unit">{k["unit"]}</span></div>
              {f'<div class="kpi-sub">{k["subtitle"]}</div>' if k.get("subtitle") else ""}
            </div>'''
        for k in kpis
    ])
    st.markdown(f'<div class="kpi-strip">{cells}</div>', unsafe_allow_html=True)
```

### 4.6 Block header (breadcrumb + número)

```python
def render_block_header(t, num: str, title: str, subtitle: str, breadcrumb: list[str]):
    crumb_html = " › ".join(
        f'<span class="breadcrumb-active">{x}</span>' if i == len(breadcrumb)-1
        else f'<span>{x}</span>'
        for i, x in enumerate(breadcrumb)
    )
    st.markdown(f'''
    <div class="block-body">
      <div class="breadcrumb">{crumb_html}</div>
      <div class="block-head">
        <div class="block-num">{num}</div>
        <div>
          <div class="block-title">{title}</div>
          <div class="block-sub">{subtitle}</div>
        </div>
      </div>
    </div>
    ''', unsafe_allow_html=True)
```

---

## 5. Mapeo de bloques actuales → v4

| Bloque actual (`bloque_trabajo`) | ID v4 (`active_block`) | Grupo sidebar | Requiere PF |
|---|---|---|---|
| `"Modelo Base"` | `"modelo"` | Setup | ✅ |
| `"Carga de Datos"` | `"carga"` | Setup | ✅ |
| `"Configuración"` | `"config"` | Setup | ❌ |
| `"Análisis SCADA/EMF"` | `"scada"` | Análisis | ❌ |
| `"Análisis Simulación"` | `"sim"` | Análisis | ✅ |
| `"Real vs Simulación"` | `"comp"` | Análisis | ✅ |
| `"Reporte Técnico"` | `"reporte"` | Salida | ❌ |
| `"Config. Gráficas"` | `"grafico"` | Salida | ❌ |

La variable `bloque_trabajo` del código actual se puede dejar como un **alias interno**:
```python
bloque_trabajo = st.session_state.get("active_block", "scada")
```
Esto permite que **todo el código existente siga leyendo `bloque_trabajo`** sin cambios.

---

## 6. Problemas conocidos de la UI actual que el rediseño resuelve

| Problema | Solución en v4 |
|---|---|
| La unidad activa se pierde visualmente al scrollear | Unit ctx bar sticky en la parte superior del área principal |
| El badge de unidad (HTML grande) ocupa mucho espacio | Barra compacta con stats al costado |
| No hay indicación de progreso en el flujo de 8 bloques | Stepper con estado completado/activo/futuro |
| Los bloques de Setup están mezclados con los de Análisis | Sidebar agrupado Setup / Análisis / Salida |
| Sin soporte tema claro | TOKENS dict con modo light/dark, toggle en top bar |
| "No se encontró el runner" en cloud | IS_CLOUD ya resuelto + badge "solo local" en stepper para bloques PF |

---

## 7. Plan de implementación por fases

### Fase 0 — Preparación (sin cambios en producción)
**Tiempo estimado: 1 sesión**

- [ ] Copiar `interfaz_analisis_RPF.py` → `interfaz_analisis_RPF_v4.py`
- [ ] Agregar `interfaz_analisis_RPF_v4.py` a `.gitignore` temporalmente (WIP)
- [ ] Crear archivo `ui_v4_css.py` con la función `inject_css(t)` completa (solo CSS, sin lógica)
- [ ] Verificar que `interfaz_analisis_RPF.py` sigue corriendo en cloud sin cambios

### Fase 1 — Shell básico (top bar + stepper + sidenav)
**Tiempo estimado: 2-3 sesiones**

- [ ] Implementar `inject_css()` con todos los tokens light/dark
- [ ] Implementar `render_topbar()` con selectores de semestre + evento + chips desconectadas
- [ ] Conectar `st.session_state.semestre` y `evento_global` al top bar
- [ ] Implementar stepper como HTML visual + `st.segmented_control` estilizado por debajo
- [ ] Implementar `render_sidenav()` usando `st.sidebar` nativo
- [ ] Verificar navegación entre los 8 bloques (funcionamiento idéntico al actual)
- [ ] **Checkpoint:** La app navega igual que la actual, pero con nuevo shell visual

### Fase 2 — Unit context bar
**Tiempo estimado: 1 sesión**

- [ ] Implementar `render_unit_ctx()` con `st.selectbox` estilizado
- [ ] Mantener `key="global_selected_unit"` (compatible con todo el código existente)
- [ ] Añadir stats de Pmax / Tecnología / Droop / Estado
- [ ] Eliminar el badge HTML grande de la barra global actual
- [ ] **Checkpoint:** La selección de unidad funciona igual, ahora en barra compacta

### Fase 3 — Cabeceras y KPI strip (Bloque 3)
**Tiempo estimado: 1-2 sesiones**

- [ ] Implementar `render_block_header()` con breadcrumb + número grande
- [ ] Aplicar a Bloque 3 (SCADA/EMF)
- [ ] Implementar `render_kpi_strip()` con los 10 KPIs de SCADA
- [ ] Conectar valores reales de `_cndc_kpis()` a la strip
- [ ] **Checkpoint:** Bloque 3 visualmente completo con diseño v4

### Fase 4 — Propagar cabeceras a los demás bloques
**Tiempo estimado: 1-2 sesiones**

- [ ] Aplicar `render_block_header()` a Bloques 0, 1, 2, 4, 5, 6, 7
- [ ] Bloque 5: añadir tabla comparativa estilizada (reemplaza `st.dataframe` actual)
- [ ] Bloque 4: KPI strip con resultados de simulación
- [ ] **Checkpoint:** Todos los bloques tienen header v4

### Fase 5 — Tema claro y toggle modo
**Tiempo estimado: 1 sesión**

- [ ] Conectar toggle de tema al `st.session_state.theme`
- [ ] Verificar que `inject_css(t)` reproduce correctamente ambos temas
- [ ] Conectar toggle Local/Nube a `st.session_state.mode` (referencia `IS_CLOUD`)
- [ ] **Checkpoint:** App funciona en modo claro y oscuro

### Fase 6 — Deployment
**Tiempo estimado: 0.5 sesión**

- [ ] Remover `interfaz_analisis_RPF_v4.py` del `.gitignore`
- [ ] Actualizar `requirements.txt` si se añadieron dependencias
- [ ] Cambiar el entry point en Streamlit Cloud a `Interfaz_RPF/interfaz_analisis_RPF_v4.py`
- [ ] O bien: renombrar `_v4.py` → reemplazar `interfaz_analisis_RPF.py`

---

## 8. Riesgos y mitigaciones

| Riesgo | Probabilidad | Impacto | Mitigación |
|---|---|---|---|
| CSS de `position:sticky` no funciona en algunas versiones de Streamlit | Media | Alto | Probar en Streamlit ≥ 1.32; alternativa: usar `st.container` sin sticky y duplicar la barra |
| Los `st.session_state` keys colisionan entre el shell v4 y el código de bloques | Baja | Alto | Auditar keys antes de Fase 1 con `grep "session_state\[" interfaz_analisis_RPF.py` |
| La sidebar nativa de Streamlit se "decolapsa" en mobile | Alta | Bajo | Aceptable; el diseño v4 ya está orientado a desktop |
| Cambiar `bloque_trabajo` a `active_block` rompe referencias en el código | Media | Alto | Usar alias `bloque_trabajo = st.session_state.get("active_block", "scada")` como primera línea de la sección de dispatch |
| El CSS inyectado es frágil ante actualizaciones de Streamlit | Media | Medio | Fijar `streamlit==X.Y.Z` en `requirements.txt` |

---

## 9. Decisiones de diseño a confirmar antes de implementar

1. **¿Se mantiene el `st.segmented_control` como stepper estilizado o se usa HTML+botones overlay?**  
   _Recomendación: mantener `st.segmented_control` + CSS override = mucho menos código, mismo efecto visual._

2. **¿Tema claro activado por defecto, o solo modo oscuro como ahora?**  
   _El diseño v4 muestra ambos; actualmente la app solo tiene dark._

3. **¿Se reemplaza `interfaz_analisis_RPF.py` o se crea un archivo nuevo paralelo?**  
   _Recomendación: archivo nuevo `_v4.py` hasta que esté validado, luego reemplazar._

4. **¿Se incluye el paquete `streamlit-extras` para componentes adicionales?**  
   _Actualmente no está en `requirements.txt`. Solo agregarlo si se necesita algo específico._

---

## 10. Referencia rápida de design tokens

```python
# light mode
PRIMARY   = "#2E5C8A"
ACCENT    = "#2563EB"   # Frecuencia (eje izquierdo)
ACCENT2   = "#F97316"   # Potencia (eje derecho)
SUCCESS   = "#10B981"
WARNING   = "#F59E0B"
DANGER    = "#DC2626"
BG        = "#F5F7FA"
SURFACE   = "#FFFFFF"
TEXT      = "#111827"
TEXT_MUTED= "#6B7280"

# dark mode (los que ya usa la app actual)
BG_DARK        = "#0B0F19"
SURFACE_DARK   = "#141925"
ACCENT_DARK    = "#60A5FA"
ACCENT2_DARK   = "#FB923C"
SUCCESS_DARK   = "#34D399"
```

---

*Documento creado: 2026-05-21*  
*App en producción: https://github.com/fernando22vg/rpf-interfaz-streamlit-*
