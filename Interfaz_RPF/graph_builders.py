"""
graph_builders.py
Funciones reutilizables para crear gráficas estándares en interfaz_analisis_RPF.py

Proporciona:
  · create_dual_axis_timeseries() - Gráfica de frecuencia + potencia
  · create_comparison_chart() - Comparativa real vs. simulación
  · apply_standard_layout() - Aplica formato estándar a figuras
  · add_kpi_markers() - Añade marcadores CNDC
  · apply_reference_lines() - Añade líneas de referencia
"""

import pandas as pd
import plotly.graph_objects as go
from copy import deepcopy
from graph_config import (
    COLOR_PALETTE, LINE_WIDTHS, MARKER_SIZES, MARKER_SYMBOLS,
    LAYOUT_PRESETS, XAXIS_TIME, XAXIS_TIME_HHMMSS, 
    YAXIS_FREQUENCY, YAXIS_POWER, LEGEND_CONFIGS,
    ANNOTATION_STYLES, DEFAULT_GRAPH_CONFIG
)


# ─────────────────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────────────────

def _to_plotly_time(t_val, show_hhmmss):
    """Convierte segundos a Datetime para que Plotly los maneje correctamente en el eje 'date'."""
    if not show_hhmmss:
        return t_val
    # Usar to_datetime asegura que Plotly interprete los segundos como marcas de tiempo reales desde el origen
    res = pd.to_datetime(t_val, unit='s')
    # Solución TypeError (Pandas 2.x): Plotly usa sum() internamente en anotaciones de add_vline/add_hline.
    # El objeto pd.Timestamp no permite sumas con el '0' inicial de sum(). Al devolver milisegundos 
    # desde la época (float), se permite el cálculo de promedios para posicionar etiquetas.
    if isinstance(res, pd.Timestamp):
        return res.timestamp() * 1000
    return res


def _get_config_value(key, default=None):
    """Obtiene un valor de configuración de gráficas (compatible con st.session_state)."""
    try:
        import streamlit as st
        if hasattr(st, 'session_state') and hasattr(st.session_state, 'graph_config'):
            cfg = st.session_state.graph_config
            return cfg.get(key, default)
    except (ImportError, AttributeError):
        pass
    
    # Fallback a configuración por defecto
    return DEFAULT_GRAPH_CONFIG.get(key, default)


# ─────────────────────────────────────────────────────────────────────────────
# CREACIÓN DE GRÁFICAS: DUAL AXIS (Frecuencia + Potencia)
# ─────────────────────────────────────────────────────────────────────────────

def create_dual_axis_timeseries(
    t_data,
    freq_data,
    pot_data,
    title="",
    freq_label="Frecuencia (Hz)",
    pot_label="Potencia (MW)",
    show_hhmmss=False,
    freq_color=None,
    pot_color=None,
    line_width=None,
    template=None,
    height=None,
    legend_position="bottom_center",
    x_range=None,
    y1_range=None,
    y2_range=None,
):
    """
    Crea una gráfica de series de tiempo con dos ejes Y (frecuencia y potencia).
    
    Args:
        t_data: Array/Series con tiempo (segundos)
        freq_data: Array/Series con datos de frecuencia
        pot_data: Array/Series con datos de potencia
        title: Título de la gráfica
        freq_label: Etiqueta del eje Y1 (frecuencia)
        pot_label: Etiqueta del eje Y2 (potencia)
        show_hhmmss: Si True, muestra tiempo en formato HH:MM:SS
        freq_color: Color de la línea de frecuencia (default: config)
        pot_color: Color de la línea de potencia (default: config)
        line_width: Grosor de línea (default: config)
        template: Template Plotly (default: config)
        height: Altura de la gráfica (default: config)
        legend_position: Posición de la leyenda (default: "bottom_center")
        x_range: Rango X [min, max] o None para auto
        y1_range: Rango Y1 [min, max] o None para auto
        y2_range: Rango Y2 [min, max] o None para auto
        
    Returns:
        plotly.graph_objects.Figure
    """
    # Obtener configuración por defecto si no se proporciona
    freq_color = freq_color or _get_config_value("freq_color_real", COLOR_PALETTE["freq_real"])
    pot_color = pot_color or _get_config_value("pot_color_real", COLOR_PALETTE["power_real"])
    line_width = line_width or _get_config_value("line_width", LINE_WIDTHS["normal"])
    template = template or _get_config_value("template", LAYOUT_PRESETS["default"]["template"])
    height = height or _get_config_value("plot_height", LAYOUT_PRESETS["default"]["height"])
    
    # Convertir tiempo a Plotly si es necesario
    t_plotly = _to_plotly_time(t_data, show_hhmmss)
    
    # Crear figura
    fig = go.Figure()
    
    # Añadir traza de frecuencia (eje Y1)
    fig.add_trace(go.Scatter(
        x=t_plotly,
        y=freq_data,
        name=freq_label,
        line=dict(color=freq_color, width=line_width),
        yaxis="y",
        hovertemplate=f"<b>{freq_label}</b><br>Tiempo: %{{x:.2f}} s<br>Valor: %{{y:.4f}} Hz<extra></extra>",
    ))
    
    # Añadir traza de potencia (eje Y2)
    fig.add_trace(go.Scatter(
        x=t_plotly,
        y=pot_data,
        name=pot_label,
        line=dict(color=pot_color, width=line_width),
        yaxis="y2",
        hovertemplate=f"<b>{pot_label}</b><br>Tiempo: %{{x:.2f}} s<br>Valor: %{{y:.3f}} MW<extra></extra>",
    ))
    
    # Configurar ejes X e Y
    xaxis_config = deepcopy(XAXIS_TIME_HHMMSS if show_hhmmss else XAXIS_TIME)
    if x_range:
        xaxis_config["range"] = [_to_plotly_time(x_range[0], show_hhmmss), 
                                 _to_plotly_time(x_range[1], show_hhmmss)]
    
    yaxis1_config = deepcopy(YAXIS_FREQUENCY)
    yaxis1_config["title"]["font"]["color"] = freq_color
    yaxis1_config["tickfont"]["color"] = freq_color
    if y1_range:
        yaxis1_config["range"] = y1_range
    
    yaxis2_config = deepcopy(YAXIS_POWER)
    yaxis2_config["title"]["font"]["color"] = pot_color
    yaxis2_config["tickfont"]["color"] = pot_color
    if y2_range:
        yaxis2_config["range"] = y2_range
    
    # Obtener configuración de layout
    layout_preset = deepcopy(LAYOUT_PRESETS["default"])
    layout_preset["template"] = template
    layout_preset["height"] = height
    
    legend_cfg = LEGEND_CONFIGS.get(legend_position, LEGEND_CONFIGS["bottom_center"])
    
    # Actualizar layout
    fig.update_layout(
        title=title,
        xaxis=xaxis_config,
        yaxis=yaxis1_config,
        yaxis2=yaxis2_config,
        legend=legend_cfg,
        **layout_preset,
    )
    
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# CREACIÓN DE GRÁFICAS: COMPARATIVA REAL VS. SIMULACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def create_comparison_chart(
    t_data,
    real_freq_data,
    simu_freq_data,
    real_pot_data,
    simu_pot_data,
    title="",
    show_hhmmss=False,
    line_width=None,
    template=None,
    height=None,
    legend_position="bottom_center",
    x_range=None,
    y1_range=None,
    y2_range=None,
):
    """
    Crea una gráfica comparativa real vs. simulación con doble eje Y.
    
    Args:
        t_data: Array/Series con tiempo (segundos)
        real_freq_data: Array/Series con frecuencia real/SCADA
        simu_freq_data: Array/Series con frecuencia simulada
        real_pot_data: Array/Series con potencia real
        simu_pot_data: Array/Series con potencia simulada
        title: Título de la gráfica
        show_hhmmss: Si True, muestra tiempo en formato HH:MM:SS
        line_width: Grosor de línea (default: config)
        template: Template Plotly (default: config)
        height: Altura de la gráfica (default: config)
        legend_position: Posición de la leyenda
        x_range, y1_range, y2_range: Rangos de ejes (None = auto)
        
    Returns:
        plotly.graph_objects.Figure
    """
    line_width = line_width or _get_config_value("line_width", LINE_WIDTHS["normal"])
    template = template or _get_config_value("template", LAYOUT_PRESETS["default"]["template"])
    height = height or _get_config_value("plot_height", LAYOUT_PRESETS["default"]["height"])
    
    # Convertir tiempo
    t_plotly = _to_plotly_time(t_data, show_hhmmss)
    
    # Crear figura
    fig = go.Figure()
    
    # Frecuencia Real (eje Y1)
    fig.add_trace(go.Scatter(
        x=t_plotly,
        y=real_freq_data,
        name="Frecuencia Real (Hz)",
        line=dict(color=COLOR_PALETTE["freq_real"], width=line_width),
        yaxis="y",
    ))
    
    # Frecuencia Simulada (eje Y1)
    fig.add_trace(go.Scatter(
        x=t_plotly,
        y=simu_freq_data,
        name="Frecuencia Simulada (Hz)",
        line=dict(color=COLOR_PALETTE["freq_simulated"], width=line_width, dash="dash"),
        yaxis="y",
    ))
    
    # Potencia Real (eje Y2)
    fig.add_trace(go.Scatter(
        x=t_plotly,
        y=real_pot_data,
        name="Potencia Real (MW)",
        line=dict(color=COLOR_PALETTE["power_real"], width=line_width),
        yaxis="y2",
    ))
    
    # Potencia Simulada (eje Y2)
    fig.add_trace(go.Scatter(
        x=t_plotly,
        y=simu_pot_data,
        name="Potencia Simulada (MW)",
        line=dict(color=COLOR_PALETTE["power_simulated"], width=line_width, dash="dash"),
        yaxis="y2",
    ))
    
    # Configurar ejes
    xaxis_config = deepcopy(XAXIS_TIME_HHMMSS if show_hhmmss else XAXIS_TIME)
    if x_range:
        xaxis_config["range"] = [_to_plotly_time(x_range[0], show_hhmmss), 
                                 _to_plotly_time(x_range[1], show_hhmmss)]
    
    yaxis1_config = deepcopy(YAXIS_FREQUENCY)
    if y1_range:
        yaxis1_config["range"] = y1_range
    
    yaxis2_config = deepcopy(YAXIS_POWER)
    if y2_range:
        yaxis2_config["range"] = y2_range
    
    layout_preset = deepcopy(LAYOUT_PRESETS["default"])
    layout_preset["template"] = template
    layout_preset["height"] = height
    legend_cfg = LEGEND_CONFIGS.get(legend_position, LEGEND_CONFIGS["bottom_center"])
    
    fig.update_layout(
        title=title,
        xaxis=xaxis_config,
        yaxis=yaxis1_config,
        yaxis2=yaxis2_config,
        legend=legend_cfg,
        **layout_preset,
    )
    
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# AÑADIR MARCADORES DE KPI CNDC
# ─────────────────────────────────────────────────────────────────────────────

def add_kpi_markers(
    fig,
    t_fault_abs,
    kpi_dict,
    show_hhmmss=False,
    dt_seconds=None,
    marker_size=None,
    freq_color=None,
    pot_color=None,
    # Si pasas tiempos plotables (en la misma escala del eje X de la figura),
    # los marcadores quedan 100% alineados con el gráfico.
    t0_plot=None,
    tmin_plot=None,
    tdt_plot=None,
):

    """
    Añade marcadores de KPI CNDC a una figura existente.
    
    Marcadores:
      · ○ f₀, P₀ en t₀ (círculos abiertos)
      · × f_min en nadir (cruz)
      · ● f_Δt, P_Δt en t₀+Δt (círculos rellenos)
    
    Args:
        fig: Figura go.Figure existente
        t_fault_abs: Tiempo absoluto de falla (segundos)
        kpi_dict: Dict con claves 'f0', 'p0', 'f_min', 't_min', 'f_dt', 'p_dt'
        show_hhmmss: Si True, convierte tiempo a formato HH:MM:SS
        dt_seconds: Duración Δt (para calcular t₀+Δt)
        marker_size: Tamaño de marcadores (default: config)
        freq_color: Color para marcadores de frecuencia (default: config)
        pot_color: Color para marcadores de potencia (default: config)
        
    Returns:
        Figura actualizada (go.Figure)
    """
    if not kpi_dict:
        return fig
    
    marker_size = marker_size or _get_config_value("marker_size", MARKER_SIZES["normal"])
    freq_color = freq_color or _get_config_value("freq_color_real", COLOR_PALETTE["freq_real"])
    pot_color = pot_color or _get_config_value("pot_color_real", COLOR_PALETTE["power_real"])
    
    show_initial = _get_config_value("show_initial", True)
    show_nadir = _get_config_value("show_nadir", True)
    show_dt_eval = _get_config_value("show_dt_eval", True)
    
    # Al usar eje 'date', los marcadores deben estar en la misma escala (datetime)
    
    # ○ f₀ y P₀ en t₀
    if show_initial:
        x0 = t0_plot if t0_plot is not None else _to_plotly_time(t_fault_abs, show_hhmmss)
        fig.add_trace(go.Scatter(
            x=[x0], y=[kpi_dict['f0']], mode='markers+text',

            marker=dict(symbol=MARKER_SYMBOLS["initial"], size=marker_size, 
                       color=freq_color, line=dict(width=2.5)),
            text=["  f₀"], textposition="top right", yaxis='y', 
            showlegend=False, hoverinfo='skip',
            name="KPI: f₀"
        ))
        fig.add_trace(go.Scatter(
            x=[x0], y=[kpi_dict['p0']], mode='markers+text',

            marker=dict(symbol=MARKER_SYMBOLS["initial"], size=marker_size, 
                       color=pot_color, line=dict(width=2.5)),
            text=["  P₀"], textposition="bottom right", yaxis='y2', 
            showlegend=False, hoverinfo='skip',
            name="KPI: P₀"
        ))
    
    # × f_min en nadir
    if show_nadir:
        x_min = tmin_plot if tmin_plot is not None else None
        if x_min is None:
            t_min_abs = t_fault_abs + kpi_dict['t_min']
            x_min = _to_plotly_time(t_min_abs, show_hhmmss)

        fig.add_trace(go.Scatter(
            x=[x_min], y=[kpi_dict['f_min']], mode='markers+text',
            marker=dict(symbol=MARKER_SYMBOLS["nadir"], size=marker_size, 
                       color=freq_color, line=dict(width=2.5)),
            text=["  f_min"], textposition="bottom right", yaxis='y', 
            showlegend=False, hoverinfo='skip',
            name="KPI: f_min"
        ))
    
    # ● f_Δt y P_Δt en t₀+Δt
    if show_dt_eval and dt_seconds is not None:
        x_dt = tdt_plot if tdt_plot is not None else None
        if x_dt is None:
            t_dt_abs = t_fault_abs + dt_seconds
            x_dt = _to_plotly_time(t_dt_abs, show_hhmmss)

        fig.add_trace(go.Scatter(
            x=[x_dt], y=[kpi_dict['f_dt']], mode='markers+text',
            marker=dict(symbol=MARKER_SYMBOLS["dt_eval"], size=marker_size, 
                       color=freq_color, line=dict(width=1.5, color='white')),
            text=["  f_Δt"], textposition="top right", yaxis='y', 
            showlegend=False, hoverinfo='skip',
            name="KPI: f_Δt"
        ))
        fig.add_trace(go.Scatter(
            x=[x_dt], y=[kpi_dict['p_dt']], mode='markers+text',
            marker=dict(symbol=MARKER_SYMBOLS["dt_eval"], size=marker_size,
                       color=pot_color, line=dict(width=1.5, color='white')),
            text=["  P_Δt"], textposition="top right", yaxis='y2',
            showlegend=False, hoverinfo='skip',
            name="KPI: P_Δt"
        ))
    
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# AÑADIR LÍNEAS DE REFERENCIA
# ─────────────────────────────────────────────────────────────────────────────

def add_reference_lines(
    fig,
    t_fault_abs=None,
    t_eval_abs=None,
    show_hhmmss=False,
    deadband_high=50.025,
    deadband_low=49.975,
    show_deadband=True,
    show_fault_line=True,
    show_eval_line=True,
    eval_line_label=None,
):
    """
    Añade líneas de referencia estándar a una figura.
    
    Líneas:
      · Banda muerta (deadband): líneas horizontales punteadas
      · Línea de falla: línea vertical en t₀
      · Línea de evaluación: línea vertical en t₀+Δt
    
    Args:
        fig: Figura go.Figure existente
        t_fault_abs: Tiempo absoluto de falla (segundos)
        t_eval_abs: Tiempo absoluto de evaluación (segundos)
        show_hhmmss: Si True, convierte tiempo a HH:MM:SS
        deadband_high: Valor alto de la banda muerta (default: 50.025 Hz)
        deadband_low: Valor bajo de la banda muerta (default: 49.975 Hz)
        show_deadband: Si True, muestra líneas de banda muerta
        show_fault_line: Si True, muestra línea vertical en t₀
        show_eval_line: Si True, muestra línea vertical en t₀+Δt
        eval_line_label: Etiqueta personalizada para línea de evaluación
        
    Returns:
        Figura actualizada (go.Figure)
    """
    # Banda muerta (líneas horizontales)
    if show_deadband:
        fig.add_hline(
            y=deadband_high,
            line_dash="dash",
            line_color=COLOR_PALETTE["deadband_line"],
            line_width=1,
            opacity=0.5,
            annotation_text="50.025 Hz",
            annotation_position="right",
        )
        fig.add_hline(
            y=deadband_low,
            line_dash="dash",
            line_color=COLOR_PALETTE["deadband_line"],
            line_width=1,
            opacity=0.5,
            annotation_text="49.975 Hz",
            annotation_position="right",
        )
    
    # Línea de falla (vertical)
    if show_fault_line and t_fault_abs is not None:
        t_fault_plotly = _to_plotly_time(t_fault_abs, show_hhmmss)
        fig.add_vline(
            x=t_fault_plotly,
            line_dash="dash",
            line_color=COLOR_PALETTE["fault_line"],
            line_width=1.5,
            annotation_text="t₀",
            annotation_position="top right",
            annotation_font_color=COLOR_PALETTE["fault_line"],
        )
    
    # Línea de evaluación (vertical)
    if show_eval_line and t_eval_abs is not None:
        t_eval_plotly = _to_plotly_time(t_eval_abs, show_hhmmss)
        label = eval_line_label or f"t₀+Δt"
        fig.add_vline(
            x=t_eval_plotly,
            line_dash="dash",
            line_color=COLOR_PALETTE["eval_line"],
            line_width=1.5,
            annotation_text=label,
            annotation_position="top left",
            annotation_font_color=COLOR_PALETTE["eval_line"],
        )
    
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# APLICAR LAYOUT ESTÁNDAR
# ─────────────────────────────────────────────────────────────────────────────

def create_scada_emf_comparison_chart(
    t_scada_aligned,
    scada_freq,
    scada_pot,
    t_emf_aligned,
    emf_freq,
    emf_pot,
    title="",
    show_hhmmss=False,
    line_width=None,
    template=None,
    height=None,
    legend_position="bottom_center",
    x_range=None,
    y1_range=None,
    y2_range=None,
    show_reference=True,
    t_fault_abs_scada=0.0,
    t_eval_abs_scada=35.0,
    show_deadband=None,
    show_fault_line=True,
    show_eval_line=True,
    eval_line_label=None,
    kpi_dict=None,
    dt_seconds=None,
):
    """Builder estandarizado para Comparativa: SCADA (real) vs EMF/CNDC.

    - Aplica siempre: contrato de ejes (doble eje Y), colores/config estándar, referencias CNDC y layout estándar.
    - No toca la lógica de cálculo de KPIs: solo consume `kpi_dict` si se provee.
    """
    _gcfg_show_deadband = _get_config_value("show_deadband", True) if show_deadband is None else show_deadband
    line_width = line_width or _get_config_value("line_width", LINE_WIDTHS["normal"])
    template = template or _get_config_value("template", LAYOUT_PRESETS["default"]["template"])
    height = height or _get_config_value("plot_height", LAYOUT_PRESETS["default"]["height"])

    # Base con SCADA usando el builder estándar
    fig = create_dual_axis_timeseries(
        t_data=t_scada_aligned,
        freq_data=scada_freq,
        pot_data=scada_pot,
        title=title,
        freq_label="Frecuencia SCADA (Hz)",
        pot_label="Potencia (MW)",
        show_hhmmss=show_hhmmss,
        freq_color=_get_config_value("freq_color_real", COLOR_PALETTE["freq_real"]),
        pot_color=_get_config_value("pot_color_real", COLOR_PALETTE["power_real"]),
        line_width=line_width,
        template=template,
        height=height,
        legend_position=legend_position,
        x_range=x_range,
        y1_range=y1_range,
        y2_range=y2_range,
    )

    # Capas EMF/CNDC usando colores/config estándar (mismas keys que se usan para real en esta UI)
    # SCADA y EMF comparten convención visual: si quieres otra paleta por EMF, se hará vía graph_config.
    emf_freq_color = _get_config_value("freq_color_real", COLOR_PALETTE["freq_real"])
    emf_pot_color = _get_config_value("pot_color_real", COLOR_PALETTE["power_real"])

    fig.add_trace(
        go.Scatter(
            x=_to_plotly_time(t_emf_aligned, show_hhmmss),
            y=emf_freq,
            name="Frecuencia CNDC (EMF)",
            line=dict(color=emf_freq_color, width=line_width),
            yaxis="y",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=_to_plotly_time(t_emf_aligned, show_hhmmss),
            y=emf_pot,
            name="Potencia CNDC (EMF)",
            line=dict(color=emf_pot_color, width=line_width),
            yaxis="y2",
        )
    )

    if show_reference:
        fig = add_reference_lines(
            fig,
            t_fault_abs=t_fault_abs_scada,
            t_eval_abs=t_eval_abs_scada,
            show_hhmmss=show_hhmmss,
            show_deadband=_gcfg_show_deadband,
            show_fault_line=show_fault_line,
            show_eval_line=show_eval_line,
            eval_line_label=eval_line_label,
        )

    # KPIs opcionales
    if kpi_dict:
        fig = add_kpi_markers(
            fig,
            t_fault_abs=t_fault_abs_scada,
            kpi_dict=kpi_dict,
            show_hhmmss=show_hhmmss,
            dt_seconds=dt_seconds,
            marker_size=_get_config_value("marker_size", MARKER_SIZES["normal"]),
            freq_color=_get_config_value("freq_color_real", COLOR_PALETTE["freq_real"]),
            pot_color=_get_config_value("pot_color_real", COLOR_PALETTE["power_real"]),
        )

    # Layout estándar al final (para asegurar consistencia)
    fig = apply_standard_layout(
        fig,
        title=title,
        legend_position=legend_position,
        template=template,
        height=height,
        show_grid=_get_config_value("show_grid", True),
    )

    return fig


def apply_standard_layout(
    fig,
    title="",
    xaxis_title=None,
    yaxis_title=None,
    yaxis2_title=None,
    preset="default",
    template=None,
    height=None,
    legend_position="bottom_center",
    show_grid=True,
):
    """
    Aplica un layout estándar a una figura existente.
    
    Args:
        fig: Figura go.Figure existente
        title: Título de la gráfica
        preset: Nombre del preset ("default", "compact", "expanded")
        template: Template Plotly personalizado (sobrescribe preset)
        height: Altura personalizada (sobrescribe preset)
        legend_position: Posición de la leyenda
        show_grid: Si True, muestra grilla
        
    Returns:
        Figura actualizada (go.Figure)
    """
    # Obtener preset base
    layout_config = deepcopy(LAYOUT_PRESETS.get(preset, LAYOUT_PRESETS["default"]))
    
    # Sobrescribir si se proporcionan valores personalizados
    if template:
        layout_config["template"] = template
    if height:
        layout_config["height"] = height
    
    # Configurar grilla
    if not show_grid:
        layout_config["xaxis"] = {"showgrid": False}
        layout_config["yaxis"] = {"showgrid": False}
    
    # Aplicar leyenda
    legend_cfg = LEGEND_CONFIGS.get(legend_position, LEGEND_CONFIGS["bottom_center"])
    layout_config["legend"] = legend_cfg
    
    # Actualizar figura
    fig.update_layout(
        title=title,
        **layout_config,
    )

    if xaxis_title: fig.update_layout(xaxis_title=xaxis_title)
    if yaxis_title: fig.update_layout(yaxis_title=yaxis_title)
    if yaxis2_title: fig.update_layout(yaxis2={"title": {"text": yaxis2_title}})
    
    return fig
