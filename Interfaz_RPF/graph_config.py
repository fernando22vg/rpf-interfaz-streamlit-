"""
graph_config.py
Configuración centralizada para todas las gráficas de interfaz_analisis_RPF.py

Define:
  · Paleta de colores estándar
  · Presets de layouts y ejes
  · Convenciones de estilo de Plotly
"""

# ─────────────────────────────────────────────────────────────────────────────
# PALETA DE COLORES ESTÁNDAR
# ─────────────────────────────────────────────────────────────────────────────
COLOR_PALETTE = {
    # Colores principales (datos reales vs. simulación)
    "freq_real":      "#1f77b4",      # Azul fuerte - Frecuencia SCADA/Real
    "freq_simulated": "#ff7f0e",      # Naranja - Frecuencia simulada
    "power_real":     "#2ca02c",      # Verde - Potencia SCADA/Real
    "power_simulated":"#d62728",      # Rojo - Potencia simulada
    
    # Colores de marcadores KPI (CNDC)
    "marker_initial":   "#1f77b4",    # Azul - f₀, P₀ iniciales
    "marker_nadir":     "#ff7f0e",    # Naranja - f_min (nadir)
    "marker_dt_eval":   "#2ca02c",    # Verde - f_Δt, P_Δt (evaluación)
    "marker_error":     "#d62728",    # Rojo - Errores/fuera de banda
    
    # Líneas de referencia
    "deadband_line":    "#808080",    # Gris - Banda muerta
    "fault_line":       "#808080",    # Gris - Línea de falla t₀
    "eval_line":        "#4682b4",    # Steel blue - Línea de evaluación t₀+Δt
    "grid":             "#d3d3d3",    # Gris claro - Grilla
    
    # Colores de estado
    "success":          "#C6EFCE",     # Verde claro - Estado OK
    "warning":          "#FFEB9C",     # Amarillo - Advertencia
    "error":            "#FFCCCC",     # Rojo claro - Error
    
    # Colores adicionales
    "text_dark":        "#333333",
    "text_light":       "#999999",
}

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE LÍNEAS Y MARCADORES
# ─────────────────────────────────────────────────────────────────────────────
LINE_STYLES = {
    "solid":     "solid",
    "dash":      "dash",
    "dot":       "dot",
    "dashdot":   "dashdot",
}

MARKER_SYMBOLS = {
    "initial":   "circle-open",    # ○ - Círculo abierto
    "nadir":     "x",              # × - Cruz
    "dt_eval":   "circle",         # ● - Círculo relleno
    "point":     "circle-dot",     # ◐ - Punto
}

LINE_WIDTHS = {
    "thin":      1.0,
    "normal":    2.0,
    "thick":     2.5,
    "bold":      3.0,
}

MARKER_SIZES = {
    "small":     6,
    "normal":    10,
    "large":     14,
    "xlarge":    18,
}

# ─────────────────────────────────────────────────────────────────────────────
# PRESETS DE LAYOUT PLOTLY
# ─────────────────────────────────────────────────────────────────────────────
LAYOUT_PRESETS = {
    "default": {
        "template": "plotly_white",
        "height": 700,
        "font": {
            "family": "Arial, sans-serif",
            "size": 11,
            "color": COLOR_PALETTE["text_dark"],
        },
        "margin": {
            "l": 80,
            "r": 80,
            "t": 60,
            "b": 50,
            "pad": 10,
        },
        "hovermode": "x unified",
        "plot_bgcolor": "rgba(240, 240, 240, 0.5)",
        "paper_bgcolor": "white",
    },
    
    "compact": {
        "template": "plotly_white",
        "height": 400,
        "font": {
            "family": "Arial, sans-serif",
            "size": 10,
        },
        "margin": {
            "l": 60,
            "r": 60,
            "t": 40,
            "b": 40,
        },
        "hovermode": "x unified",
    },
    
    "expanded": {
        "template": "plotly_white",
        "height": 800,
        "font": {
            "family": "Arial, sans-serif",
            "size": 12,
        },
        "margin": {
            "l": 100,
            "r": 100,
            "t": 80,
            "b": 80,
        },
        "hovermode": "x unified",
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE EJE X (Tiempo)
# ─────────────────────────────────────────────────────────────────────────────
XAXIS_TIME = {
    "title": "Tiempo (s)",
    "showgrid": True,
    "gridwidth": 1,
    "gridcolor": COLOR_PALETTE["grid"],
    "zeroline": False,
    "showline": True,
    "linewidth": 1,
    "linecolor": COLOR_PALETTE["text_dark"],
}

XAXIS_TIME_HHMMSS = {
    "title": "Tiempo HH:MM:SS",
    "type": "date",
    "tickformat": "%H:%M:%S",
    "showgrid": True,
    "gridwidth": 1,
    "gridcolor": COLOR_PALETTE["grid"],
    "zeroline": False,
    "showline": True,
    "linewidth": 1,
    "linecolor": COLOR_PALETTE["text_dark"],
}

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE EJES Y (Frecuencia, Potencia)
# ─────────────────────────────────────────────────────────────────────────────
YAXIS_FREQUENCY = {
    "title": {
        "text": "Frecuencia (Hz)",
        "font": {"color": COLOR_PALETTE["freq_real"]},
    },
    "tickfont": {"color": COLOR_PALETTE["freq_real"]},
    "showgrid": True,
    "gridwidth": 1,
    "gridcolor": COLOR_PALETTE["grid"],
    "zeroline": False,
    "showline": True,
    "linewidth": 1,
    "linecolor": COLOR_PALETTE["text_dark"],
}

YAXIS_FREQUENCY_SIMULATED = {
    "title": {
        "text": "Frecuencia Simulada (Hz)",
        "font": {"color": COLOR_PALETTE["freq_simulated"]},
    },
    "tickfont": {"color": COLOR_PALETTE["freq_simulated"]},
    "showgrid": True,
    "gridwidth": 1,
    "gridcolor": COLOR_PALETTE["grid"],
}

YAXIS_POWER = {
    "title": {
        "text": "Potencia (MW)",
        "font": {"color": COLOR_PALETTE["power_real"]},
    },
    "tickfont": {"color": COLOR_PALETTE["power_real"]},
    "overlaying": "y",
    "side": "right",
    "showgrid": False,
    "zeroline": False,
    "showline": True,
    "linewidth": 1,
    "linecolor": COLOR_PALETTE["text_dark"],
}

YAXIS_POWER_SIMULATED = {
    "title": {
        "text": "Potencia Simulada (MW)",
        "font": {"color": COLOR_PALETTE["power_simulated"]},
    },
    "tickfont": {"color": COLOR_PALETTE["power_simulated"]},
    "overlaying": "y",
    "side": "right",
    "showgrid": False,
}

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE LEYENDA
# ─────────────────────────────────────────────────────────────────────────────
LEGEND_CONFIGS = {
    "bottom_center": {
        "orientation": "h",
        "yanchor": "top",
        "y": -0.12,
        "font": {"size": 10},
        "xanchor": "center",
        "x": 0.5,
        "bgcolor": "rgba(255, 255, 255, 0.8)",
        "bordercolor": COLOR_PALETTE["text_light"],
        "borderwidth": 1,
    },
    
    "top_left": {
        "orientation": "v",
        "yanchor": "top",
        "y": 0.99,
        "xanchor": "left",
        "x": 0.01,
        "bgcolor": "rgba(255, 255, 255, 0.8)",
        "bordercolor": COLOR_PALETTE["text_light"],
        "borderwidth": 1,
    },
    
    "top_right": {
        "orientation": "v",
        "yanchor": "top",
        "y": 0.99,
        "xanchor": "right",
        "x": 0.99,
        "bgcolor": "rgba(255, 255, 255, 0.8)",
        "bordercolor": COLOR_PALETTE["text_light"],
        "borderwidth": 1,
    },
    
    "outside": {
        "orientation": "v",
        "yanchor": "middle",
        "y": 0.5,
        "xanchor": "left",
        "x": 1.05,
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE ANOTACIONES (Líneas de referencia, text)
# ─────────────────────────────────────────────────────────────────────────────
ANNOTATION_STYLES = {
    "fault_time": {
        "line_dash": "solid",
        "line_color": COLOR_PALETTE["fault_line"],
        "line_width": 1.5,
        "annotation_text": "t₀",
        "annotation_position": "top right",
        "annotation_font_color": COLOR_PALETTE["fault_line"],
    },
    
    "eval_time": {
        "line_dash": "solid",
        "line_color": COLOR_PALETTE["eval_line"],
        "line_width": 1.5,
        "annotation_position": "top left",
        "annotation_font_color": COLOR_PALETTE["eval_line"],
    },
    
    "deadband_high": {
        "line_dash": "dash",
        "line_color": COLOR_PALETTE["deadband_line"],
        "line_width": 1,
        "opacity": 0.5,
    },
    
    "deadband_low": {
        "line_dash": "dash",
        "line_color": COLOR_PALETTE["deadband_line"],
        "line_width": 1,
        "opacity": 0.5,
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# DICCIONARIO DE CONFIGURACIÓN GLOBAL (para st.session_state)
# ─────────────────────────────────────────────────────────────────────────────
DEFAULT_GRAPH_CONFIG = {
    # Colores - Datos Reales
    "freq_color_real": COLOR_PALETTE["freq_real"],
    "pot_color_real": COLOR_PALETTE["power_real"],
    
    # Colores - Simulaciones (escenarios E.0 y E.1)
    "freq_color_simulated": COLOR_PALETTE["freq_simulated"],
    "pot_color_simulated": COLOR_PALETTE["power_simulated"],
    "freq_color_sim0": "#1565C0",      # Azul oscuro - Escenario 0
    "pot_color_sim0": "#C62828",       # Rojo oscuro - Escenario 0
    "freq_color_sim1": "#29B6F6",      # Azul claro - Escenario 1
    "pot_color_sim1": "#E64A19",       # Naranja - Escenario 1
    
    # Estilos de línea
    "line_width": LINE_WIDTHS["normal"],
    "line_style_real": LINE_STYLES["solid"],
    "line_style_sim":  LINE_STYLES["solid"],
    
    # Marcadores KPI
    "marker_size": MARKER_SIZES["normal"],
    "show_initial": True,       # ○ f₀, P₀
    "show_nadir": True,          # × f_min
    "show_dt_eval": True,        # ● f_Δt, P_Δt
    "show_pmax_marker": True,    # × P_máxima (post-nadir hasta t₀+Δt)
    
    # Líneas de referencia
    "show_deadband": True,
    "show_fault_line": True,
    "show_eval_line": True,
    
    # Layout
    "template": LAYOUT_PRESETS["default"]["template"],
    "plot_height": LAYOUT_PRESETS["default"]["height"],
    "legend_position": "bottom_center",
    
    # Grid
    "show_grid": True,
}
