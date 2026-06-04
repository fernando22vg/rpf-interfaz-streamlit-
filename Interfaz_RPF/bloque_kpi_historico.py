"""
bloque_kpi_historico.py — Bloque 06: Análisis Histórico RPF
Conecta a PostgreSQL del servidor (rpf_intelligence.rpf_kpi_cobee)
y genera 5 gráficas interactivas con Plotly.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# ─── Conexión PostgreSQL ──────────────────────────────────────────────────────

@st.cache_data(ttl=300, show_spinner=False)
def _load_data() -> pd.DataFrame:
    """Carga todos los KPIs desde el servidor. Cachea 5 minutos."""
    try:
        import psycopg2
        s = st.secrets.get("postgres", {})
        conn = psycopg2.connect(
            host=s.get("host", "192.168.0.92"),
            port=int(s.get("port", 5432)),
            dbname=s.get("dbname", "rpf_intelligence"),
            user=s.get("user", "n8n"),
            password=s.get("password", ""),
            connect_timeout=5,
        )
        df = pd.read_sql("SELECT * FROM rpf_kpi_cobee ORDER BY semestre, evento, unidad", conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"No se pudo conectar a PostgreSQL del servidor: {e}")
        return pd.DataFrame()


def _theme():
    """Colores del tema activo de la interfaz."""
    dark = st.session_state.get("theme") == "dark"
    return {
        "bg":      "#13161f" if dark else "#ffffff",
        "surface": "#1a1e2e" if dark else "#f8fafc",
        "border":  "#252a3d" if dark else "#e2e8f0",
        "text":    "#e2e8f0" if dark else "#1e293b",
        "muted":   "#64748b",
        "grid":    "#1e293b" if dark else "#e2e8f0",
        "paper":   "rgba(0,0,0,0)" ,
        "si":      "#22c55e",
        "no":      "#ef4444",
        "pm":      "#f59e0b",
        "fs":      "#475569",
        "accent":  "#6366f1",
        "accent2": "#38bdf8",
    }


def _base_layout(t: dict, **kwargs) -> dict:
    return dict(
        paper_bgcolor=t["paper"],
        plot_bgcolor=t["paper"],
        font=dict(family="Inter, sans-serif", color=t["muted"], size=11),
        xaxis=dict(gridcolor=t["grid"], zerolinecolor=t["grid"]),
        yaxis=dict(gridcolor=t["grid"], zerolinecolor=t["grid"]),
        margin=dict(t=20, r=16, b=40, l=60),
        **kwargs,
    )


# ─── Filtros ──────────────────────────────────────────────────────────────────

def _render_filters(df: pd.DataFrame):
    col1, col2, col3 = st.columns([2, 2, 2])
    semestres = ["Todos"] + sorted(df["semestre"].dropna().unique().tolist())
    unidades  = ["Todas"] + sorted(df["unidad"].dropna().unique().tolist())

    with col1:
        sel_sem = st.selectbox("Semestre", semestres, key="kpi_sem")
    with col2:
        sel_ev  = st.selectbox("Evento", ["Todos"], key="kpi_ev")
    with col3:
        sel_uni = st.selectbox("Unidad", unidades, key="kpi_uni")

    dff = df.copy()
    if sel_sem != "Todos":
        dff = dff[dff["semestre"] == sel_sem]
        eventos = ["Todos"] + sorted(dff["evento"].dropna().unique().tolist())
        sel_ev  = st.session_state.get("kpi_ev", "Todos")
    if sel_ev != "Todos":
        dff = dff[dff["evento"] == sel_ev]
    if sel_uni != "Todas":
        dff = dff[dff["unidad"] == sel_uni]

    return dff


# ─── Tab 1: Cumplimiento ──────────────────────────────────────────────────────

def _tab_cumplimiento(df: pd.DataFrame, t: dict):
    # ── Heatmap ──────────────────────────────────────────────────────────────
    st.markdown("#### Mapa de Cumplimiento por Unidad y Evento")

    # Construir etiqueta de evento con semestre
    df2 = df.copy()
    df2["ev_label"] = df2["semestre"].str.replace("_", " ") + " " + df2["evento"]

    # Codificar estado
    ESTADO_NUM = {"Sí": 1, "Pot. máx": 0.5, "No": 0, None: -0.5}
    df2["estado_num"] = df2["aporta_rpf"].map(
        lambda x: ESTADO_NUM.get(x, -1)  # -1 = sin datos
    )
    df2["estado_txt"] = df2["aporta_rpf"].fillna("F.S.")
    df2.loc[df2["aporta_rpf"].isna(), "estado_num"] = -0.5

    pivot_num = df2.pivot_table(index="unidad", columns="ev_label",
                                values="estado_num", aggfunc="first")
    pivot_txt = df2.pivot_table(index="unidad", columns="ev_label",
                                values="estado_txt", aggfunc="first")

    # Ordenar unidades por % incumplimiento
    inc_order = (df2[df2["aporta_rpf"] == "No"]
                 .groupby("unidad").size()
                 .reindex(pivot_num.index, fill_value=0)
                 .sort_values(ascending=False).index.tolist())
    pivot_num = pivot_num.reindex(inc_order)
    pivot_txt = pivot_txt.reindex(inc_order)

    colorscale = [
        [0,    "#0f172a"],  # sin datos
        [0.25, "#334155"],  # fuera de servicio
        [0.5,  "#ef4444"],  # No
        [0.75, "#f59e0b"],  # Pot. máx
        [1,    "#22c55e"],  # Sí
    ]

    fig = go.Figure(go.Heatmap(
        z=pivot_num.values,
        x=pivot_num.columns.tolist(),
        y=pivot_num.index.tolist(),
        text=pivot_txt.values,
        texttemplate="%{text}",
        textfont=dict(size=9, color="rgba(255,255,255,0.6)"),
        colorscale=colorscale,
        showscale=False,
        xgap=2, ygap=2,
        hovertemplate="<b>%{y}</b> · %{x}<br>%{text}<extra></extra>",
        zmin=-1, zmax=1,
    ))
    fig.update_layout(
        **_base_layout(t),
        height=420,
        margin=dict(t=10, r=16, b=90, l=70),
        xaxis=dict(tickangle=-40, tickfont=dict(size=9), gridcolor=t["grid"]),
        yaxis=dict(autorange="reversed", tickfont=dict(size=10), gridcolor=t["grid"]),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # Leyenda manual
    leg1, leg2, leg3, leg4, leg5 = st.columns(5)
    for col, color, label in [
        (leg1, "#22c55e", "Sí aportó"),
        (leg2, "#ef4444", "No aportó"),
        (leg3, "#f59e0b", "Pot. máx"),
        (leg4, "#475569", "Fuera de servicio"),
        (leg5, "#0f172a", "Sin datos"),
    ]:
        col.markdown(
            f'<div style="display:flex;align-items:center;gap:6px;font-size:12px">'
            f'<div style="width:12px;height:12px;background:{color};border-radius:2px;'
            f'border:1px solid #334155"></div>{label}</div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Ranking incumplimiento ────────────────────────────────────────────────
    st.markdown("#### Ranking de Incumplimiento RPF por Unidad")

    inc_df = (
        df2[df2["aporta_rpf"].notna()]
        .groupby("unidad")["aporta_rpf"]
        .value_counts(normalize=True)
        .mul(100).round(1)
        .reset_index(name="pct")
    )

    no_df = inc_df[inc_df["aporta_rpf"] == "No"].sort_values("pct", ascending=True)
    colors = ["#ef4444" if v > 50 else "#f59e0b" if v > 25 else "#22c55e"
              for v in no_df["pct"]]

    fig2 = go.Figure(go.Bar(
        x=no_df["pct"], y=no_df["unidad"],
        orientation="h",
        marker=dict(color=colors, line=dict(color="rgba(0,0,0,0.2)", width=1)),
        text=no_df["pct"].map(lambda v: f"{v:.1f}%"),
        textposition="outside",
        textfont=dict(size=10),
        hovertemplate="<b>%{y}</b><br>Incumplimiento: <b>%{x:.1f}%</b><extra></extra>",
    ))
    fig2.add_vline(x=50, line=dict(color="#ef4444", dash="dot", width=1.5))
    fig2.add_annotation(x=51, y=0, text="50%", showarrow=False,
                        font=dict(size=9, color="#ef4444"), xanchor="left")
    fig2.update_layout(
        **_base_layout(t),
        height=380,
        margin=dict(t=10, r=70, b=40, l=70),
        xaxis=dict(title=dict(text="% eventos sin aporte RPF", font=dict(size=11)),
                   range=[0, 90], gridcolor=t["grid"]),
        yaxis=dict(gridcolor=t["grid"]),
    )
    st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})


# ─── Tab 2: Frecuencia ────────────────────────────────────────────────────────

def _tab_frecuencia(df: pd.DataFrame, t: dict):
    col_a, col_b = st.columns(2)

    # ── Timeline f_min ────────────────────────────────────────────────────────
    with col_a:
        st.markdown("#### Evolución de f_min por Evento")
        ev_df = (df.groupby(["semestre", "evento", "fecha_evento"])
                 .agg(f_min=("f_min_hz", "mean"), f_0=("f_0_hz", "mean"))
                 .reset_index()
                 .dropna(subset=["fecha_evento"])
                 .sort_values("fecha_evento"))

        sem_colors = {"2024_sem2": "#6366f1", "2025_sem1": "#38bdf8", "2025_sem2": "#22c55e"}
        fig = go.Figure()
        for sem, grp in ev_df.groupby("semestre"):
            fig.add_trace(go.Scatter(
                x=grp["fecha_evento"], y=grp["f_min"],
                mode="lines+markers", name=sem.replace("_", " "),
                line=dict(color=sem_colors.get(sem, "#94a3b8"), width=2),
                marker=dict(size=7),
                hovertemplate="<b>%{x}</b><br>f_min: <b>%{y:.3f} Hz</b><extra></extra>",
            ))
        fig.add_hline(y=49.5, line=dict(color="#ef4444", dash="dot", width=1.5),
                      annotation_text="Límite CNDC", annotation_font_size=10)
        fig.update_layout(
            **_base_layout(t),
            height=300,
            yaxis=dict(title=dict(text="f_min [Hz]"), range=[49.1, 50.05],
                       gridcolor=t["grid"]),
            xaxis=dict(type="date", gridcolor=t["grid"]),
            legend=dict(orientation="h", x=0, y=1.1, font=dict(size=10)),
            margin=dict(t=30, r=16, b=40, l=60),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # ── Scatter droop ─────────────────────────────────────────────────────────
    with col_b:
        st.markdown("#### Droop Declarado vs Calculado")
        droop_df = (df.dropna(subset=["droop_inf_pct", "droop_calc_pct"])
                    .groupby("unidad")
                    .agg(droop_inf=("droop_inf_pct", "mean"),
                         droop_calc=("droop_calc_pct", "mean"),
                         pct_no=("aporta_rpf", lambda x: (x == "No").mean() * 100))
                    .reset_index())

        colors = droop_df["pct_no"].map(
            lambda v: "#ef4444" if v > 50 else "#f59e0b" if v > 25 else "#22c55e"
        )

        # Zona CDM válida (6–12%)
        fig2 = go.Figure()
        fig2.add_shape(type="rect", x0=6, x1=12, y0=0,
                       y1=droop_df["droop_calc"].max() * 1.15,
                       fillcolor="rgba(99,102,241,0.08)",
                       line=dict(color="rgba(99,102,241,0.3)", width=1))
        fig2.add_trace(go.Scatter(
            x=[5, 14], y=[5, 14], mode="lines",
            line=dict(color="#334155", dash="dash", width=1.5),
            hoverinfo="skip", showlegend=False,
        ))
        fig2.add_trace(go.Scatter(
            x=droop_df["droop_inf"], y=droop_df["droop_calc"],
            mode="markers+text", text=droop_df["unidad"],
            textposition="top center",
            textfont=dict(size=9, color=t["muted"]),
            marker=dict(color=colors, size=10,
                        line=dict(color="rgba(0,0,0,0.3)", width=1)),
            hovertemplate=(
                "<b>%{text}</b><br>"
                "Declarado: %{x:.1f}%<br>"
                "Calculado: %{y:.1f}%<extra></extra>"
            ),
        ))
        fig2.update_layout(
            **_base_layout(t),
            height=300,
            xaxis=dict(title=dict(text="Droop declarado [%]"),
                       range=[4, 14], tickvals=[6, 8, 10, 12],
                       gridcolor=t["grid"]),
            yaxis=dict(title=dict(text="Droop calculado [%]"),
                       gridcolor=t["grid"]),
            annotations=[
                dict(x=9, y=droop_df["droop_calc"].max() * 1.1,
                     text="Zona CDM válida (6–12%)",
                     showarrow=False, font=dict(size=9, color="#6366f1")),
            ],
            margin=dict(t=30, r=16, b=40, l=60),
        )
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})


# ─── Tab 3: Reservas ──────────────────────────────────────────────────────────

def _tab_reservas(df: pd.DataFrame, t: dict):
    st.markdown("#### Reserva Disponible vs Potencia Entregada por Unidad")

    res_df = (df.groupby("unidad")
              .agg(reserva=("r_inicial_mw", "mean"),
                   entregada=("p_entregada_mw", "mean"),
                   p_max=("p_max_mw", "mean"))
              .reset_index()
              .sort_values("reserva", ascending=False))

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Reserva disponible [MW]",
        x=res_df["unidad"], y=res_df["reserva"].clip(lower=0),
        marker=dict(color="rgba(99,102,241,0.75)",
                    line=dict(color="#6366f1", width=1)),
        hovertemplate="<b>%{x}</b><br>Reserva: %{y:.2f} MW<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name="Potencia entregada [MW]",
        x=res_df["unidad"], y=res_df["entregada"],
        marker=dict(color="rgba(34,197,94,0.8)",
                    line=dict(color="#22c55e", width=1)),
        hovertemplate="<b>%{x}</b><br>Entregada: %{y:.2f} MW<extra></extra>",
    ))
    fig.update_layout(
        **_base_layout(t),
        barmode="group", height=360,
        margin=dict(t=10, r=16, b=60, l=55),
        yaxis=dict(title=dict(text="[MW]"), gridcolor=t["grid"]),
        xaxis=dict(tickangle=-40, tickfont=dict(size=9), gridcolor=t["grid"]),
        legend=dict(orientation="h", x=0, y=1.05, font=dict(size=10)),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # ── Estadísticas resumidas ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### Estadísticas por Semestre")

    sem_df = (df.groupby("semestre").agg(
        n_eventos=("evento", "nunique"),
        f_min_prom=("f_min_hz", "mean"),
        f_min_peor=("f_min_hz", "min"),
        reserva_prom=("r_inicial_mw", "mean"),
        pct_cumple=("aporta_rpf", lambda x: (x == "Sí").mean() * 100),
    ).reset_index().sort_values("semestre"))

    sem_df.columns = ["Semestre", "Eventos", "f_min prom [Hz]",
                      "f_min peor [Hz]", "Reserva prom [MW]", "% Cumple RPF"]
    sem_df = sem_df.round(3)
    sem_df["% Cumple RPF"] = sem_df["% Cumple RPF"].map(lambda v: f"{v:.1f}%")

    st.dataframe(
        sem_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "f_min peor [Hz]": st.column_config.NumberColumn(format="%.3f"),
            "f_min prom [Hz]": st.column_config.NumberColumn(format="%.3f"),
        },
    )


# ─── Render principal ─────────────────────────────────────────────────────────

def render_bloque_kpi(session_state):
    t = _theme()

    # Cargar datos
    with st.spinner("Cargando datos históricos RPF desde el servidor..."):
        df = _load_data()

    if df.empty:
        st.warning(
            "No hay datos disponibles. Verifica la conexión al servidor "
            "(192.168.0.92:5432) y que la tabla `rpf_kpi_cobee` tenga registros.",
            icon="⚠️",
        )
        st.code(
            "# Agregar en .streamlit/secrets.toml:\n"
            "[postgres]\n"
            'host = "192.168.0.92"\n'
            "port = 5432\n"
            'dbname = "rpf_intelligence"\n'
            'user = "n8n"\n'
            'password = "..."',
            language="toml",
        )
        return

    # Métricas rápidas
    total_ev  = df.groupby(["semestre", "evento"]).ngroups
    pct_si    = (df["aporta_rpf"] == "Sí").mean() * 100
    f_min_min = df["f_min_hz"].min()
    n_units   = df["unidad"].nunique()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Eventos analizados", f"{total_ev}", f"{df['semestre'].nunique()} semestres")
    m2.metric("Cumplimiento RPF", f"{pct_si:.1f}%")
    m3.metric("f_min histórico", f"{f_min_min:.3f} Hz", "↓ peor nadir")
    m4.metric("Unidades COBEE", f"{n_units}")

    st.markdown("---")

    # Filtros
    col1, col2, col3 = st.columns([2, 2, 2])
    semestres = ["Todos"] + sorted(df["semestre"].dropna().unique().tolist())
    unidades  = ["Todas"] + sorted(df["unidad"].dropna().unique().tolist())

    with col1:
        sel_sem = st.selectbox("Semestre", semestres, key="kpi_h_sem")
    with col2:
        eventos_disp = (["Todos"] + sorted(
            df[df["semestre"] == sel_sem]["evento"].dropna().unique().tolist())
            if sel_sem != "Todos" else ["Todos"])
        sel_ev = st.selectbox("Evento", eventos_disp, key="kpi_h_ev")
    with col3:
        sel_uni = st.selectbox("Unidad", unidades, key="kpi_h_uni")

    dff = df.copy()
    if sel_sem != "Todos":
        dff = dff[dff["semestre"] == sel_sem]
    if sel_ev != "Todos":
        dff = dff[dff["evento"] == sel_ev]
    if sel_uni != "Todas":
        dff = dff[dff["unidad"] == sel_uni]

    if dff.empty:
        st.info("Sin registros para los filtros seleccionados.")
        return

    # Tabs
    tab1, tab2, tab3 = st.tabs(["Cumplimiento", "Frecuencia & Droop", "Reservas"])
    with tab1:
        _tab_cumplimiento(dff, t)
    with tab2:
        _tab_frecuencia(dff, t)
    with tab3:
        _tab_reservas(dff, t)
