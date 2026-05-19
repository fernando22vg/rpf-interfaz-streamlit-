# -*- coding: utf-8 -*-
"""
visualizar_red_pf.py

Aplicacion Streamlit para visualizar la topologia de la red electrica
exportada por TopologiaCompleta_PF.py (archivo topologia_grafo.json).

Uso:
    streamlit run visualizar_red_pf.py

Dependencias:
    pip install streamlit networkx plotly pandas
"""

import json
from pathlib import Path

import networkx as nx
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# =============================================================================
# CONFIGURACION
# =============================================================================
DEFAULT_JSON = r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Topologia\topologia_grafo.json"

# Colores por nivel de tension — mismos que loc_namesLineas.py
_COLOR_KV = {
    500: "#D9B3FF",
    230: "#BDD7EE",
    115: "#C5E0B4",
    69:  "#FFE699",
    44:  "#FCE4D6",
    25:  "#F4CCCC",
    24:  "#F4CCCC",
}
# Colores de aristas por tipo de elemento
_COLOR_EDGE = {
    "Linea":              "#2E75B6",
    "Transformador_2dev": "#ED7D31",
    "Transformador_3dev": "#FFC000",
    "Interruptor":        "#808080",
    "Compensador":        "#70AD47",
    "Otro":               "#BFBFBF",
}
_DASH_EDGE = {
    "Transformador_2dev": "dot",
    "Transformador_3dev": "dot",
}

# =============================================================================
# HELPERS
# =============================================================================
def _kv_color(kv) -> str:
    if kv is None:
        return "#EDEDED"
    try:
        kv = float(kv)
    except Exception:
        return "#EDEDED"
    for nivel in sorted(_COLOR_KV.keys(), reverse=True):
        if kv >= nivel * 0.9:
            return _COLOR_KV[nivel]
    return "#EDEDED"


def _kv_label(kv) -> str:
    if kv is None:
        return "kV ?"
    try:
        return f"{int(round(float(kv)))} kV"
    except Exception:
        return "kV ?"


def _kv_sort(label: str) -> float:
    try:
        return float(label.split()[0])
    except Exception:
        return -1.0


def _fmt(v, dec: int = 2, unit: str = "") -> str:
    if v is None:
        return "?"
    try:
        s = f"{float(v):.{dec}f}"
        return f"{s} {unit}" if unit else s
    except Exception:
        return str(v)


# =============================================================================
# CARGA DE DATOS
# =============================================================================
@st.cache_data(show_spinner="Cargando JSON...")
def _load_from_path(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


@st.cache_data(show_spinner="Cargando JSON...")
def _load_from_bytes(content: bytes) -> dict:
    return json.loads(content)


# =============================================================================
# FILTROS
# =============================================================================
def _apply_filters(nodes: list, edges: list, f: dict):
    zone_set = set(f["zones"]) if f["zones"] else None
    tipo_set = set(f["tipos"]) if f["tipos"] else None
    kv_lo, kv_hi = f["kv_range"]
    comp_set = set(str(c) for c in f["comps"]) if f["comps"] else None
    only_serv = f["in_service"]

    accepted: set = set()
    fn = []
    for n in nodes:
        kv = n.get("kv")
        kv_ok   = (kv is None) or (kv_lo <= float(kv) <= kv_hi)
        zone_ok = (zone_set is None) or ((n.get("zone") or "Sin zona") in zone_set)
        serv_ok = (not only_serv) or (n.get("en_servicio") == "Si")
        comp_ok = (comp_set is None) or (str(n.get("componente", "")) in comp_set)
        if kv_ok and zone_ok and serv_ok and comp_ok:
            fn.append(n)
            accepted.add(n["id"])

    fe = []
    for e in edges:
        if e.get("source") not in accepted or e.get("target") not in accepted:
            continue
        tipo_ok = (tipo_set is None) or (e.get("tipo", "Otro") in tipo_set)
        serv_ok = (not only_serv) or (e.get("en_servicio") == "Si")
        if tipo_ok and serv_ok:
            fe.append(e)

    return fn, fe


# =============================================================================
# GRAFO NETWORKX
# =============================================================================
def _build_graph(nodes: list, edges: list) -> nx.Graph:
    G = nx.Graph()
    for n in nodes:
        G.add_node(n["id"], **n)
    for e in edges:
        s, t = e.get("source", ""), e.get("target", "")
        if G.has_node(s) and G.has_node(t):
            G.add_edge(s, t, **e)
    return G


def _compute_positions(G: nx.Graph, method: str) -> dict:
    n = len(G.nodes())
    if n == 0:
        return {}
    if method == "spring":
        k = 2.5 / max(n ** 0.5, 1)
        return nx.spring_layout(G, k=k, iterations=60, seed=42)
    if method == "kamada_kawai":
        try:
            return nx.kamada_kawai_layout(G)
        except Exception:
            return nx.spring_layout(G, seed=42)
    if method == "spectral":
        try:
            return nx.spectral_layout(G)
        except Exception:
            return nx.spring_layout(G, seed=42)
    if method == "shell (por zona)":
        zones: dict = {}
        for nd, d in G.nodes(data=True):
            z = d.get("zone") or "Sin zona"
            zones.setdefault(z, []).append(nd)
        shells = list(zones.values())
        if len(shells) > 1:
            try:
                return nx.shell_layout(G, nlist=shells)
            except Exception:
                pass
        return nx.spring_layout(G, seed=42)
    return nx.spring_layout(G, seed=42)


# =============================================================================
# FIGURA PLOTLY
# =============================================================================
def _make_figure(G: nx.Graph, pos: dict, show_labels: bool) -> go.Figure:
    traces = []

    # ── Aristas por tipo ──────────────────────────────────────────────────────
    edge_groups: dict = {}
    for u, v, d in G.edges(data=True):
        tipo = d.get("tipo", "Otro")
        edge_groups.setdefault(tipo, {"x": [], "y": [], "hover": []})
        x0, y0 = pos.get(u, (0.0, 0.0))
        x1, y1 = pos.get(v, (0.0, 0.0))
        loading = d.get("loading_pct")
        hover = (
            f"<b>{d.get('element', tipo)}</b><br>"
            f"Tipo: {tipo} | {d.get('clase_pf', '?')}<br>"
            f"Carga: {_fmt(loading, 1, '%')}<br>"
            f"P_from: {_fmt(d.get('P_from_MW'), 2, 'MW')} &nbsp; "
            f"P_to: {_fmt(d.get('P_to_MW'), 2, 'MW')}<br>"
            f"Perdidas: {_fmt(d.get('perdidas_MW'), 3, 'MW')}<br>"
            f"kV nom: {_fmt(d.get('kv_nom'), 1, 'kV')} &nbsp; "
            f"MVA nom: {_fmt(d.get('mva_nom'), 1, 'MVA')}<br>"
            f"Km: {_fmt(d.get('km'), 2, 'km')}<br>"
            f"En servicio: {d.get('en_servicio', '?')}"
        )
        edge_groups[tipo]["x"] += [x0, x1, None]
        edge_groups[tipo]["y"] += [y0, y1, None]
        edge_groups[tipo]["hover"] += [hover, hover, None]

    is_first_edge = True
    for tipo, gd in edge_groups.items():
        color = _COLOR_EDGE.get(tipo, "#BFBFBF")
        dash  = _DASH_EDGE.get(tipo, "solid")
        traces.append(go.Scatter(
            x=gd["x"], y=gd["y"],
            mode="lines",
            name=tipo,
            legendgroup="aristas",
            legendgrouptitle_text="Aristas" if is_first_edge else None,
            line=dict(color=color, width=1.8, dash=dash),
            hovertext=gd["hover"],
            hoverinfo="text",
        ))
        is_first_edge = False

    # ── Nodos agrupados por kV ────────────────────────────────────────────────
    kv_groups: dict = {}
    for nd, d in G.nodes(data=True):
        kv    = d.get("kv")
        label = _kv_label(kv)
        color = d.get("kv_color") or _kv_color(kv)
        kv_groups.setdefault(label, {
            "x": [], "y": [], "text": [], "hover": [], "size": [], "color": color,
        })
        x, y = pos.get(nd, (0.0, 0.0))
        deg   = G.degree(nd)
        size  = 9 + min(deg * 1.2, 14)
        attached = d.get("attached") or {}
        hover = (
            f"<b>{nd}</b><br>"
            f"kV nom: {_fmt(kv, 1, 'kV')}<br>"
            f"Zona: {d.get('zone') or '?'}<br>"
            f"V(pu): {_fmt(d.get('v_pu'))}<br>"
            f"V(kV): {_fmt(d.get('v_kv'), 2, 'kV')}<br>"
            f"Angulo: {_fmt(d.get('angle_deg'), 2, '°')}<br>"
            f"P: {_fmt(d.get('P_MW'), 2, 'MW')} &nbsp; "
            f"Q: {_fmt(d.get('Q_Mvar'), 2, 'Mvar')}<br>"
            f"Componente: {d.get('componente', '?')}<br>"
            f"Aislada: {'SI ⚠' if d.get('aislada') == 'Si' else 'No'}<br>"
            f"En servicio: {d.get('en_servicio', '?')}<br>"
            f"Vecinos: {deg}"
        )
        for cat_key, cat_label in (
            ("cargas", "Cargas"),
            ("generadores", "Generadores"),
            ("shunts", "Shunts"),
            ("compensadores", "Compensadores"),
        ):
            items = attached.get(cat_key, [])
            if items:
                preview = ", ".join(items[:3]) + ("…" if len(items) > 3 else "")
                hover += f"<br>{cat_label} ({len(items)}): {preview}"

        kv_groups[label]["x"].append(x)
        kv_groups[label]["y"].append(y)
        kv_groups[label]["text"].append(nd)
        kv_groups[label]["hover"].append(hover)
        kv_groups[label]["size"].append(size)

    is_first_node = True
    for label, gd in sorted(kv_groups.items(), key=lambda x: -_kv_sort(x[0])):
        traces.append(go.Scatter(
            x=gd["x"], y=gd["y"],
            mode="markers+text" if show_labels else "markers",
            name=label,
            legendgroup="nodos",
            legendgrouptitle_text="Barras" if is_first_node else None,
            marker=dict(
                color=gd["color"],
                size=gd["size"],
                symbol="circle",
                line=dict(color="#333", width=0.8),
            ),
            text=gd["text"] if show_labels else None,
            textposition="top center",
            textfont=dict(size=7, color="#333"),
            hovertext=gd["hover"],
            hoverinfo="text",
        ))
        is_first_node = False

    fig = go.Figure(
        data=traces,
        layout=go.Layout(
            showlegend=True,
            hovermode="closest",
            xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
            yaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
            height=680,
            margin=dict(t=20, b=10, l=10, r=180),
            paper_bgcolor="white",
            plot_bgcolor="#F8F9FA",
            legend=dict(
                x=1.01, y=1,
                bgcolor="white",
                bordercolor="#DDD",
                borderwidth=1,
                font=dict(size=10),
                groupclick="toggleitem",
            ),
        ),
    )
    return fig


# =============================================================================
# LEYENDA DE COLORES
# =============================================================================
def _render_legend():
    items = sorted(_COLOR_KV.items(), reverse=True)
    cols = st.columns(len(items) + 1)
    for i, (kv, hex_color) in enumerate(items):
        with cols[i]:
            st.markdown(
                f"<div style='background:{hex_color};border:1px solid #999;"
                f"border-radius:4px;padding:4px 8px;text-align:center;"
                f"font-size:12px'>{kv} kV</div>",
                unsafe_allow_html=True,
            )
    with cols[-1]:
        st.markdown(
            "<div style='background:#EDEDED;border:1px solid #999;"
            "border-radius:4px;padding:4px 8px;text-align:center;"
            "font-size:12px'>? kV</div>",
            unsafe_allow_html=True,
        )


# =============================================================================
# PANEL DE METRICAS
# =============================================================================
def _metrics_row(nodes: list, edges: list, meta: dict):
    n_barras = len(nodes)
    n_serv   = sum(1 for n in nodes if n.get("en_servicio") == "Si")
    n_ais    = sum(1 for n in nodes if n.get("aislada") == "Si")
    n_edges  = len(edges)
    n_lne    = sum(1 for e in edges if e.get("tipo") == "Linea")
    n_xfo    = sum(1 for e in edges if "Transformador" in (e.get("tipo") or ""))

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Barras",          n_barras)
    c2.metric("En servicio",     n_serv)
    c3.metric("Aisladas",        n_ais,    delta=None if n_ais == 0 else f"⚠ {n_ais}", delta_color="inverse")
    c4.metric("Conexiones",      n_edges)
    c5.metric("Lineas",          n_lne)
    c6.metric("Transformadores", n_xfo)

    if meta:
        st.caption(
            f"Proyecto PF: **{meta.get('proyecto', '?')}** &nbsp;|&nbsp; "
            f"Extraccion: {meta.get('timestamp', '?')} &nbsp;|&nbsp; "
            f"Total barras modelo: {meta.get('n_barras', '?')} &nbsp;|&nbsp; "
            f"Total elementos: {meta.get('n_elementos', '?')}"
        )


# =============================================================================
# SIDEBAR — carga y filtros
# =============================================================================
def _sidebar(nodes: list, edges: list):
    with st.sidebar:
        st.header("Fuente de datos")
        src = st.radio("", ["Ruta de archivo", "Subir archivo"], horizontal=True, label_visibility="collapsed")

        data = None
        if src == "Ruta de archivo":
            path = st.text_input("Ruta JSON", value=DEFAULT_JSON, label_visibility="collapsed")
            if st.button("Cargar archivo", use_container_width=True):
                if Path(path).exists():
                    st.session_state["data"]      = _load_from_path(path)
                    st.session_state.pop("pos", None)
                else:
                    st.error("Archivo no encontrado.")
        else:
            uploaded = st.file_uploader("Seleccionar .json", type="json", label_visibility="collapsed")
            if uploaded:
                content = uploaded.read()
                st.session_state["data"]  = _load_from_bytes(content)
                st.session_state.pop("pos", None)

        data = st.session_state.get("data")

        if data is None:
            return None, {}, "spring", False

        nodes_all = data.get("nodes", [])
        edges_all = data.get("edges", [])

        st.divider()
        st.header("Filtros")

        all_zones = sorted({n.get("zone") or "Sin zona" for n in nodes_all})
        sel_zones = st.multiselect("Zonas", all_zones, default=all_zones)

        kvs = sorted(filter(None, {n.get("kv") for n in nodes_all}))
        if kvs:
            kv_lo, kv_hi = float(min(kvs)), float(max(kvs))
            kv_range = st.slider("Rango kV nom.", kv_lo, kv_hi, (kv_lo, kv_hi))
        else:
            kv_range = (0.0, 9999.0)

        all_tipos = sorted({e.get("tipo", "Otro") for e in edges_all})
        sel_tipos = st.multiselect("Tipos de conexion", all_tipos, default=all_tipos)

        in_serv = st.checkbox("Solo en servicio", value=True)

        all_comps = sorted({str(n.get("componente")) for n in nodes_all if n.get("componente") is not None})
        if len(all_comps) > 1:
            sel_comps = st.multiselect("Componentes", all_comps, default=all_comps)
        else:
            sel_comps = all_comps

        st.divider()
        st.header("Visualizacion")
        layout_method = st.selectbox(
            "Algoritmo de posicion",
            ["spring", "kamada_kawai", "spectral", "shell (por zona)"],
        )
        show_labels = st.checkbox("Mostrar etiquetas de barras", value=len(nodes_all) <= 60)
        if st.button("Recalcular posiciones", use_container_width=True):
            st.session_state.pop("pos", None)
            st.session_state.pop("pos_key", None)

    filters = {
        "zones":      sel_zones or None,
        "kv_range":   kv_range,
        "tipos":      sel_tipos or None,
        "in_service": in_serv,
        "comps":      sel_comps or None,
    }
    return data, filters, layout_method, show_labels


# =============================================================================
# MAIN
# =============================================================================
def main():
    st.set_page_config(
        page_title="Topologia Red PF",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.title("⚡ Visualizador de Topologia — DIgSILENT PowerFactory")

    # Placeholder para los filtros (necesitamos data primero)
    data, filters, layout_method, show_labels = _sidebar([], [])

    if data is None:
        st.info(
            "Cargue el archivo **topologia_grafo.json** generado por "
            "`TopologiaCompleta_PF.py` usando la barra lateral."
        )
        with st.expander("Instrucciones"):
            st.markdown(
                "1. Ejecutar `TopologiaCompleta_PF.py` con PowerFactory abierto.\n"
                "2. El script genera `topologia_grafo.json` en el directorio de salida configurado.\n"
                "3. Ingresar la ruta en la barra lateral y hacer clic en **Cargar archivo**.\n\n"
                "**Dependencias:**\n"
                "```\npip install streamlit networkx plotly pandas\n```\n\n"
                "**Ejecutar:**\n"
                "```\nstreamlit run visualizar_red_pf.py\n```"
            )
        st.stop()

    nodes_all = data.get("nodes", [])
    edges_all = data.get("edges", [])
    meta      = data.get("metadata", {})

    # Filtrar
    fn, fe = _apply_filters(nodes_all, edges_all, filters)
    if not fn:
        st.warning("Ningun nodo cumple los filtros actuales. Ajuste los filtros en la barra lateral.")
        st.stop()

    # Grafo NetworkX
    G = _build_graph(fn, fe)

    # Posiciones — cache en session_state para evitar recalculos
    pos_key = f"{layout_method}_{len(fn)}_{len(fe)}"
    if st.session_state.get("pos_key") != pos_key or "pos" not in st.session_state:
        with st.spinner("Calculando posiciones..."):
            st.session_state["pos"]     = _compute_positions(G, layout_method)
            st.session_state["pos_key"] = pos_key
    pos = st.session_state["pos"]

    # ── Tabs ─────────────────────────────────────────────────────────────────
    tab_graf, tab_barras, tab_elem, tab_resumen = st.tabs(
        ["Grafico de Red", "Barras", "Elementos / Conexiones", "Resumen"]
    )

    with tab_graf:
        _metrics_row(fn, fe, meta)
        st.divider()
        _render_legend()
        st.divider()
        fig = _make_figure(G, pos, show_labels)
        st.plotly_chart(fig, use_container_width=True)
        st.caption(
            f"Mostrando {len(fn)} barras y {len(fe)} conexiones "
            f"(de {len(nodes_all)} / {len(edges_all)} totales en el modelo)."
        )

    with tab_barras:
        df_n = pd.DataFrame(fn)
        if not df_n.empty:
            preferred = ["id", "zone", "kv", "v_pu", "v_kv", "angle_deg",
                         "P_MW", "Q_Mvar", "en_servicio", "componente", "aislada"]
            cols = [c for c in preferred if c in df_n.columns]
            cols += [c for c in df_n.columns if c not in cols and c not in ("kv_color", "attached")]
            st.dataframe(df_n[cols], use_container_width=True, height=520)
            st.download_button(
                "Descargar CSV",
                df_n[cols].to_csv(index=False, encoding="utf-8-sig"),
                "barras_filtradas.csv", "text/csv",
            )

    with tab_elem:
        df_e = pd.DataFrame(fe)
        if not df_e.empty:
            preferred = ["element", "tipo", "clase_pf", "source", "target",
                         "en_servicio", "loading_pct", "P_from_MW", "P_to_MW",
                         "perdidas_MW", "kv_nom", "mva_nom", "km",
                         "uktr_pct", "curmg_pct", "pfe_kW", "pcutr_kW"]
            cols = [c for c in preferred if c in df_e.columns]
            cols += [c for c in df_e.columns if c not in cols]
            st.dataframe(df_e[cols], use_container_width=True, height=520)
            st.download_button(
                "Descargar CSV",
                df_e[cols].to_csv(index=False, encoding="utf-8-sig"),
                "conexiones_filtradas.csv", "text/csv",
            )

    with tab_resumen:
        st.subheader("Estadisticas por nivel de tension")
        df_kv = (
            pd.DataFrame(fn)
            .assign(kV=lambda df: df["kv"].apply(_kv_label))
            .groupby("kV", sort=False)
            .agg(
                N_barras=("id", "count"),
                En_servicio=("en_servicio", lambda x: (x == "Si").sum()),
                Aisladas=("aislada", lambda x: (x == "Si").sum()),
            )
            .sort_values("N_barras", ascending=False)
            .reset_index()
        )
        st.dataframe(df_kv, use_container_width=True, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Por zona")
            df_z = (
                pd.DataFrame(fn)
                .assign(Zona=lambda df: df["zone"].fillna("Sin zona"))
                .groupby("Zona")
                .agg(N_barras=("id", "count"))
                .sort_values("N_barras", ascending=False)
                .reset_index()
            )
            st.dataframe(df_z, use_container_width=True, hide_index=True)

        with col2:
            st.subheader("Por tipo de conexion")
            df_t = (
                pd.DataFrame(fe)
                .assign(Tipo=lambda df: df["tipo"].fillna("Otro"))
                .groupby("Tipo")
                .agg(N=("element", "count"))
                .sort_values("N", ascending=False)
                .reset_index()
            )
            st.dataframe(df_t, use_container_width=True, hide_index=True)

        st.subheader("Componentes conectados")
        df_c = (
            pd.DataFrame(fn)
            [["id", "componente"]]
            .dropna(subset=["componente"])
            .groupby("componente")
            .agg(N_barras=("id", "count"))
            .reset_index()
            .rename(columns={"componente": "Componente"})
            .sort_values("N_barras", ascending=False)
            .reset_index(drop=True)
        )
        df_c["Aislada"] = df_c["N_barras"].apply(lambda x: "Si" if x == 1 else "No")
        st.dataframe(df_c, use_container_width=True, hide_index=True)

        if meta:
            st.subheader("Metadata de extraccion")
            st.json(meta)


if __name__ == "__main__":
    main()
