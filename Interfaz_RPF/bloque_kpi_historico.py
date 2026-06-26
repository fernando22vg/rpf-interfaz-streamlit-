"""
bloque_kpi_historico.py — Bloque 06: Análisis Histórico RPF

Fuentes de datos (en orden de prioridad):
  1. PostgreSQL en red local (192.168.0.92) — más actualizado
  2. rpf_kpi_cobee.csv en SharePoint (03_DATOS GEN) — acceso remoto
  3. Cache local C:\\Datos Cobee\\03_DATOS GEN\\rpf_kpi_cobee.csv — sin red
"""

import json
import glob as _glob
import os
import re
import numpy as np
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from kpi_calc import (
    _find_pmax_time, _cndc_kpis,
    _load_pmax_cargado, _get_pmax_from_cargado,
    _load_tech_map, _get_rp_default,
    _robust_col_detect,
)

_SCADA_SUBDIR = "Graficas Registro 1SEG COBEE"
_SIM_CSV      = r"C:\Datos Cobee\03_DATOS GEN\rpf_kpi_sim.csv"
_LOC_PATH_DEFAULT = (
    r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT"
    r"\Designacion de loc_name\loc_names_gen.xlsx"
)

# Ruta local del CSV cache
_CSV_LOCAL = r"C:\Datos Cobee\03_DATOS GEN\rpf_kpi_cobee.csv"
# Ruta en SharePoint (carpeta relativa dentro del share link)
_SP_FOLDER = "04_Interfaz/Datos Cobee/03_DATOS GEN"
_SP_FILE   = "rpf_kpi_cobee.csv"

# Config Plotly: muestra solo botón de descarga como imagen PNG
_CHART_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
    "modeBarButtonsToKeep": ["toImage"],
    "toImageButtonOptions": {
        "format": "png",
        "scale": 2,
        "width": 1400,
        "height": 700,
    },
}


#  Carga con fallback automático 

@st.cache_data(ttl=300, show_spinner=False)
def _load_data() -> tuple[pd.DataFrame, str, list]:
    """
    Intenta cargar desde PostgreSQL → SharePoint → cache local.
    Devuelve (DataFrame, fuente_usada, [errores_diagnostico]).
    """
    _errors: list[str] = []

    # 1. PostgreSQL
    try:
        import psycopg2
        s = st.secrets.get("postgres", {})
        conn = psycopg2.connect(
            host=s.get("host", "192.168.0.92"),
            port=int(s.get("port", 5432)),
            dbname=s.get("dbname", "rpf_intelligence"),
            user=s.get("user", "n8n"),
            password=s.get("password", ""),
            connect_timeout=3,
        )
        df = pd.read_sql(
            "SELECT * FROM rpf_kpi_cobee ORDER BY semestre, evento, unidad", conn
        )
        conn.close()
        if not df.empty:
            os.makedirs(os.path.dirname(_CSV_LOCAL), exist_ok=True)
            df.to_csv(_CSV_LOCAL, index=False)
            try:
                import sharepoint_client as _spc
                import threading
                csv_bytes = df.to_csv(index=False).encode("utf-8")
                def _upload():
                    try:
                        session, site_url, root_path = _spc._get_session()
                        sp_folder = f"{root_path}/{_SP_FOLDER}"
                        _spc._upload_sp_file(session, site_url, sp_folder, _SP_FILE, csv_bytes)
                    except Exception:
                        pass
                threading.Thread(target=_upload, daemon=True).start()
            except Exception:
                pass
            return df, "🟢 PostgreSQL (servidor local)", _errors
    except Exception as e:
        _errors.append(f"PostgreSQL: {e}")

    # 2. SharePoint
    try:
        import sharepoint_client as _spc
        session, site_url, root_path = _spc._get_session()
        _errors.append(f"SP root_path: {root_path}")
        sp_file_path = f"{root_path}/{_SP_FOLDER}/{_SP_FILE}"
        _errors.append(f"SP file path construida: {sp_file_path}")
        # Listar carpetas en root_path para verificar estructura
        try:
            folders = _spc._list_folders(session, site_url, root_path)
            _errors.append(f"Carpetas en root_path: {[f['Name'] for f in folders]}")
            files = _spc._list_files(session, site_url, root_path)
            _errors.append(f"Archivos en root_path: {[f['Name'] for f in files]}")
        except Exception as le:
            _errors.append(f"No se pudo listar root_path: {le}")
        dl_url = (
            f"{site_url}/_api/web"
            f"/GetFileByServerRelativeUrl('{_spc._sp_path(sp_file_path)}')/$value"
        )
        r = session.get(
            dl_url,
            headers={"Accept": "application/json;odata=nometadata"},
            timeout=20,
        )
        _errors.append(f"SharePoint HTTP status: {r.status_code}")
        if r.status_code == 200:
            import io
            df = pd.read_csv(io.BytesIO(r.content))
            if not df.empty:
                try:
                    os.makedirs(os.path.dirname(_CSV_LOCAL), exist_ok=True)
                    df.to_csv(_CSV_LOCAL, index=False)
                except Exception:
                    pass  # ruta Windows no válida en Linux/cloud
                return df, "🔵 SharePoint (nube COBEE)", _errors
            else:
                _errors.append("SharePoint: CSV descargado pero está vacío")
        elif r.status_code == 404:
            _errors.append("SharePoint: archivo no encontrado en la ruta construida")
        else:
            r.raise_for_status()
    except Exception as e:
        _errors.append(f"SharePoint excepción: {e}")

    # 3. Cache local
    if os.path.exists(_CSV_LOCAL):
        try:
            df = pd.read_csv(_CSV_LOCAL)
            mtime = os.path.getmtime(_CSV_LOCAL)
            from datetime import datetime
            fecha = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            return df, f"🟡 Cache local (actualizado {fecha})", _errors
        except Exception as e:
            _errors.append(f"Cache local: {e}")

    return pd.DataFrame(), "", _errors


@st.cache_data(ttl=60, show_spinner=False)
def _load_sim_data() -> pd.DataFrame:
    """Carga rpf_kpi_sim.csv (KPIs P_max guardados desde Bloque 5)."""
    if not os.path.exists(_SIM_CSV):
        return pd.DataFrame()
    try:
        return pd.read_csv(_SIM_CSV)
    except Exception:
        return pd.DataFrame()


def _b6_save_pmax_kpis(semestre: str, evento: str, unidad: str, fuente: str,
                        source_file: str, kpi_pm: dict, p_max: float,
                        t_pmax_al: float, p_pmax: float, rp: float):
    """Guarda KPIs en P_máxima en rpf_kpi_sim.csv con upsert por (semestre, evento, unidad, fuente)."""
    if kpi_pm is None:
        return False, "No hay KPIs en P_máxima disponibles."

    droop_calc_v = kpi_pm.get("droop_calc")
    try:
        if not isinstance(droop_calc_v, (int, float)) or pd.isna(droop_calc_v):
            droop_calc_v = None
    except Exception:
        droop_calc_v = None

    record = {
        "semestre": semestre, "evento": str(evento), "fecha_evento": None,
        "unidad": unidad, "fuente": fuente, "source_file": source_file,
        "p_max_mw": round(float(p_max), 3),
        "p_0_mw": round(float(kpi_pm["p0"]), 3),
        "p_pmax_mw": round(float(p_pmax), 3),
        "dp_mw_pmax": round(float(kpi_pm["dp"]), 3),
        "dp_pct_pmax": round(float(kpi_pm["dp_pct"]), 3),
        "f_0_hz": round(float(kpi_pm["f0"]), 4),
        "f_min_hz": round(float(kpi_pm["f_min"]), 4),
        "f_pmax_hz": round(float(kpi_pm["f_dt"]), 4),
        "t_min_s": round(float(kpi_pm["t_min"]), 1),
        "t_pmax_s": round(float(t_pmax_al), 1),
        "r_inicial_mw": round(float(kpi_pm["r_inic"]), 3),
        "r_inicial_pct": round(float(kpi_pm["r_inic_pct"]), 2),
        "droop_inf_pct": round(float(rp * 100), 1),
        "droop_calc_pmax": round(float(droop_calc_v), 2) if droop_calc_v is not None else None,
        "aporta_pmax": "Sí" if kpi_pm["aporta"] else "No",
    }

    new_df = pd.DataFrame([record])

    if os.path.exists(_SIM_CSV):
        try:
            existing = pd.read_csv(_SIM_CSV)
            mask_keep = ~(
                (existing["semestre"] == semestre) &
                (existing["evento"] == str(evento)) &
                (existing["unidad"] == unidad) &
                (existing["fuente"] == fuente)
            )
            new_df = pd.concat([existing[mask_keep], new_df], ignore_index=True)
        except Exception:
            pass

    try:
        os.makedirs(os.path.dirname(_SIM_CSV), exist_ok=True)
        new_df.to_csv(_SIM_CSV, index=False)
        return True, "KPI guardado correctamente."
    except Exception as exc:
        return False, str(exc)


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
        **kwargs,
    )


def _axis(t: dict, **kwargs) -> dict:
    """Eje con grid del tema + kwargs adicionales."""
    return dict(gridcolor=t["grid"], zerolinecolor=t["grid"], **kwargs)


#  Filtros 

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
        eventos = ["Todos"] + sorted(dff["evento"].dropna().unique().tolist(),
                                       key=lambda e: int(m.group(1)) if (m := re.search(r"(\d+)$", e)) else -1)
        sel_ev  = st.session_state.get("kpi_ev", "Todos")
    if sel_ev != "Todos":
        dff = dff[dff["evento"] == sel_ev]
    if sel_uni != "Todas":
        dff = dff[dff["unidad"] == sel_uni]

    return dff


# ── Helpers para verificación SCADA en P_max ─────────────────────────────────

def _b6_parse_to_seconds(series: pd.Series) -> pd.Series:
    """Convierte columna de tiempo (HH:MM:SS o float) a segundos."""
    s = series.astype(str).str.strip().str.replace(',', '.', regex=False)
    result = pd.Series(0.0, index=series.index)
    has_colon = s.str.contains(':')
    if has_colon.any():
        parts = s[has_colon].str.split(':')
        h   = pd.to_numeric(parts.str[0], errors='coerce').fillna(0)
        m   = pd.to_numeric(parts.str[1], errors='coerce').fillna(0)
        sec = pd.to_numeric(parts.str[2], errors='coerce').fillna(0)
        result[has_colon] = h * 3600 + m * 60 + sec
    result[~has_colon] = pd.to_numeric(s[~has_colon], errors='coerce').fillna(0.0)
    return result


def _b6_load_event_cfg(ev_path: str) -> dict:
    p = os.path.join(ev_path, "event_config.json")
    try:
        if os.path.isfile(p):
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


@st.cache_data(ttl=120, show_spinner=False)
def _b6_load_scada(scada_path: str) -> tuple:
    """Carga SCADA y devuelve (t_al_arr, freq_arr, pot_arr, f_col, p_col)."""
    df = pd.read_excel(scada_path, engine="calamine").dropna(how="all")
    tc, fc_col, pc_col = _robust_col_detect(df)
    t_raw  = _b6_parse_to_seconds(df[tc])
    t_norm = (t_raw - t_raw.min()).values
    freq   = pd.to_numeric(df[fc_col], errors='coerce').ffill().bfill().values
    pot    = pd.to_numeric(df[pc_col], errors='coerce').ffill().bfill().values
    return t_norm, freq, pot, fc_col, pc_col


def _b6_kpi_comparison(kpi: dict, kpi_pm, p_max: float, dt: int,
                        t_pmax: float, t: dict):
    """Tabla HTML con comparación de KPIs estándar vs P_max."""

    def _badge(val: bool) -> str:
        c, lbl = ("#22c55e", "✓ Sí") if val else ("#ef4444", "✗ No")
        return f'<span style="color:{c};font-weight:600">{lbl}</span>'

    rows = [
        ("P_max nominal [MW]",  f"{p_max:.1f}",                f"{p_max:.1f}"),
        ("P₀ [MW]",             f"{kpi['p0']:.3f}",            f"{kpi['p0']:.3f}"),
        ("P evaluada [MW]",     f"{kpi['p_dt']:.3f}",          f"{kpi_pm['p_dt']:.3f}"  if kpi_pm else "—"),
        ("ΔP [MW]",             f"{kpi['dp']:.3f}",            f"{kpi_pm['dp']:.3f}"    if kpi_pm else "—"),
        ("ΔP% [%]",             f"{kpi['dp_pct']:.2f}%",       f"{kpi_pm['dp_pct']:.2f}%" if kpi_pm else "—"),
        ("Aporta RPF",          _badge(kpi["aporta"]),         _badge(kpi_pm["aporta"]) if kpi_pm else "—"),
        ("Punto evaluado",      f"t₀+{dt} s",                  f"t = {t_pmax:.1f} s"    if t_pmax is not None else "—"),
    ]

    html = (
        f'<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:10px">'
        f'<tr style="border-bottom:2px solid {t["border"]}">'
        f'<th style="text-align:left;padding:8px 12px;color:{t["muted"]}">Métrica</th>'
        f'<th style="text-align:center;padding:8px 12px;color:#4682b4">t₀+{dt} s (estándar)</th>'
        f'<th style="text-align:center;padding:8px 12px;color:#ef4444">P_máxima SCADA</th>'
        f'</tr>'
    )
    for i, (label, v_std, v_pm) in enumerate(rows):
        bg = t["surface"] if i % 2 == 0 else t["bg"]
        html += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:7px 12px;color:{t["muted"]}">{label}</td>'
            f'<td style="text-align:center;padding:7px 12px;color:{t["text"]}">{v_std}</td>'
            f'<td style="text-align:center;padding:7px 12px;color:{t["text"]}">{v_pm}</td>'
            f'</tr>'
        )
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)


def _b6_render_scada_pmax(sel_uni: str, t: dict):
    """Sección interactiva: carga SCADA del evento activo y recalcula P_max según Δt."""

    st.markdown("---")
    st.markdown("#### Verificación SCADA en P_máxima")
    st.caption(
        "Carga el archivo SCADA del **evento activo** (cargado en Bloque 1) para la unidad "
        "seleccionada y recalcula dinámicamente el punto de P_máxima según el intervalo Δt configurado."
    )

    ev_path  = st.session_state.get("ev_path_global")
    n_evento = st.session_state.get("n_evento_global")

    if not ev_path or not os.path.isdir(str(ev_path)):
        st.info("ℹ️ Carga un evento en el **Bloque 1** para activar la verificación SCADA.")
        return

    st.caption(f"Evento activo: **{n_evento}** — `{ev_path}`")

    # Buscar archivo SCADA de la unidad
    scada_exact = os.path.join(ev_path, _SCADA_SUBDIR, f"{sel_uni}.xlsx")
    if os.path.isfile(scada_exact):
        scada_path = scada_exact
    else:
        candidates = _glob.glob(
            os.path.join(ev_path, _SCADA_SUBDIR, f"*{sel_uni}*.xlsx")
        )
        if not candidates:
            st.warning(
                f"No se encontró archivo SCADA para **{sel_uni}** en "
                f"`{_SCADA_SUBDIR}/`. Verifica que el evento esté cargado con "
                "los archivos SCADA de esa unidad."
            )
            return
        scada_path = candidates[0]

    try:
        t_norm_arr, freq_arr, pot_arr, fc_col, pc_col = _b6_load_scada(scada_path)
    except Exception as e:
        st.error(f"Error al leer SCADA: {e}")
        return

    # Determinar t₀ desde event_config (compartido entre unidades)
    ev_cfg = _b6_load_event_cfg(ev_path)
    t0_s = ev_cfg.get("scada_t0_s", None)
    if t0_s is not None:
        idx_t0 = int(np.argmin(np.abs(t_norm_arr - float(t0_s))))
    else:
        df_dt = np.gradient(freq_arr)
        drop = np.where(df_dt < -0.02)[0]
        idx_t0 = max(0, int(drop[0]) - 2) if len(drop) > 0 else len(t_norm_arr) // 3

    t_al = t_norm_arr - t_norm_arr[idx_t0]

    # Control Δt (reactive: cada cambio dispara rerun → P_max se recalcula)
    _cc1, _ = st.columns([1, 5])
    with _cc1:
        dt_b6 = st.number_input(
            "Δt CNDC [s]", value=35, min_value=10, max_value=120, step=1,
            key="b6_scada_dt",
            help="Ventana de evaluación CNDC [t₀, t₀+Δt]. "
                 "Al cambiar este valor el punto P_máxima se recalcula automáticamente.",
        )

    # P_max nominal y Rp de la unidad
    loc_path = st.session_state.get("cfg_LOC_NAMES_GEN_PATH", _LOC_PATH_DEFAULT)
    try:
        pmax_cargado = _load_pmax_cargado(ev_path, n_evento)
        tech_map     = _load_tech_map(loc_path) if os.path.isfile(str(loc_path)) else {}
        p_max, tk, _ = _get_pmax_from_cargado(sel_uni, pmax_cargado, tech_map)
    except Exception:
        p_max, tk = 100.0, sel_uni
    rp = _get_rp_default(tk, loc_path) / 100.0 if os.path.isfile(str(loc_path)) else 0.10

    # KPIs estándar (evaluación en t₀+Δt)
    kpi = _cndc_kpis(t_al, freq_arr, pot_arr, p_max, rp, int(dt_b6))
    if kpi is None:
        st.warning("No se pudieron calcular los KPIs (datos insuficientes en el archivo SCADA).")
        return

    # P_max en ventana [t_nadir, t₀+Δt] — se recalcula con cada cambio de dt_b6
    t_nadir   = float(kpi["t_min"])
    t_pmax_al, p_pmax = _find_pmax_time(t_al, pot_arr, int(dt_b6), t_min_eval=t_nadir)

    # KPIs en P_max
    kpi_pm = None
    if t_pmax_al is not None:
        kpi_pm = _cndc_kpis(t_al, freq_arr, pot_arr, p_max, rp, t_pmax_al)

    # ── Gráfico dual-eje ──────────────────────────────────────────────────────
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=t_al, y=freq_arr, name="Frecuencia (Hz)", yaxis="y",
        line=dict(color="#1f77b4", width=2),
        hovertemplate="t=%{x:.1f} s — f=%{y:.4f} Hz<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=t_al, y=pot_arr, name=f"Potencia {sel_uni} (MW)", yaxis="y2",
        line=dict(color="#2ca02c", width=2),
        hovertemplate="t=%{x:.1f} s — P=%{y:.3f} MW<extra></extra>",
    ))

    # Líneas de referencia t₀ y t₀+Δt
    fig.add_vline(x=0, line=dict(color="#94a3b8", width=1.5),
                  annotation_text="t₀", annotation_position="top right",
                  annotation_font=dict(color="#94a3b8", size=11))
    fig.add_vline(x=int(dt_b6), line=dict(color="#4682b4", dash="dash", width=1.5),
                  annotation_text=f"t₀+{dt_b6}s", annotation_position="top left",
                  annotation_font=dict(color="#4682b4", size=11))

    # Marcadores KPI estándar
    fig.add_trace(go.Scatter(
        x=[0], y=[kpi["f0"]], yaxis="y", mode="markers",
        name=f"f₀ = {kpi['f0']:.4f} Hz",
        marker=dict(symbol="circle-open", color="#1f77b4", size=10, line=dict(width=2)),
    ))
    fig.add_trace(go.Scatter(
        x=[kpi["t_min"]], y=[kpi["f_min"]], yaxis="y", mode="markers",
        name=f"f_min = {kpi['f_min']:.4f} Hz",
        marker=dict(symbol="x", color="#ff7f0e", size=12, line=dict(width=2)),
    ))
    fig.add_trace(go.Scatter(
        x=[int(dt_b6)], y=[kpi["f_dt"]], yaxis="y", mode="markers",
        name=f"f_Δt = {kpi['f_dt']:.4f} Hz",
        marker=dict(symbol="circle", color="#4682b4", size=10),
    ))
    fig.add_trace(go.Scatter(
        x=[0], y=[kpi["p0"]], yaxis="y2", mode="markers",
        name=f"P₀ = {kpi['p0']:.3f} MW",
        marker=dict(symbol="circle-open", color="#2ca02c", size=10, line=dict(width=2)),
    ))
    fig.add_trace(go.Scatter(
        x=[int(dt_b6)], y=[kpi["p_dt"]], yaxis="y2", mode="markers",
        name=f"P_Δt = {kpi['p_dt']:.3f} MW",
        marker=dict(symbol="circle", color="#2ca02c", size=10),
    ))

    # Marcador P_max (se mueve con Δt)
    if t_pmax_al is not None:
        idx_pm = int(np.argmin(np.abs(t_al - t_pmax_al)))
        fig.add_trace(go.Scatter(
            x=[t_pmax_al], y=[p_pmax], yaxis="y2", mode="markers",
            name=f"P_max = {p_pmax:.3f} MW @ t={t_pmax_al:.1f}s",
            marker=dict(symbol="x", color="#ef4444", size=14, line=dict(width=2.5)),
        ))
        fig.add_trace(go.Scatter(
            x=[t_pmax_al], y=[float(freq_arr[idx_pm])], yaxis="y", mode="markers",
            name=f"f@P_max = {float(freq_arr[idx_pm]):.4f} Hz",
            marker=dict(symbol="x", color="#ef4444", size=12, line=dict(width=2)),
        ))

    fig.update_layout(
        **_base_layout(t),
        height=450,
        margin=dict(t=20, r=110, b=140, l=65),
        xaxis=dict(title="Tiempo relativo a t₀ [s]", gridcolor=t["grid"]),
        yaxis=dict(title="Frecuencia [Hz]", gridcolor=t["grid"],
                   tickfont=dict(color="#1f77b4")),
        yaxis2=dict(title="Potencia [MW]", overlaying="y", side="right",
                    showgrid=False, tickfont=dict(color="#2ca02c")),
        legend=dict(orientation="h", x=0, y=-0.38, font=dict(size=9),
                    bgcolor="rgba(0,0,0,0)"),
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True, config={
        **_CHART_CONFIG,
        "toImageButtonOptions": {
            **_CHART_CONFIG["toImageButtonOptions"],
            "filename": f"RPF_scada_pmax_{sel_uni}",
        },
    })

    # ── Tabla comparativa KPIs ────────────────────────────────────────────────
    _b6_kpi_comparison(kpi, kpi_pm, p_max, int(dt_b6), t_pmax_al, t)

    # ── Guardar KPIs P_max en histórico ──────────────────────────────────────
    if kpi_pm is not None and t_pmax_al is not None:
        _s_col, _ = st.columns([1, 3])
        with _s_col:
            if st.button(
                "💾 Guardar KPIs P_max SCADA",
                key="b6_save_scada_pmax",
                help=(
                    "Guarda los KPIs calculados en P_máxima en rpf_kpi_sim.csv "
                    "(fuente=SCADA) para acumular el histórico computado visible "
                    "en la sección 'Evaluación P_max computada' del Tab Cumplimiento."
                ),
                use_container_width=True,
            ):
                _semestre = st.session_state.get("semestre_global", "")
                _ok, _msg = _b6_save_pmax_kpis(
                    semestre=_semestre,
                    evento=str(n_evento) if n_evento else "",
                    unidad=sel_uni,
                    fuente="SCADA",
                    source_file=scada_path,
                    kpi_pm=kpi_pm,
                    p_max=p_max,
                    t_pmax_al=t_pmax_al,
                    p_pmax=p_pmax,
                    rp=rp,
                )
                if _ok:
                    _load_sim_data.clear()
                    st.toast("KPIs P_max SCADA guardados correctamente.", icon="✅")
                else:
                    st.error(f"Error al guardar: {_msg}")


def _b6_render_sim_pmax(sel_uni: str, fuente: str, t: dict):
    """Análisis gráfico de P_max para simulación E.0/E.1 con Δt reactivo."""

    st.markdown("---")
    st.markdown(f"#### Verificación en P_máxima — Simulación {fuente}")
    st.caption(
        f"Carga el archivo de simulación **{fuente}** del evento activo y recalcula "
        "dinámicamente el punto de P_máxima según el Δt configurado."
    )

    # Buscar registro en CSV de simulaciones
    if not os.path.exists(_SIM_CSV):
        st.info("ℹ️ No hay datos de simulación guardados. Usa Bloque 5 → 💾 Guardar KPIs P_max en B6.")
        return

    try:
        df_sim_all = pd.read_csv(_SIM_CSV)
    except Exception as e:
        st.error(f"Error al leer {_SIM_CSV}: {e}")
        return

    row = df_sim_all[
        (df_sim_all["unidad"] == sel_uni) & (df_sim_all["fuente"] == fuente)
    ]
    if row.empty:
        st.info(f"No hay datos guardados para **{sel_uni}** · {fuente}. Ejecuta Bloque 5 primero.")
        return

    row = row.iloc[-1]  # más reciente
    source_file = str(row.get("source_file", ""))
    if not source_file or not os.path.isfile(source_file):
        st.warning(f"Archivo de simulación no encontrado: `{source_file}`")
        return

    st.caption(f"Archivo: `{source_file}`")

    try:
        t_norm_arr, freq_arr, pot_arr, fc_col, pc_col = _b6_load_scada(source_file)
    except Exception as e:
        st.error(f"Error al cargar simulación: {e}")
        return

    # t₀ desde event_config (t_sim_falla para simulaciones)
    ev_path = st.session_state.get("ev_path_global", "")
    ev_cfg  = _b6_load_event_cfg(ev_path) if ev_path else {}
    t0_s    = ev_cfg.get("t_sim_falla", None)
    if t0_s is not None:
        idx_t0 = int(np.argmin(np.abs(t_norm_arr - float(t0_s))))
    else:
        df_dt  = np.gradient(freq_arr)
        drop   = np.where(df_dt < -0.02)[0]
        idx_t0 = max(0, int(drop[0]) - 2) if len(drop) > 0 else len(t_norm_arr) // 3

    t_al = t_norm_arr - t_norm_arr[idx_t0]

    _cc1, _ = st.columns([1, 5])
    with _cc1:
        dt_sim = st.number_input(
            "Δt CNDC [s]", value=35, min_value=10, max_value=120, step=1,
            key=f"b6_sim_dt_{fuente.replace('.', '_')}",
            help="Ventana [t₀, t₀+Δt]. Al cambiar, P_máxima se recalcula automáticamente.",
        )

    # P_max nominal y Rp del registro guardado
    p_max = float(row.get("p_max_mw", 100.0))
    rp    = float(row.get("droop_inf_pct", 10.0)) / 100.0

    kpi    = _cndc_kpis(t_al, freq_arr, pot_arr, p_max, rp, int(dt_sim))
    if kpi is None:
        st.warning("No se pudieron calcular los KPIs (datos insuficientes).")
        return

    t_nadir   = float(kpi["t_min"])
    t_pmax_al, p_pmax = _find_pmax_time(t_al, pot_arr, int(dt_sim), t_min_eval=t_nadir)

    kpi_pm = None
    if t_pmax_al is not None:
        kpi_pm = _cndc_kpis(t_al, freq_arr, pot_arr, p_max, rp, t_pmax_al)

    # Colores por fuente
    _c_f = "#29B6F6" if "1" in fuente else "#1565C0"
    _c_p = "#E64A19" if "1" in fuente else "#C62828"

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=t_al, y=freq_arr, name="Frecuencia (Hz)", yaxis="y",
                             line=dict(color=_c_f, width=2),
                             hovertemplate="t=%{x:.1f} s — f=%{y:.4f} Hz<extra></extra>"))
    fig.add_trace(go.Scatter(x=t_al, y=pot_arr, name=f"Potencia {fuente} (MW)", yaxis="y2",
                             line=dict(color=_c_p, width=2),
                             hovertemplate="t=%{x:.1f} s — P=%{y:.3f} MW<extra></extra>"))

    fig.add_vline(x=0, line=dict(color="#94a3b8", width=1.5),
                  annotation_text="t₀", annotation_position="top right",
                  annotation_font=dict(color="#94a3b8", size=11))
    fig.add_vline(x=int(dt_sim), line=dict(color="#4682b4", dash="dash", width=1.5),
                  annotation_text=f"t₀+{dt_sim}s", annotation_position="top left",
                  annotation_font=dict(color="#4682b4", size=11))

    fig.add_trace(go.Scatter(x=[0], y=[kpi["f0"]], yaxis="y", mode="markers",
                             name=f"f₀={kpi['f0']:.4f} Hz",
                             marker=dict(symbol="circle-open", color=_c_f, size=10, line=dict(width=2))))
    fig.add_trace(go.Scatter(x=[kpi["t_min"]], y=[kpi["f_min"]], yaxis="y", mode="markers",
                             name=f"f_min={kpi['f_min']:.4f} Hz",
                             marker=dict(symbol="x", color="#ff7f0e", size=12, line=dict(width=2))))
    fig.add_trace(go.Scatter(x=[int(dt_sim)], y=[kpi["f_dt"]], yaxis="y", mode="markers",
                             name=f"f_Δt={kpi['f_dt']:.4f} Hz",
                             marker=dict(symbol="circle", color="#4682b4", size=10)))
    fig.add_trace(go.Scatter(x=[0], y=[kpi["p0"]], yaxis="y2", mode="markers",
                             name=f"P₀={kpi['p0']:.3f} MW",
                             marker=dict(symbol="circle-open", color=_c_p, size=10, line=dict(width=2))))
    fig.add_trace(go.Scatter(x=[int(dt_sim)], y=[kpi["p_dt"]], yaxis="y2", mode="markers",
                             name=f"P_Δt={kpi['p_dt']:.3f} MW",
                             marker=dict(symbol="circle", color=_c_p, size=10)))

    if t_pmax_al is not None:
        idx_pm = int(np.argmin(np.abs(t_al - t_pmax_al)))
        fig.add_trace(go.Scatter(x=[t_pmax_al], y=[p_pmax], yaxis="y2", mode="markers",
                                 name=f"P_max={p_pmax:.3f} MW @ t={t_pmax_al:.1f}s",
                                 marker=dict(symbol="x", color="#ef4444", size=14, line=dict(width=2.5))))
        fig.add_trace(go.Scatter(x=[t_pmax_al], y=[float(freq_arr[idx_pm])], yaxis="y", mode="markers",
                                 name=f"f@P_max={float(freq_arr[idx_pm]):.4f} Hz",
                                 marker=dict(symbol="x", color="#ef4444", size=12, line=dict(width=2))))

    fig.update_layout(
        **_base_layout(t),
        height=450,
        margin=dict(t=20, r=110, b=140, l=65),
        xaxis=dict(title="Tiempo relativo a t₀ [s]", gridcolor=t["grid"]),
        yaxis=dict(title="Frecuencia [Hz]", gridcolor=t["grid"], tickfont=dict(color=_c_f)),
        yaxis2=dict(title="Potencia [MW]", overlaying="y", side="right",
                    showgrid=False, tickfont=dict(color=_c_p)),
        legend=dict(orientation="h", x=0, y=-0.38, font=dict(size=9), bgcolor="rgba(0,0,0,0)"),
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True, config={
        **_CHART_CONFIG,
        "toImageButtonOptions": {
            **_CHART_CONFIG["toImageButtonOptions"],
            "filename": f"RPF_sim_pmax_{fuente}_{sel_uni}",
        },
    })

    _b6_kpi_comparison(kpi, kpi_pm, p_max, int(dt_sim), t_pmax_al, t)


#  Tab 1: Cumplimiento

def _tab_cumplimiento(df: pd.DataFrame, t: dict, sel_uni: str = "Todas", fuente: str = "SCADA"):
    #  Heatmap 
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

    # Celdas vacías = unidad no despachada en ese evento → F.S.
    pivot_num = pivot_num.fillna(-0.5)
    pivot_txt = pivot_txt.fillna("F.S.")

    # Reordenar columnas numéricamente: "2025 sem2 Evento_10" después de "Evento_9"
    def _ev_col_key(label):
        m = re.match(r"(\d{4})\s+sem(\d+).*?(\d+)$", label)
        return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else (9999, 9, 9999)

    col_order = sorted(pivot_num.columns.tolist(), key=_ev_col_key)
    pivot_num = pivot_num[col_order]
    pivot_txt = pivot_txt[col_order]

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
    st.plotly_chart(fig, use_container_width=True, config={**_CHART_CONFIG,
        "toImageButtonOptions": {**_CHART_CONFIG["toImageButtonOptions"], "filename": "RPF_heatmap_cumplimiento"}})

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

    #  Ranking incumplimiento 
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
    st.plotly_chart(fig2, use_container_width=True, config={**_CHART_CONFIG,
        "toImageButtonOptions": {**_CHART_CONFIG["toImageButtonOptions"], "filename": "RPF_ranking_incumplimiento"}})

    # ── EVALUACIÓN EN P_MÁXIMA ───────────────────────────────────────────────
    st.markdown("---")
    if fuente == "SCADA":
        # Para SCADA no se muestra la evaluación derivada de aporta_rpf.
        # La sección de más abajo muestra los KPIs realmente computados cuando
        # el usuario ha guardado datos con el botón 💾 de la verificación individual.
        st.markdown("#### Cumplimiento evaluado en P_máxima SCADA")
        st.info(
            "Esta sección se activa una vez que se guarden los KPIs P_max desde la "
            "verificación individual. Selecciona una unidad y pulsa "
            "**💾 Guardar KPIs P_max SCADA** para acumular el histórico computado.",
            icon="ℹ️",
        )
    else:
        # Para E.0/E.1 los datos ya son evaluaciones en P_max — se muestran directamente.
        st.markdown(f"#### Cumplimiento evaluado en P_máxima — Simulación {fuente}")
        st.caption(
            f"Para datos de simulación **{fuente}**, los KPIs ya están evaluados "
            "directamente en el punto de P_máxima registrada."
        )

        df_pm = df2.copy()

        df_pm["cumple_pmax"] = df_pm["aporta_rpf"].map(
            lambda x: ("Sí" if x in ("Sí", "Pot. máx") else ("No" if x == "No" else None))
        )
        df_pm["estado_pmax_num"] = df_pm["cumple_pmax"].map(
            {"Sí": 1, "No": 0, None: -0.5}
        ).fillna(-0.5)
        df_pm["estado_pmax_lbl"] = df_pm["cumple_pmax"].fillna("F.S.")

        pivot_pmax_num = df_pm.pivot_table(
            index="unidad", columns="ev_label", values="estado_pmax_num", aggfunc="first"
        ).fillna(-0.5)
        pivot_pmax_txt = df_pm.pivot_table(
            index="unidad", columns="ev_label", values="estado_pmax_lbl", aggfunc="first"
        ).fillna("F.S.")

        _pmax_cols = [c for c in col_order if c in pivot_pmax_num.columns]
        pivot_pmax_num = pivot_pmax_num[_pmax_cols]
        pivot_pmax_txt = pivot_pmax_txt[_pmax_cols]

        _inc_pmax = (df_pm[df_pm["cumple_pmax"] == "No"]
                     .groupby("unidad").size()
                     .reindex(pivot_pmax_num.index, fill_value=0)
                     .sort_values(ascending=False).index.tolist())
        _rest_pmax = [u for u in pivot_pmax_num.index if u not in _inc_pmax]
        pivot_pmax_num = pivot_pmax_num.reindex(_inc_pmax + _rest_pmax)
        pivot_pmax_txt = pivot_pmax_txt.reindex(_inc_pmax + _rest_pmax)

        fig_pmax = go.Figure(go.Heatmap(
            z=pivot_pmax_num.values,
            x=pivot_pmax_num.columns.tolist(),
            y=pivot_pmax_num.index.tolist(),
            text=pivot_pmax_txt.values,
            texttemplate="%{text}",
            textfont=dict(size=9, color="rgba(255,255,255,0.6)"),
            colorscale=[
                [0,    "#0f172a"],
                [0.25, "#334155"],
                [0.5,  "#ef4444"],
                [1,    "#22c55e"],
            ],
            showscale=False, xgap=2, ygap=2,
            hovertemplate="<b>%{y}</b> · %{x}<br>%{text}<extra></extra>",
            zmin=-1, zmax=1,
        ))
        fig_pmax.update_layout(
            **_base_layout(t),
            height=420,
            margin=dict(t=10, r=16, b=90, l=70),
            xaxis=dict(tickangle=-40, tickfont=dict(size=9), gridcolor=t["grid"]),
            yaxis=dict(autorange="reversed", tickfont=dict(size=10), gridcolor=t["grid"]),
        )
        st.plotly_chart(fig_pmax, use_container_width=True, config={**_CHART_CONFIG,
            "toImageButtonOptions": {**_CHART_CONFIG["toImageButtonOptions"],
                                     "filename": "RPF_heatmap_cumplimiento_pmax"}})

        _lp1, _lp2, _lp3, _lp4 = st.columns(4)
        for _col, _color, _lbl in [
            (_lp1, "#22c55e", "Sí aportó en P_máx"),
            (_lp2, "#ef4444", "No aportó en P_máx"),
            (_lp3, "#475569", "Fuera de servicio"),
            (_lp4, "#0f172a", "Sin datos"),
        ]:
            _col.markdown(
                f'<div style="display:flex;align-items:center;gap:6px;font-size:12px">'
                f'<div style="width:12px;height:12px;background:{_color};border-radius:2px;'
                f'border:1px solid #334155"></div>{_lbl}</div>',
                unsafe_allow_html=True,
            )

        st.markdown("---")
        st.markdown("#### Ranking de Incumplimiento RPF por Unidad — evaluado en P_máxima")

        _disp_pm = df_pm[df_pm["cumple_pmax"].notna()]
        inc_pm_df = (
            _disp_pm.groupby("unidad")["cumple_pmax"]
            .value_counts(normalize=True).mul(100).round(1)
            .reset_index(name="pct")
        )
        no_pm_df = inc_pm_df[inc_pm_df["cumple_pmax"] == "No"].sort_values("pct", ascending=True)
        if no_pm_df.empty:
            st.info("No hay unidades con incumplimiento en P_máxima para los filtros seleccionados.")
        else:
            colors_pm = ["#ef4444" if v > 50 else "#f59e0b" if v > 25 else "#22c55e"
                         for v in no_pm_df["pct"]]
            fig_rank_pm = go.Figure(go.Bar(
                x=no_pm_df["pct"], y=no_pm_df["unidad"],
                orientation="h",
                marker=dict(color=colors_pm, line=dict(color="rgba(0,0,0,0.2)", width=1)),
                text=no_pm_df["pct"].map(lambda v: f"{v:.1f}%"),
                textposition="outside", textfont=dict(size=10),
                hovertemplate="<b>%{y}</b><br>No cumple P_max: <b>%{x:.1f}%</b><extra></extra>",
            ))
            fig_rank_pm.add_vline(x=50, line=dict(color="#ef4444", dash="dot", width=1.5))
            fig_rank_pm.add_annotation(x=51, y=0, text="50%", showarrow=False,
                                       font=dict(size=9, color="#ef4444"), xanchor="left")
            fig_rank_pm.update_layout(
                **_base_layout(t),
                height=380,
                margin=dict(t=10, r=70, b=40, l=70),
                xaxis=dict(title=dict(text="% eventos sin aporte RPF en P_máxima", font=dict(size=11)),
                           range=[0, 90], gridcolor=t["grid"]),
                yaxis=dict(gridcolor=t["grid"]),
            )
            st.plotly_chart(fig_rank_pm, use_container_width=True, config={**_CHART_CONFIG,
                "toImageButtonOptions": {**_CHART_CONFIG["toImageButtonOptions"],
                                         "filename": "RPF_ranking_incumplimiento_pmax"}})

    # ── Evaluación P_max computada desde archivos guardados (fuente=SCADA) ────
    if fuente == "SCADA":
        _df_spm_all = _load_sim_data()
        if not _df_spm_all.empty and "fuente" in _df_spm_all.columns:
            _df_spm = _df_spm_all[_df_spm_all["fuente"] == "SCADA"].copy()
            if not _df_spm.empty:
                # Filtrar para que coincida con los eventos del df actual
                _ev_cur = df2[["semestre", "evento"]].drop_duplicates().copy()
                _ev_cur["evento"] = _ev_cur["evento"].astype(str)
                _df_spm = _df_spm.copy()
                _df_spm["evento"] = _df_spm["evento"].astype(str)
                _df_spm_filt = _df_spm.merge(_ev_cur, on=["semestre", "evento"], how="inner")

                st.markdown("---")
                st.markdown("#### Evaluación en P_máxima — Datos computados desde archivos SCADA")
                st.caption(
                    "KPIs calculados directamente **en el punto P_máxima** de cada archivo SCADA registrado, "
                    "usando el botón 💾 de la sección de verificación individual. "
                    "A diferencia del mapa derivado de arriba, aquí los valores provienen del cálculo "
                    "explícito en el instante de mayor potencia medida."
                )

                if _df_spm_filt.empty:
                    st.info(
                        "No hay KPIs P_max computados para los eventos del filtro actual. "
                        "Selecciona una unidad y usa **💾 Guardar KPIs P_max SCADA** para acumular datos."
                    )
                else:
                    _df_spm_filt["ev_label"] = (
                        _df_spm_filt["semestre"].str.replace("_", " ") + " " + _df_spm_filt["evento"]
                    )
                    _df_spm_filt["estado_num"] = _df_spm_filt["aporta_pmax"].map(
                        {"Sí": 1, "No": 0}
                    ).fillna(-0.5)
                    _df_spm_filt["estado_txt"] = _df_spm_filt["aporta_pmax"].fillna("N/D")

                    _piv_spm_n = _df_spm_filt.pivot_table(
                        index="unidad", columns="ev_label", values="estado_num", aggfunc="first"
                    ).fillna(-0.5)
                    _piv_spm_t = _df_spm_filt.pivot_table(
                        index="unidad", columns="ev_label", values="estado_txt", aggfunc="first"
                    ).fillna("N/D")

                    _col_ord_spm = sorted(_piv_spm_n.columns.tolist(), key=_ev_col_key)
                    _piv_spm_n = _piv_spm_n[_col_ord_spm]
                    _piv_spm_t = _piv_spm_t[_col_ord_spm]

                    _inc_spm = (
                        _df_spm_filt[_df_spm_filt["aporta_pmax"] == "No"]
                        .groupby("unidad").size()
                        .reindex(_piv_spm_n.index, fill_value=0)
                        .sort_values(ascending=False).index.tolist()
                    )
                    _rest_spm = [u for u in _piv_spm_n.index if u not in _inc_spm]
                    _piv_spm_n = _piv_spm_n.reindex(_inc_spm + _rest_spm)
                    _piv_spm_t = _piv_spm_t.reindex(_inc_spm + _rest_spm)

                    fig_spm = go.Figure(go.Heatmap(
                        z=_piv_spm_n.values,
                        x=_piv_spm_n.columns.tolist(),
                        y=_piv_spm_n.index.tolist(),
                        text=_piv_spm_t.values,
                        texttemplate="%{text}",
                        textfont=dict(size=9, color="rgba(255,255,255,0.6)"),
                        colorscale=[
                            [0,    "#0f172a"],
                            [0.375, "#334155"],
                            [0.75, "#ef4444"],
                            [1,    "#22c55e"],
                        ],
                        showscale=False,
                        xgap=2, ygap=2,
                        hovertemplate="<b>%{y}</b> · %{x}<br>%{text}<extra></extra>",
                        zmin=-1, zmax=1,
                    ))
                    fig_spm.update_layout(
                        **_base_layout(t),
                        height=max(300, 30 * len(_piv_spm_n) + 60),
                        margin=dict(t=10, r=16, b=90, l=70),
                        xaxis=dict(tickangle=-40, tickfont=dict(size=9), gridcolor=t["grid"]),
                        yaxis=dict(autorange="reversed", tickfont=dict(size=10), gridcolor=t["grid"]),
                    )
                    st.plotly_chart(fig_spm, use_container_width=True, config={**_CHART_CONFIG,
                        "toImageButtonOptions": {**_CHART_CONFIG["toImageButtonOptions"],
                                                 "filename": "RPF_heatmap_pmax_computado_scada"}})

                    # Leyenda
                    _ls1, _ls2, _ls3 = st.columns(3)
                    for _lc, _lcolor, _llbl in [
                        (_ls1, "#22c55e", "Sí aportó en P_máx"),
                        (_ls2, "#ef4444", "No aportó en P_máx"),
                        (_ls3, "#334155", "Sin datos guardados"),
                    ]:
                        _lc.markdown(
                            f'<div style="display:flex;align-items:center;gap:6px;font-size:12px">'
                            f'<div style="width:12px;height:12px;background:{_lcolor};border-radius:2px;'
                            f'border:1px solid #475569"></div>{_llbl}</div>',
                            unsafe_allow_html=True,
                        )

                    # Ranking computado
                    _no_spm = (
                        _df_spm_filt[_df_spm_filt["aporta_pmax"].notna()]
                        .groupby("unidad")["aporta_pmax"]
                        .value_counts(normalize=True).mul(100).round(1)
                        .reset_index(name="pct")
                    )
                    _no_spm = _no_spm[_no_spm["aporta_pmax"] == "No"].sort_values("pct", ascending=True)
                    if not _no_spm.empty:
                        st.markdown("##### Ranking de incumplimiento (P_max computado)")
                        _clrs_spm = [
                            "#ef4444" if v > 50 else "#f59e0b" if v > 25 else "#22c55e"
                            for v in _no_spm["pct"]
                        ]
                        fig_rk_spm = go.Figure(go.Bar(
                            x=_no_spm["pct"], y=_no_spm["unidad"],
                            orientation="h",
                            marker=dict(color=_clrs_spm,
                                        line=dict(color="rgba(0,0,0,0.2)", width=1)),
                            text=_no_spm["pct"].map(lambda v: f"{v:.1f}%"),
                            textposition="outside", textfont=dict(size=10),
                            hovertemplate="<b>%{y}</b><br>No cumple P_max: <b>%{x:.1f}%</b><extra></extra>",
                        ))
                        fig_rk_spm.add_vline(x=50, line=dict(color="#ef4444", dash="dot", width=1.5))
                        fig_rk_spm.add_annotation(x=51, y=0, text="50%", showarrow=False,
                                                   font=dict(size=9, color="#ef4444"), xanchor="left")
                        fig_rk_spm.update_layout(
                            **_base_layout(t),
                            height=max(250, 30 * len(_no_spm) + 80),
                            margin=dict(t=10, r=70, b=40, l=70),
                            xaxis=dict(
                                title=dict(text="% eventos sin aporte en P_max computado",
                                           font=dict(size=11)),
                                range=[0, 90], gridcolor=t["grid"],
                            ),
                            yaxis=dict(gridcolor=t["grid"]),
                        )
                        st.plotly_chart(fig_rk_spm, use_container_width=True, config={**_CHART_CONFIG,
                            "toImageButtonOptions": {**_CHART_CONFIG["toImageButtonOptions"],
                                                     "filename": "RPF_ranking_pmax_computado_scada"}})

    # Verificación individual en P_max (solo cuando hay una unidad específica seleccionada)
    if sel_uni != "Todas":
        if fuente == "SCADA":
            _b6_render_scada_pmax(sel_uni, t)
        else:
            _b6_render_sim_pmax(sel_uni, fuente, t)


#  Tab 2: Frecuencia

def _tab_frecuencia(df: pd.DataFrame, t: dict):
    st.markdown("#### Evolución de f_min por Evento")
    ev_df = (df.groupby(["semestre", "evento", "fecha_evento"])
             .agg(f_min=("f_min_hz", "mean"), f_0=("f_0_hz", "mean"))
             .reset_index()
             .dropna(subset=["fecha_evento"])
             .sort_values("fecha_evento"))

    if ev_df.empty:
        st.info("No hay registros de f_min con fecha de evento para los filtros seleccionados.")
        return

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
        height=320,
        yaxis=dict(title=dict(text="f_min [Hz]"), range=[49.1, 50.05], gridcolor=t["grid"]),
        xaxis=dict(type="date", gridcolor=t["grid"]),
        legend=dict(orientation="h", x=0, y=1.08, font=dict(size=10)),
        margin=dict(t=30, r=16, b=40, l=60),
    )
    st.plotly_chart(fig, use_container_width=True, config={**_CHART_CONFIG,
        "toImageButtonOptions": {**_CHART_CONFIG["toImageButtonOptions"],
                                 "filename": "RPF_evolucion_fmin"}})


#  Tab 3: Reservas 

def _tab_reservas(df: pd.DataFrame, t: dict):
    st.markdown("#### Reserva Disponible vs Potencia Entregada por Unidad")

    if "p_entregada_mw" not in df.columns or df["p_entregada_mw"].isna().all():
        st.info("No hay datos de potencia entregada para los filtros seleccionados.")
        return

    res_df = (df.dropna(subset=["r_inicial_mw", "p_entregada_mw"])
              .groupby("unidad")
              .agg(reserva=("r_inicial_mw", "mean"),
                   entregada=("p_entregada_mw", "mean"),
                   p_max=("p_max_mw", "mean"))
              .reset_index()
              .sort_values("reserva", ascending=False))

    if res_df.empty:
        st.info("No hay datos de reserva para los filtros seleccionados.")
        return

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
    st.plotly_chart(fig, use_container_width=True, config={**_CHART_CONFIG,
        "toImageButtonOptions": {**_CHART_CONFIG["toImageButtonOptions"], "filename": "RPF_reservas_por_unidad"}})

    #  Estadísticas resumidas
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


#  Render principal 

def render_bloque_kpi(session_state):
    t = _theme()

    # ── Selector de fuente (por encima de todo) ───────────────────────────────
    _fuente_sel = st.radio(
        "Fuente de análisis",
        ["SCADA (histórico)", "Simulación E.0", "Simulación E.1"],
        horizontal=True, key="b06_fuente",
    )
    _fuente_key = (
        "SCADA" if "SCADA" in _fuente_sel
        else ("E.0" if "E.0" in _fuente_sel else "E.1")
    )

    # Botón para forzar recarga (limpia cache)
    _col_reload, _col_sp = st.columns([1, 5])
    with _col_reload:
        if st.button("🔄 Recargar datos", key="b06_reload"):
            _load_data.clear()
            _load_sim_data.clear()
            st.rerun()

    # ── Carga de datos según fuente ───────────────────────────────────────────
    if _fuente_key == "SCADA":
        with st.spinner("Cargando datos históricos RPF..."):
            df, _fuente_lbl, _diag_errors = _load_data()

        if df.empty:
            st.error(
                "No se encontraron datos en ninguna fuente disponible.\n\n"
                "**Opciones:** conecta a la red local (192.168.0.92) para usar PostgreSQL, "
                "o asegúrate de que `rpf_kpi_cobee.csv` esté en SharePoint `03_DATOS GEN`.",
                icon="❌",
            )
            if _diag_errors:
                with st.expander("🔍 Diagnóstico de fuentes (ver detalles del error)"):
                    for msg in _diag_errors:
                        st.code(msg)
            return

        st.caption(f"Fuente de datos: {_fuente_lbl}")
        if _diag_errors:
            with st.expander("🔍 Diagnóstico de fuentes intentadas", expanded=False):
                for msg in _diag_errors:
                    st.code(msg)
    else:
        # Simulaciones E.0 / E.1
        with st.spinner("Cargando KPIs de simulación..."):
            df_sim_full = _load_sim_data()
        _diag_errors = []

        if df_sim_full.empty:
            st.warning(
                f"No hay KPIs de simulación guardados. "
                "Ejecuta el **Bloque 5**, selecciona una unidad y pulsa "
                "**💾 Guardar KPIs P_max en Bloque 6**."
            )
            return

        df = df_sim_full[df_sim_full["fuente"] == _fuente_key].copy()
        if df.empty:
            st.warning(
                f"No hay datos guardados para **Simulación {_fuente_key}**. "
                "Ejecuta el Bloque 5 con archivos de esa versión de simulación."
            )
            return

        # Aplicar alias de columnas para reusar tabs existentes
        df["aporta_rpf"]      = df["aporta_pmax"]
        df["p_entregada_mw"]  = df["p_pmax_mw"]
        df["p_entregada_pct"] = df["dp_pct_pmax"]
        df["droop_calc_pct"]  = df["droop_calc_pmax"]
        df["f_35_hz"]         = df["f_pmax_hz"]
        df["t_35"]            = df["t_pmax_s"]
        if "fecha_evento" not in df.columns:
            df["fecha_evento"] = None

        _fuente_lbl = f"rpf_kpi_sim.csv — Simulación {_fuente_key}"
        st.caption(f"Fuente de datos: {_fuente_lbl}")

    # Métricas rápidas
    total_ev  = df.groupby(["semestre", "evento"]).ngroups
    _aporta_col = "aporta_rpf" if "aporta_rpf" in df.columns else "aporta_pmax"
    pct_si    = (df[_aporta_col] == "Sí").mean() * 100
    f_min_min = df["f_min_hz"].min() if "f_min_hz" in df.columns and not df["f_min_hz"].isna().all() else float("nan")
    n_units   = df["unidad"].nunique()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Eventos analizados", f"{total_ev}", f"{df['semestre'].nunique()} semestres")
    m2.metric("Cumplimiento RPF", f"{pct_si:.1f}%")
    m3.metric("f_min histórico", f"{f_min_min:.3f} Hz" if not pd.isna(f_min_min) else "N/D", "↓ peor nadir")
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
            df[df["semestre"] == sel_sem]["evento"].dropna().unique().tolist(),
            key=lambda e: int(m.group(1)) if (m := re.search(r"(\d+)$", e)) else -1)
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

    # Barra de tabs — mismo patrón que _v4_tab_bar del interfaz principal
    _tab_defs = [
        {"id": "cumpl", "label": "Cumplimiento"},
        {"id": "freq",  "label": "Frecuencia & Droop"},
        {"id": "res",   "label": "Reservas"},
    ]
    _sk = "v4_tab_b06_kpi"
    _ids = [td["id"] for td in _tab_defs]
    if _sk not in st.session_state or st.session_state[_sk] not in _ids:
        st.session_state[_sk] = _ids[0]
    _active = st.session_state[_sk]

    _cols = st.columns(len(_tab_defs))
    for _td, _col in zip(_tab_defs, _cols):
        with _col:
            if st.button(
                _td["label"],
                key=f"{_sk}_{_td['id']}",
                type="primary" if _td["id"] == _active else "secondary",
                use_container_width=True,
            ):
                st.session_state[_sk] = _td["id"]
                st.rerun()

    if _active == "cumpl":
        _tab_cumplimiento(dff, t, sel_uni=sel_uni, fuente=_fuente_key)
    elif _active == "freq":
        _tab_frecuencia(dff, t)
    elif _active == "res":
        _tab_reservas(dff, t)
