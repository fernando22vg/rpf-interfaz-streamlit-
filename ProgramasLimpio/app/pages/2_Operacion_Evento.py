"""
▶️ Operación por Evento — Fases 2 y 3
Repetir para cada uno de los ~28 eventos CNDC por semestre.
"""
from __future__ import annotations

import io
import os
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Optional

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Operación por Evento — Pipeline SIN",
    page_icon="▶️",
    layout="wide",
    initial_sidebar_state="expanded",
)

from app.components.sidebar import render_sidebar
from core.config import PipelinePaths
from core.data_extraction import (
    ExtraccionResult,
    cargar_condiciones_iniciales,
    cargar_datos_simulacion,
    extraer_datos_evento,
    leer_tabla_eventos,
)

render_sidebar()

paths: PipelinePaths   = st.session_state.paths
semestre: Optional[str] = st.session_state.get("semestre")
evento_num: Optional[int] = st.session_state.get("evento_num")

# ── Guard: verificar selección ────────────────────────────────────────────────
st.title("▶️ Operación por Evento")

if not semestre or not evento_num:
    st.warning(
        "Selecciona un **semestre** y un **evento** en el menú lateral para continuar.",
        icon="👈",
    )
    st.stop()

ev_path = paths.evento_path(semestre, evento_num)
st.caption(
    f"**{semestre}  ›  Evento {evento_num}**  |  `{ev_path}`"
)

# ── Cargar datos existentes en sesión ─────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _load_datos_simulacion(path: str) -> Optional[ExtraccionResult]:
    return cargar_datos_simulacion(path)

@st.cache_data(show_spinner=False)
def _load_cond_iniciales(path: str) -> Optional[dict]:
    return cargar_condiciones_iniciales(path)

@st.cache_data(show_spinner=False)
def _load_tabla_eventos(semestre: str, raiz: str):
    try:
        evs, _ = leer_tabla_eventos(semestre, raiz)
        return pd.DataFrame(evs)
    except Exception:
        return pd.DataFrame()

datos_sim_path = paths.datos_simulacion_path(semestre, evento_num)
cond_ini_path  = paths.condiciones_iniciales_path(semestre, evento_num)

datos_sim: Optional[ExtraccionResult] = (
    _load_datos_simulacion(str(datos_sim_path)) if datos_sim_path else None
)
cond_ini: Optional[dict] = (
    _load_cond_iniciales(str(cond_ini_path)) if cond_ini_path else None
)

# ═══════════════════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Resumen del Evento",
    "📥 Datos de Entrada",
    "⚡ Condiciones Iniciales",
    "🖥 Estado de Simulación",
    "📊 Resultados",
])


# ──────────────────────────────────────────────────────────────────────────────
# TAB 1 — Resumen del Evento
# ──────────────────────────────────────────────────────────────────────────────
with tab1:
    col_meta, col_tabla = st.columns([1, 2])

    with col_meta:
        st.subheader("Metadata del evento")
        if datos_sim and not datos_sim.info_evento.empty:
            df_info = datos_sim.info_evento.set_index("Campo")
            for campo, row in df_info.iterrows():
                st.metric(str(campo), str(row.iloc[0]))
        else:
            # Intentar leer directamente de la tabla de eventos
            df_tabla = _load_tabla_eventos(semestre, paths.raiz_cndc)
            if not df_tabla.empty:
                ev_row = df_tabla[df_tabla["num"] == evento_num]
                if not ev_row.empty:
                    r = ev_row.iloc[0]
                    st.metric("Fecha y hora", r.get("fecha_hora", "—"))
                    st.metric("Disparo", r.get("desconexion", "—"))
                    st.metric("Potencia desconectada", f"{r.get('pot_desc_MW', '—')} MW")
                    st.metric("Demanda SIN", f"{r.get('demanda_MW', '—')} MW")
                    st.metric("f₀", f"{r.get('f0_Hz', '—')} Hz")
                    st.metric("f_min", f"{r.get('fmin_Hz', '—')} Hz")
            else:
                st.info("Ejecuta la Fase 2A para cargar los datos del evento.")

    with col_tabla:
        st.subheader(f"Todos los eventos — {semestre}")
        df_tabla = _load_tabla_eventos(semestre, paths.raiz_cndc)
        if not df_tabla.empty:
            # Columna estado
            def _estado_ev(row):
                ev = int(row["num"])
                p2 = paths.datos_simulacion_path(semestre, ev)
                p3 = paths.condiciones_iniciales_path(semestre, ev)
                if p2 and p3:
                    return "✅ Fase 2 completa"
                if p2:
                    return "🔶 Fase 2A lista"
                return "⏳ Pendiente"

            df_tabla["Estado"] = df_tabla.apply(_estado_ev, axis=1)
            # Resaltar evento seleccionado
            def _highlight_ev(row):
                if row["num"] == evento_num:
                    return ["background-color: #E3F2FD"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df_tabla.style.apply(_highlight_ev, axis=1),
                use_container_width=True,
                hide_index=True,
                height=400,
            )
        else:
            st.warning("No se encontró Tabla_Eventos_*.xlsx en este semestre.")

    st.divider()

    # Botón para cargar evento (limpiar cache y recargar)
    if st.button("🔄 Recargar datos del evento", key="btn_reload_ev"):
        st.cache_data.clear()
        st.rerun()


# ──────────────────────────────────────────────────────────────────────────────
# TAB 2 — Datos de Entrada
# ──────────────────────────────────────────────────────────────────────────────
with tab2:
    st.subheader("Datos de entrada CNDC")

    if datos_sim is None:
        st.info(
            "No hay datos de simulación para este evento. "
            "Ejecuta la **Fase 2A** en la sección de Condiciones Iniciales.",
            icon="ℹ️",
        )
    else:
        # KPI cards
        df_gen = datos_sim.generadores_pgini
        df_car = datos_sim.cargas_plini

        # Columna hora del evento (última hora numérica)
        hora_cols_gen = [c for c in df_gen.columns
                         if ":" in str(c) and c not in ("Generador_CNDC", "Codigo STI", "Estado")]
        hora_cols_car = [c for c in df_car.columns
                         if ":" in str(c) and c not in ("Nodo_CNDC",)]

        hora_ev_gen = hora_cols_gen[-1] if hora_cols_gen else None
        hora_ev_car = hora_cols_car[-1] if hora_cols_car else None

        p_gen_total = df_gen[hora_ev_gen].sum() if hora_ev_gen else 0.0
        p_dem_total = df_car[hora_ev_car].sum() if hora_ev_car else 0.0
        balance = p_gen_total - p_dem_total

        kc1, kc2, kc3, kc4, kc5 = st.columns(5)
        kc1.metric("Unidades generadoras", len(df_gen))
        kc2.metric("Nodos de demanda", len(df_car))
        kc3.metric("ΣPgen [MW]", f"{p_gen_total:.2f}")
        kc4.metric("ΣPdem [MW]", f"{p_dem_total:.2f}")
        delta_color = "normal" if abs(balance) < 50 else "inverse"
        kc5.metric("ΔBalance [MW]", f"{balance:+.2f}", delta=f"{balance:+.1f}", delta_color=delta_color)

        st.divider()

        sub1, sub2, sub3 = st.tabs(["⚡ Generadores (DC+DCDR)", "📉 Demanda (DEENER)", "🔍 P₀ Medido"])

        with sub1:
            filtro_gen = st.text_input("Filtrar generador…", key="flt_gen")
            df_show = df_gen if not filtro_gen else df_gen[
                df_gen["Generador_CNDC"].str.contains(filtro_gen, case=False, na=False)
            ]
            # Color-coding por Estado
            def _color_gen(row):
                if row.get("Estado") == "Mantenimiento":
                    return ["background-color:#e0e0e0"] * len(row)
                if hora_ev_gen and pd.notna(row.get(hora_ev_gen)):
                    return ["background-color:#e8f5e9"] * len(row)
                return ["background-color:#fff9c4"] * len(row)

            st.dataframe(
                df_show.style.apply(_color_gen, axis=1),
                use_container_width=True,
                height=400,
                hide_index=True,
            )
            st.caption(
                "🟩 Verde = P₀ real de tabla_resultados  |  "
                "🟨 Amarillo = valor de hora anterior (fallback)  |  "
                "⬜ Gris = mantenimiento"
            )

        with sub2:
            filtro_car = st.text_input("Filtrar nodo…", key="flt_car")
            df_car_show = df_car if not filtro_car else df_car[
                df_car["Nodo_CNDC"].str.contains(filtro_car, case=False, na=False)
            ]
            st.dataframe(df_car_show, use_container_width=True, height=400, hide_index=True)

        with sub3:
            df_po = datos_sim.p0_inicial
            if df_po.empty:
                st.info("No se encontraron datos P₀ para este evento.")
            else:
                c_po1, c_po2 = st.columns([2, 1])
                with c_po1:
                    st.dataframe(df_po, use_container_width=True, hide_index=True)
                with c_po2:
                    fuentes = df_po["Fuente"].value_counts().reset_index()
                    fuentes.columns = ["Carpeta", "N° unidades"]
                    fig_po = px.bar(
                        fuentes, x="N° unidades", y="Carpeta", orientation="h",
                        title="Unidades P₀ por carpeta",
                        color_discrete_sequence=["#1976D2"],
                    )
                    fig_po.update_layout(margin=dict(l=10, r=10, t=40, b=10), height=300)
                    st.plotly_chart(fig_po, use_container_width=True)

    st.divider()
    # Exportar datos de entrada a Excel
    if datos_sim:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            datos_sim.info_evento.to_excel(w, sheet_name="Info_Evento", index=False)
            datos_sim.generadores_pgini.to_excel(w, sheet_name="Generadores_pgini", index=False)
            datos_sim.cargas_plini.to_excel(w, sheet_name="Cargas_plini", index=False)
            datos_sim.p0_inicial.to_excel(w, sheet_name="P0_inicial", index=False)
        st.download_button(
            "⬇ Descargar datos_simulacion.xlsx",
            data=buf.getvalue(),
            file_name=f"datos_simulacion_Ev{evento_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ──────────────────────────────────────────────────────────────────────────────
# TAB 3 — Condiciones Iniciales
# ──────────────────────────────────────────────────────────────────────────────
with tab3:
    st.subheader("Condiciones Iniciales para PowerFactory")

    # ── Panel de ejecución Fase 2A y 2B ──────────────────────────────────────
    col_2a, col_2b = st.columns(2)

    with col_2a:
        st.markdown("**Fase 2A — Extracción de flujos CNDC**")
        st.caption("`ExtFLujos2daO.py` → `datos_simulacion_*.xlsx`")
        if datos_sim_path:
            st.success(f"✅ {datos_sim_path.name}", icon="✅")
        else:
            st.warning("Aún no ejecutado")

        btn_2a = st.button(
            "▶ Ejecutar Fase 2A",
            key="btn_fase2a",
            type="primary" if not datos_sim_path else "secondary",
            use_container_width=True,
        )

        if btn_2a:
            prog_bar = st.progress(0.0, text="Iniciando…")
            log_2a   = st.empty()

            def _cb_2a(pct: float, msg: str) -> None:
                prog_bar.progress(pct, text=msg)
                log_2a.code(msg)

            try:
                result = extraer_datos_evento(
                    semestre=semestre,
                    evento_num=evento_num,
                    raiz_cndc=paths.raiz_cndc,
                    loc_names_gen_path=str(paths.loc_gen),
                    progress_cb=_cb_2a,
                )
                prog_bar.progress(1.0, text="Completado")
                st.success(
                    f"✅ Generadas {len(result.generadores_pgini)} unidades  |  "
                    f"P₀ real: {result.n_con_po}  |  "
                    f"Fallback: {result.n_fallback}  |  "
                    f"Mantenimiento: {result.n_mant}",
                    icon="✅",
                )
                if result.warnings:
                    for w in result.warnings:
                        st.warning(w)
                st.cache_data.clear()
                st.rerun()
            except Exception as exc:
                st.error(f"Error: {exc}")

    with col_2b:
        st.markdown("**Fase 2B — Cálculo de condiciones iniciales**")
        st.caption("`CondInicialesPF.py` → `condiciones_iniciales_*.xlsx`")
        if cond_ini_path:
            st.success(f"✅ {cond_ini_path.name}", icon="✅")
        else:
            st.warning("Aún no ejecutado")

        btn_2b = st.button(
            "▶ Ejecutar Fase 2B",
            key="btn_fase2b",
            type="primary" if (datos_sim_path and not cond_ini_path) else "secondary",
            use_container_width=True,
            disabled=(datos_sim_path is None),
        )

        if btn_2b:
            script_path = Path(paths.raiz_programas) / "CondInicialesPF.py"
            if not script_path.exists():
                st.error(f"No se encontró: {script_path}")
            else:
                log_2b = st.empty()
                with st.spinner("Ejecutando CondInicialesPF.py…"):
                    try:
                        proc = subprocess.run(
                            [sys.executable, str(script_path)],
                            capture_output=True, text=True,
                            cwd=str(paths.raiz_programas),
                            timeout=300,
                        )
                        log_2b.code(
                            (proc.stdout or "") + (proc.stderr or ""),
                            language="",
                        )
                        if proc.returncode == 0:
                            st.success("✅ Condiciones iniciales generadas.", icon="✅")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("❌ El script terminó con error.")
                    except subprocess.TimeoutExpired:
                        st.error("Timeout: el script tardó más de 5 minutos.")
                    except Exception as e:
                        st.error(str(e))

    st.divider()

    # ── Visualización de condiciones iniciales ────────────────────────────────
    if cond_ini is None:
        st.info("Ejecuta la Fase 2B para ver las condiciones iniciales.", icon="ℹ️")
    else:
        tab_gen, tab_car, tab_res, tab_chart = st.tabs([
            "⚡ pgini_GEN", "📉 plini_CAR", "📋 Resumen", "📊 Gráfico"
        ])

        with tab_gen:
            df_pgini = cond_ini.get("pgini_GEN", pd.DataFrame())
            if df_pgini.empty:
                st.warning("Hoja pgini_GEN vacía o no encontrada.")
            else:
                # Color por Fuente
                _COLORES_FUENTE = {
                    "P0_medido":         "#c6efce",
                    "CNDC_proporcional": "#ffeb9c",
                    "sin_despacho":      "#d9d9d9",
                    "Mantenimiento":     "#ffcccc",
                    "disparo":           "#ff9999",
                }
                def _color_pgini(row):
                    fuente = str(row.get("Fuente", "")).strip()
                    bg = _COLORES_FUENTE.get(fuente, "")
                    return [f"background-color:{bg}"] * len(row) if bg else [""] * len(row)

                st.dataframe(
                    df_pgini.style.apply(_color_pgini, axis=1),
                    use_container_width=True,
                    hide_index=True,
                    height=500,
                )
                st.caption(
                    "🟩 P0_medido  |  🟨 CNDC_proporcional  |  ⬜ sin_despacho  |  "
                    "🔴 Mantenimiento / disparo"
                )

        with tab_car:
            df_plini = cond_ini.get("plini_CAR", pd.DataFrame())
            if df_plini.empty:
                st.warning("Hoja plini_CAR vacía o no encontrada.")
            else:
                filtro_dist = st.text_input("Filtrar distribuidor…", key="flt_dist")
                if filtro_dist and "Distribuidor" in df_plini.columns:
                    df_plini = df_plini[
                        df_plini["Distribuidor"].str.contains(
                            filtro_dist, case=False, na=False
                        )
                    ]
                st.dataframe(df_plini, use_container_width=True, hide_index=True, height=480)

        with tab_res:
            df_res = cond_ini.get("Resumen", pd.DataFrame())
            if df_res.empty:
                st.warning("Hoja Resumen vacía.")
            else:
                st.dataframe(df_res, use_container_width=True, hide_index=True)

        with tab_chart:
            df_pgini_chart = cond_ini.get("pgini_GEN", pd.DataFrame())
            if "Tipo" in df_pgini_chart.columns and "pgini_MW" in df_pgini_chart.columns:
                df_tipo = (
                    df_pgini_chart.groupby("Tipo")["pgini_MW"]
                    .sum()
                    .reset_index()
                    .rename(columns={"pgini_MW": "Pgen [MW]"})
                )
                _COLORS_TIPO = {
                    "HIDRO":  "#1565C0",
                    "SOLAR":  "#F57F17",
                    "EÓLICO": "#2E7D32",
                    "TERMO":  "#B71C1C",
                }
                fig_tipo = px.bar(
                    df_tipo, x="Tipo", y="Pgen [MW]",
                    color="Tipo",
                    color_discrete_map=_COLORS_TIPO,
                    title="Generación total por tecnología",
                    text_auto=".1f",
                )
                fig_tipo.update_layout(showlegend=False, height=380)
                st.plotly_chart(fig_tipo, use_container_width=True)

                # Validador de balance
                if "Fuente" in df_pgini_chart.columns:
                    pgen_total = df_pgini_chart["pgini_MW"].sum() if "pgini_MW" in df_pgini_chart.columns else 0
                    df_res_check = cond_ini.get("Resumen", pd.DataFrame())
                    pdem_total = 0.0
                    if not df_res_check.empty and "Valor" in df_res_check.columns:
                        fila_dem = df_res_check[
                            df_res_check.get("Campo", pd.Series()).str.contains("Demanda", na=False)
                        ]
                        if not fila_dem.empty:
                            try:
                                pdem_total = float(fila_dem.iloc[0]["Valor"])
                            except Exception:
                                pass
                    if pgen_total > 0:
                        tol_mw = 50.0
                        delta = pgen_total - pdem_total
                        ok_balance = abs(delta) <= tol_mw
                        if ok_balance:
                            st.success(
                                f"✅ Balance OK: Pgen={pgen_total:.1f} MW  |  "
                                f"Pdem={pdem_total:.1f} MW  |  Δ={delta:+.1f} MW",
                            )
                        else:
                            st.warning(
                                f"⚠ Balance fuera de tolerancia (±{tol_mw} MW): "
                                f"Pgen={pgen_total:.1f}  Pdem={pdem_total:.1f}  Δ={delta:+.1f} MW"
                            )
            else:
                st.info("Las columnas 'Tipo' y 'pgini_MW' no están disponibles en pgini_GEN.")


# ──────────────────────────────────────────────────────────────────────────────
# TAB 4 — Estado de Simulación (Fase 3)
# ──────────────────────────────────────────────────────────────────────────────
with tab4:
    st.subheader("Fase 3 — Carga en PowerFactory y ejecución RMS")
    st.caption("`CargaCondIniciales_PF.py`  →  pgini/plini asignados · Load Flow AC · RMS init")

    col_btn3, col_info3 = st.columns([1, 2])
    with col_btn3:
        btn_fase3 = st.button(
            "▶ Ejecutar Fase 3 (Cargar en PowerFactory)",
            key="btn_fase3",
            type="primary",
            use_container_width=True,
            disabled=(cond_ini_path is None),
        )
        if cond_ini_path is None:
            st.caption("⚠ Primero completa la Fase 2B.")

    with col_info3:
        st.warning(
            "Esta acción modifica el escenario activo en PowerFactory. "
            "Asegúrate de que el proyecto correcto esté abierto y no haya "
            "simulaciones en curso.",
            icon="⚠️",
        )

    if btn_fase3:
        # Confirmación modal con checkbox
        confirmar = st.checkbox(
            f"Confirmo: quiero cargar el Evento {evento_num} en PowerFactory ({paths.pf_modelo})",
            key="chk_confirmar_pf",
        )
        if confirmar:
            script_path = Path(paths.raiz_programas) / "CargaCondIniciales_PF.py"
            if not script_path.exists():
                st.error(f"No se encontró: {script_path}")
            else:
                log_pf = st.empty()
                prog_pf = st.progress(0.0, text="Iniciando PowerFactory…")
                with st.spinner("Ejecutando CargaCondIniciales_PF.py…"):
                    try:
                        proc = subprocess.Popen(
                            [sys.executable, str(script_path)],
                            stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT,
                            text=True,
                            bufsize=1,
                            cwd=str(paths.raiz_programas),
                        )
                        lines: list[str] = []
                        for line in proc.stdout:
                            lines.append(line.rstrip())
                            log_pf.code("\n".join(lines[-80:]), language="")
                            # Actualizar progreso por secciones detectadas
                            txt = line.lower()
                            if "asignando pgini" in txt:
                                prog_pf.progress(0.30, "Asignando pgini…")
                            elif "asignando plini" in txt:
                                prog_pf.progress(0.55, "Asignando plini…")
                            elif "load flow" in txt:
                                prog_pf.progress(0.70, "Ejecutando Load Flow…")
                            elif "convergió" in txt or "converged" in txt:
                                prog_pf.progress(0.85, "Load Flow convergido ✓")
                            elif "rms" in txt or "cominc" in txt:
                                prog_pf.progress(0.95, "Inicializando RMS…")
                        proc.wait()
                        prog_pf.progress(1.0, "Completado")
                        if proc.returncode == 0:
                            st.success(
                                "✅ PowerFactory — Escenario cargado y RMS inicializado.",
                                icon="✅",
                            )
                            st.session_state["fase3_ok"] = True
                        else:
                            st.error("❌ El script terminó con error. Revisa el log.")
                    except Exception as exc:
                        st.error(str(exc))
        else:
            st.info("Marca la casilla de confirmación para ejecutar.", icon="☑️")

    # Estado actual Fase 3
    st.divider()
    st.markdown("**Verificación rápida de artefactos**")
    fase_status = paths.fase_status(semestre, evento_num)
    items = [
        ("datos_simulacion_*.xlsx", datos_sim_path),
        ("condiciones_iniciales_*.xlsx", cond_ini_path),
    ]
    for nombre, p in items:
        if p and p.exists():
            mtime = time.strftime("%d/%m/%Y %H:%M", time.localtime(p.stat().st_mtime))
            st.markdown(f"✅ `{nombre}` — {p.name} ({mtime})")
        else:
            st.markdown(f"❌ `{nombre}` — no encontrado")
    if st.session_state.get("fase3_ok"):
        st.success("✅ Fase 3 ejecutada en esta sesión.", icon="⚡")


# ──────────────────────────────────────────────────────────────────────────────
# TAB 5 — Resultados
# ──────────────────────────────────────────────────────────────────────────────
with tab5:
    st.subheader("Resultados de simulación RMS")
    st.caption(
        "Curvas de frecuencia, potencia y tensión obtenidas tras ejecutar la "
        "simulación dinámica en PowerFactory."
    )

    # Buscar archivos de resultados EMF o CSV en la carpeta del evento
    import glob as _glob
    emf_files = sorted(_glob.glob(str(ev_path / "**" / "*.emf"), recursive=True))
    csv_files = sorted(_glob.glob(str(ev_path / "**" / "*.csv"), recursive=True))
    result_files = emf_files + csv_files

    if not result_files:
        st.info(
            "No se encontraron archivos de resultados (.emf / .csv) para este evento.  \n"
            "Ejecuta la simulación RMS en PowerFactory y exporta las curvas.",
            icon="📂",
        )
    else:
        sel_file = st.selectbox(
            "Archivo de resultados",
            result_files,
            format_func=os.path.basename,
            key="sel_result_file",
        )

        if sel_file and sel_file.endswith(".csv"):
            try:
                df_res = pd.read_csv(sel_file, sep=None, engine="python")
                cols_num = df_res.select_dtypes("number").columns.tolist()
                tiempo_col = next(
                    (c for c in df_res.columns if "tiempo" in c.lower() or "time" in c.lower()),
                    cols_num[0] if cols_num else None,
                )
                if tiempo_col and len(cols_num) > 1:
                    otras = [c for c in cols_num if c != tiempo_col]
                    sel_vars = st.multiselect(
                        "Variables a graficar",
                        otras,
                        default=otras[:4],
                        key="sel_vars_csv",
                    )
                    if sel_vars:
                        fig = go.Figure()
                        for var in sel_vars:
                            fig.add_trace(go.Scatter(
                                x=df_res[tiempo_col], y=df_res[var],
                                name=var, mode="lines",
                            ))
                        fig.update_layout(
                            xaxis_title="Tiempo (s)",
                            hovermode="x unified",
                            height=450,
                        )
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.dataframe(df_res, use_container_width=True)
            except Exception as e:
                st.error(f"No se pudo leer el CSV: {e}")
        elif sel_file and sel_file.endswith(".emf"):
            st.info(
                "Para visualizar archivos .emf, usa la app **EMF Explorer** "
                "(`streamlit run streamlit.py`).",
                icon="📈",
            )

    # Descarga masiva de resultados
    st.divider()
    st.markdown("**Exportar condiciones iniciales a Excel**")
    if cond_ini:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            for sheet, df in cond_ini.items():
                df.to_excel(w, sheet_name=sheet[:31], index=False)
        st.download_button(
            "⬇ Descargar condiciones_iniciales.xlsx",
            data=buf.getvalue(),
            file_name=f"condiciones_iniciales_Ev{evento_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Condiciones iniciales no disponibles. Ejecuta la Fase 2B primero.")


# ══════════════════════════════════════════════════════════════════════════════
# Progress bar sticky (bottom) — muestra fase activa
# ══════════════════════════════════════════════════════════════════════════════
fase_status = paths.fase_status(semestre, evento_num)
fases = ["Fase 2A", "Fase 2B", "Fase 3"]
completadas = sum([
    1 if datos_sim_path else 0,
    1 if cond_ini_path else 0,
    1 if st.session_state.get("fase3_ok") else 0,
])
st.markdown("---")
prog_cols = st.columns(3)
for i, (col_prog, label) in enumerate(zip(prog_cols, fases)):
    done = i < completadas
    active = i == completadas
    bg = "#2e7d32" if done else ("#1565C0" if active else "#e0e0e0")
    txt_color = "#fff" if (done or active) else "#616161"
    icon = "✅" if done else ("🔵" if active else "⬜")
    col_prog.markdown(
        f'<div style="background:{bg};color:{txt_color};text-align:center;'
        f'padding:8px;border-radius:8px;font-weight:600;font-size:13px;">'
        f'{icon} {label}</div>',
        unsafe_allow_html=True,
    )
