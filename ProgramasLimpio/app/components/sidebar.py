# app/components/sidebar.py
"""Sidebar persistente: selectores dinámicos y chips de estado del pipeline."""
from __future__ import annotations

import os
import sys

# Asegura que la raíz del proyecto esté en sys.path
_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import streamlit as st
from core.config import PipelinePaths

# Colores para los chips de fase
_CHIP_OK  = "#2e7d32"   # verde
_CHIP_NOK = "#9e9e9e"   # gris


def _chip_html(label: str, ok: bool) -> str:
    bg    = _CHIP_OK if ok else _CHIP_NOK
    icon  = "✅" if ok else "⬜"
    return (
        f'<span style="background:{bg};color:#fff;padding:3px 10px;'
        f'border-radius:12px;font-size:12px;font-weight:600;">'
        f'{icon} {label}</span>'
    )


def render_sidebar() -> None:
    """
    Renderiza el sidebar completo.
    Debe llamarse en cada página Streamlit.
    Lee y escribe st.session_state.paths / .semestre / .evento_num.
    """
    # ── Inicializar estado ────────────────────────────────────────────────────
    if "paths" not in st.session_state:
        st.session_state.paths = PipelinePaths.from_toml()
    if "semestre" not in st.session_state:
        st.session_state.semestre = None
    if "evento_num" not in st.session_state:
        st.session_state.evento_num = None

    paths: PipelinePaths = st.session_state.paths

    # ── Cabecera ──────────────────────────────────────────────────────────────
    st.sidebar.markdown(
        "<h2 style='margin:0;color:#1565C0;'>⚡ Pipeline SIN</h2>"
        "<p style='margin:0;color:#616161;font-size:13px;'>PowerFactory · COBEE S.A.</p>",
        unsafe_allow_html=True,
    )
    st.sidebar.divider()

    # ── Configuración de rutas ────────────────────────────────────────────────
    with st.sidebar.expander("⚙ Rutas de datos", expanded=False):
        raiz_cndc = st.text_input(
            "Raíz CNDC",
            value=paths.raiz_cndc,
            key="cfg_raiz_cndc",
        )
        raiz_loc = st.text_input(
            "Carpeta loc_names",
            value=paths.raiz_loc_names,
            key="cfg_raiz_loc",
        )
        raiz_prog = st.text_input(
            "Carpeta programas",
            value=paths.raiz_programas,
            key="cfg_raiz_prog",
        )
        pf_modelo = st.text_input(
            "Nombre modelo PF",
            value=paths.pf_modelo,
            key="cfg_pf_modelo",
        )
        if st.button("💾 Guardar rutas", key="btn_save_paths", use_container_width=True):
            paths.raiz_cndc       = raiz_cndc
            paths.raiz_loc_names  = raiz_loc
            paths.raiz_programas  = raiz_prog
            paths.pf_modelo       = pf_modelo
            paths.save_toml()
            st.session_state.paths = paths
            st.success("Rutas guardadas en ~/.cobee_pipeline.toml", icon="✅")

    st.sidebar.divider()

    # ── Selección semestre ────────────────────────────────────────────────────
    semestres = paths.semestres()
    if not semestres:
        st.sidebar.warning("⚠ Raíz CNDC no encontrada o vacía.")
        st.session_state.semestre = None
        st.session_state.evento_num = None
    else:
        sem_idx = 0
        if st.session_state.semestre in semestres:
            sem_idx = semestres.index(st.session_state.semestre)
        semestre = st.sidebar.selectbox(
            "📅 Semestre",
            semestres,
            index=sem_idx,
            key="sel_semestre",
        )
        st.session_state.semestre = semestre

        # ── Selección evento ──────────────────────────────────────────────────
        eventos = paths.eventos_de_semestre(semestre)
        if not eventos:
            st.sidebar.info("Sin eventos encontrados en este semestre.")
            st.session_state.evento_num = None
        else:
            ev_idx = 0
            if st.session_state.evento_num in eventos:
                ev_idx = eventos.index(st.session_state.evento_num)
            evento_num = st.sidebar.selectbox(
                "🔢 Evento",
                eventos,
                index=ev_idx,
                format_func=lambda n: f"Evento {n}",
                key="sel_evento",
            )
            st.session_state.evento_num = evento_num

            # Mostrar metadata rápida del evento si existe
            tabla_path = paths.tabla_eventos_path(semestre)
            if tabla_path and tabla_path.exists():
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(str(tabla_path), data_only=True)
                    sh = wb.active
                    for fila in sh.iter_rows(min_row=3, values_only=True):
                        if fila[0] is not None:
                            try:
                                if int(fila[0]) == evento_num:
                                    st.sidebar.caption(
                                        f"📆 {fila[1]}  \n"
                                        f"⚡ {fila[2]}  \n"
                                        f"📊 {fila[3]} MW | f₀={fila[5]} Hz"
                                    )
                                    break
                            except (TypeError, ValueError):
                                pass
                except Exception:
                    pass

    st.sidebar.divider()

    # ── Estado del pipeline ───────────────────────────────────────────────────
    st.sidebar.markdown("**Estado del pipeline**")
    fase_status = paths.fase_status(
        st.session_state.semestre,
        st.session_state.evento_num,
    )
    for fase, label in [
        ("fase0", "Fase 0 — Extracción Red"),
        ("fase1", "Fase 1 — Catálogos"),
        ("fase2", "Fase 2 — Cond. Iniciales"),
        ("fase3", "Fase 3 — Simulación PF"),
    ]:
        ok = fase_status[fase]
        icon = "✅" if ok else "⬜"
        color = _CHIP_OK if ok else _CHIP_NOK
        st.sidebar.markdown(
            f'<div style="padding:3px 0;">'
            f'<span style="background:{color};color:#fff;padding:2px 8px;'
            f'border-radius:10px;font-size:11px;font-weight:600;">{icon} {label}</span>'
            f"</div>",
            unsafe_allow_html=True,
        )
