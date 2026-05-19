"""
Pipeline SIN PowerFactory — Interfaz Streamlit
===============================================
Ejecutar desde ProgramasLimpio/:
    streamlit run app/streamlit_app.py

Página de inicio: resumen del pipeline y estado de cada fase.
"""
from __future__ import annotations

import os
import sys

# ── Ruta raíz del proyecto en sys.path (antes de cualquier import local) ──────
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import streamlit as st

st.set_page_config(
    page_title="Pipeline SIN — PowerFactory",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

from app.components.sidebar import render_sidebar
from core.config import PipelinePaths

# ── Sidebar (inicializa session_state también) ────────────────────────────────
render_sidebar()

# ── Título ────────────────────────────────────────────────────────────────────
st.markdown(
    "<h1 style='color:#1A237E;margin-bottom:4px;'>⚡ Pipeline SIN — PowerFactory</h1>"
    "<p style='color:#616161;margin-top:0;'>COBEE S.A. · Sistema Interconectado Nacional Bolivia"
    " · DIgSILENT PowerFactory · Modelo <code>PMP_NOV25_OCT29_31102025</code></p>",
    unsafe_allow_html=True,
)
st.divider()

paths: PipelinePaths = st.session_state.paths
semestre  = st.session_state.get("semestre")
evento_num = st.session_state.get("evento_num")

# ── KPI de estado del pipeline ────────────────────────────────────────────────
fase_status = paths.fase_status(semestre, evento_num)

_COLOR = {True: "#2e7d32", False: "#9e9e9e"}
_BG    = {True: "#e8f5e9", False: "#f5f5f5"}
_ICON  = {True: "✅", False: "⏳"}

cols = st.columns(4)
for col, (fase, titulo, desc) in zip(
    cols,
    [
        ("fase0", "Fase 0", "Extracción de Red"),
        ("fase1", "Fase 1", "Catálogos de Referencia"),
        ("fase2", "Fase 2", "Condiciones Iniciales"),
        ("fase3", "Fase 3", "Simulación PF"),
    ],
):
    ok = fase_status[fase]
    col.markdown(
        f"""<div style="background:{_BG[ok]};border-left:4px solid {_COLOR[ok]};
        border-radius:8px;padding:14px 16px;">
        <div style="font-size:13px;color:{_COLOR[ok]};font-weight:700;
        text-transform:uppercase;letter-spacing:1px;">{titulo}</div>
        <div style="font-size:18px;font-weight:700;margin:4px 0;">{_ICON[ok]} {"Listo" if ok else "Pendiente"}</div>
        <div style="font-size:12px;color:#616161;">{desc}</div>
        </div>""",
        unsafe_allow_html=True,
    )

st.divider()

# ── Guía de uso ───────────────────────────────────────────────────────────────
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("⚙️ Setup del Modelo  *(1 sola vez)*")
    st.markdown("""
Ejecutar cuando cambie el modelo PowerFactory:

| Fase | Script | Salida |
|------|--------|--------|
| 0 | `DatsoGENBUSLNE.py` | `DatosSINdigsilent.xlsx` |
| 1a | `loc_namesGEN.py` | `loc_names_gen.xlsx` |
| 1b | `loc_names_xfo.py` | `loc_names_xfo.xlsx` |
| 1c | `loc_namesLineas.py` | `loc_names_lineas.xlsx` |
| 1d | `MapeoRetirosSTI_v6.py` | `loc_name_cargas.xlsx` |

→ Ve a **⚙️ Setup del Modelo** en el menú lateral.
""")

with col_b:
    st.subheader("▶️ Operación por Evento  *(~28 eventos)*")
    st.markdown("""
Para cada evento CNDC:

| Fase | Script | Salida |
|------|--------|--------|
| 2a | `ExtFLujos2daO.py` | `datos_simulacion_*.xlsx` |
| 2b | `CondInicialesPF.py` | `condiciones_iniciales_*.xlsx` |
| 3 | `CargaCondIniciales_PF.py` | Escenario PF listo ✓ |

→ Selecciona semestre + evento en el menú lateral.
→ Ve a **▶️ Operación por Evento**.
""")

st.divider()

# ── Rutas configuradas ────────────────────────────────────────────────────────
st.subheader("Rutas activas")
c1, c2 = st.columns(2)
with c1:
    st.code(f"Raíz CNDC  : {paths.raiz_cndc}")
    st.code(f"loc_names  : {paths.raiz_loc_names}")
with c2:
    st.code(f"Programas  : {paths.raiz_programas}")
    st.code(f"Modelo PF  : {paths.pf_modelo}")

if semestre:
    st.info(
        f"**Semestre activo:** {semestre}  |  "
        f"**Evento activo:** {f'Evento {evento_num}' if evento_num else '—'}",
        icon="📅",
    )
else:
    st.info(
        "Selecciona un **semestre** y **evento** en el menú lateral para habilitar "
        "las fases de operación por evento.",
        icon="👈",
    )
