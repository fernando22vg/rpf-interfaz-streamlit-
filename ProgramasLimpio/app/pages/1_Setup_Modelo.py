"""
⚙️ Setup del Modelo — Fase 0 y Fase 1
Ejecutar una sola vez cuando cambie el modelo PowerFactory.
"""
from __future__ import annotations

import io
import os
import subprocess
import sys
import time
from pathlib import Path

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Setup del Modelo — Pipeline SIN",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

from app.components.sidebar import render_sidebar
from core.config import PipelinePaths

render_sidebar()

paths: PipelinePaths = st.session_state.paths
SCRIPTS_1VEZ = Path(paths.raiz_programas) / "Programas_1_uso_modelo"

# ── Definición de pasos ───────────────────────────────────────────────────────
PASOS = [
    {
        "id":      "fase0",
        "fase":    "Fase 0",
        "titulo":  "Extracción de Red",
        "script":  "DatsoGENBUSLNE.py",
        "salida":  lambda p: p.datos_sin,
        "desc":    (
            "Conecta PowerFactory vía COM API, ejecuta Load Flow y extrae "
            "barras, líneas, generadores, cargas y transformadores del modelo activo."
        ),
        "tags":    ["COM API", "Load Flow", "Barras", "Líneas", "Generadores", "Cargas", "XFOs"],
    },
    {
        "id":      "fase1a",
        "fase":    "Fase 1a",
        "titulo":  "Catálogo de Generadores",
        "script":  "loc_namesGEN.py",
        "salida":  lambda p: p.loc_gen,
        "desc":    "Mapea generadores STI → loc_name PowerFactory + tipo (HIDRO / SOLAR / EÓLICO / TERMO).",
        "tags":    ["HIDRO", "SOLAR", "EÓLICO", "TERMO", "STI → loc_name"],
    },
    {
        "id":      "fase1b",
        "fase":    "Fase 1b",
        "titulo":  "Catálogo de Transformadores",
        "script":  "loc_names_xfo.py",
        "salida":  lambda p: p.loc_xfo,
        "desc":    "Catálogo de transformadores HV/MV/LV con tensiones nominales y potencia.",
        "tags":    ["HV/MV/LV", "Tensión nominal", "Potencia"],
    },
    {
        "id":      "fase1c",
        "fase":    "Fase 1c",
        "titulo":  "Catálogo de Líneas",
        "script":  "loc_namesLineas.py",
        "salida":  lambda p: p.loc_lineas,
        "desc":    "Catálogo de líneas con nombre descriptivo y nivel de tensión.",
        "tags":    ["Líneas", "Nivel de tensión"],
    },
    {
        "id":      "fase1d",
        "fase":    "Fase 1d",
        "titulo":  "Mapeo de Retiros STI",
        "script":  "MapeoRetirosSTI_v6.py",
        "salida":  lambda p: p.loc_cargas,
        "desc":    "Mapea ElmLod → distribuidores STI con 7 prioridades y curvas características.",
        "tags":    ["7 prioridades", "Curvas", "Distribuidores"],
    },
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _estado_paso(paso: dict) -> tuple[str, str]:
    """Retorna (estado_texto, color) según si el archivo de salida existe."""
    salida: Path = paso["salida"](paths)
    if salida.exists():
        mtime = time.strftime("%d/%m/%Y %H:%M", time.localtime(salida.stat().st_mtime))
        return f"✅ Listo — {mtime}", "#2e7d32"
    return "⏳ Pendiente", "#9e9e9e"


def _preview_xlsx(path: Path, max_rows: int = 20) -> None:
    """Muestra las primeras filas de cada hoja de un Excel."""
    try:
        xl = pd.ExcelFile(str(path))
        for sheet in xl.sheet_names[:4]:
            df = xl.parse(sheet, nrows=max_rows)
            st.caption(f"Hoja: **{sheet}** — {len(df)} filas (máx {max_rows})")
            st.dataframe(df, use_container_width=True, height=200)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")


def _run_script(script_name: str, log_placeholder) -> bool:
    """Ejecuta un script .py como subproceso y muestra stdout en tiempo real."""
    script_path = SCRIPTS_1VEZ / script_name
    if not script_path.exists():
        log_placeholder.error(f"Script no encontrado: {script_path}")
        return False
    try:
        proc = subprocess.Popen(
            [sys.executable, str(script_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            cwd=str(SCRIPTS_1VEZ),
        )
        lines: list[str] = []
        for line in proc.stdout:
            lines.append(line.rstrip())
            log_placeholder.code("\n".join(lines[-60:]), language="")
        proc.wait()
        return proc.returncode == 0
    except Exception as e:
        log_placeholder.error(str(e))
        return False


# ── Cabecera ──────────────────────────────────────────────────────────────────
st.title("⚙️ Setup del Modelo")
st.caption(
    "Ejecutar **una sola vez** cuando cambie el modelo PowerFactory o la lista "
    "de generadores/distribuidores."
)

# Botón global "Ejecutar todo"
col_btn, col_info = st.columns([1, 3])
with col_btn:
    run_all = st.button(
        "▶ Ejecutar Fase 0 + Fase 1 completa",
        type="primary",
        use_container_width=True,
        key="btn_run_all",
    )
with col_info:
    st.info(
        "Los scripts de Fase 0 y 1 requieren PowerFactory instalado y licencia activa. "
        "Asegúrate de que PF esté cerrado antes de ejecutar.",
        icon="⚠️",
    )

st.divider()

# ── Stepper: un card por paso ─────────────────────────────────────────────────
for paso in PASOS:
    estado_txt, estado_color = _estado_paso(paso)
    salida_path: Path = paso["salida"](paths)

    with st.container(border=True):
        hdr_col, estado_col = st.columns([3, 1])
        with hdr_col:
            st.markdown(
                f"**{paso['fase']} — {paso['titulo']}**  \n"
                f"`{paso['script']}`",
            )
            st.caption(paso["desc"])
            tags_html = " ".join(
                f'<span style="background:#E3F2FD;color:#1565C0;padding:2px 8px;'
                f'border-radius:10px;font-size:11px;font-weight:600;">{t}</span>'
                for t in paso["tags"]
            )
            st.markdown(tags_html, unsafe_allow_html=True)

        with estado_col:
            st.markdown(
                f'<div style="text-align:right;color:{estado_color};'
                f'font-weight:600;font-size:13px;padding-top:6px;">{estado_txt}</div>',
                unsafe_allow_html=True,
            )

        btn_col, link_col = st.columns([1, 2])
        with btn_col:
            ejecutar = st.button(
                f"▶ Ejecutar {paso['fase']}",
                key=f"btn_{paso['id']}",
                use_container_width=True,
            )
        with link_col:
            if salida_path.exists():
                with open(salida_path, "rb") as f:
                    st.download_button(
                        label=f"⬇ Descargar {salida_path.name}",
                        data=f.read(),
                        file_name=salida_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{paso['id']}",
                        use_container_width=True,
                    )

        if ejecutar or (run_all):
            log_box = st.empty()
            with st.spinner(f"Ejecutando {paso['script']}…"):
                ok = _run_script(paso["script"], log_box)
            if ok:
                st.success(f"✅ {paso['titulo']} completado.", icon="✅")
            else:
                st.error(
                    f"❌ Error al ejecutar {paso['script']}. "
                    "Revisa el log y verifica que PowerFactory esté disponible.",
                )
            # Rompe el run_all si falla
            if not ok and run_all:
                st.warning("⏹ Ejecución en cadena detenida por error en el paso anterior.")
                run_all = False  # type: ignore[assignment]

        # Preview del archivo de salida
        if salida_path.exists():
            with st.expander(f"👁 Vista previa de {salida_path.name}"):
                _preview_xlsx(salida_path)

    st.write("")  # espaciado visual entre pasos


# ── Resumen del estado actual ─────────────────────────────────────────────────
st.divider()
st.subheader("Resumen de artefactos")
data_resumen = []
for paso in PASOS:
    salida = paso["salida"](paths)
    existe = salida.exists()
    mtime  = (
        time.strftime("%d/%m/%Y %H:%M", time.localtime(salida.stat().st_mtime))
        if existe else "—"
    )
    size = f"{salida.stat().st_size / 1024:.1f} KB" if existe else "—"
    data_resumen.append({
        "Fase":   paso["fase"],
        "Script": paso["script"],
        "Salida": salida.name,
        "Estado": "✅ Existe" if existe else "❌ Falta",
        "Última modificación": mtime,
        "Tamaño": size,
    })

st.dataframe(
    pd.DataFrame(data_resumen),
    use_container_width=True,
    hide_index=True,
)
