"""
diagnostico_rocof_curvas.py — Comparador ROCOF: curva real SCADA vs simulación RMS

Discrimina la causa del retraso del nadir de frecuencia en la simulación:
  · Si la pendiente inicial (1-2 s, solo inercia) simulada es mas plana que la
    real → la inercia H del modelo PF esta sobreestimada.
  · Si las pendientes iniciales coinciden pero el nadir simulado llega tarde →
    el retraso viene de los gobernadores (constantes de tiempo DSL) y/o kpf.

Fuentes (las mismas que usa el Bloque 05 de la interfaz Streamlit):
  Real : {evento}\\Graficas Registro 1SEG COBEE\\*.xlsx  (SCADA COBEE 1SEG)
         fallback {evento}\\Resultados_COBEE\\*.xlsx     (EMF CNDC)
  Sim  : {evento}\\E{n}.0\\Datos Curvas\\F. Barras SIN.xlsx (export ShowData PF)

Uso:
  python diagnostico_rocof_curvas.py --ev "C:\\...\\Evento 1"
  python diagnostico_rocof_curvas.py --ev "..." --sim "ruta\\F. Barras SIN.xlsx"
"""

import os
import re
import sys
import glob
import json
import argparse

import numpy as np
import pandas as pd

# Forzar UTF-8 en stdout para evitar UnicodeEncodeError en consolas cp1252 (Windows)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

UMBRAL_DFDT  = -0.04   # Hz/s — caída sostenida para detectar inicio de falla (default B05)
VENTANA_SUAV = 5       # muestras — rolling mean antes de df/dt (default B05)
F_VALID_MIN, F_VALID_MAX = 45.0, 55.0   # rango físico de frecuencia [Hz]


#  Réplicas de las funciones de la interfaz (interfaz_analisis_RPF.py) 
# No se importan porque el módulo ejecuta UI Streamlit al importarse.

def _parse_to_seconds(series):
    """Convierte tiempo (HH:MM:SS o float) a segundos. Réplica de la interfaz (línea 1161)."""
    s = series.astype(str).str.strip().str.replace(',', '.')
    result = pd.Series(np.nan, index=series.index)
    has_colon = s.str.contains(':')
    if has_colon.any():
        parts = s[has_colon].str.split(':')
        h   = pd.to_numeric(parts.str[0], errors='coerce').fillna(0)
        m   = pd.to_numeric(parts.str[1], errors='coerce').fillna(0)
        sec = pd.to_numeric(parts.str[2], errors='coerce').fillna(0)
        result[has_colon] = h * 3600 + m * 60 + sec
    not_colon = ~has_colon
    if not_colon.any():
        result[not_colon] = pd.to_numeric(s[not_colon], errors='coerce')
    return result


def _detectar_inicio_falla(freq_array, umbral_dfdt=UMBRAL_DFDT, ventana_suavizado=VENTANA_SUAV):
    """Inicio de falla por df/dt sostenido sobre señal suavizada. Réplica (línea 1626)."""
    n = len(freq_array)
    if n < ventana_suavizado + 2:
        return 0
    kernel = np.ones(ventana_suavizado) / ventana_suavizado
    freq_smooth = np.convolve(freq_array.astype(float), kernel, mode='same')
    half = ventana_suavizado // 2
    freq_smooth[:half] = freq_smooth[half]
    freq_smooth[-half:] = freq_smooth[-half - 1]
    dfdt = np.diff(freq_smooth)
    condicion = (dfdt[:-1] < umbral_dfdt) & (dfdt[1:] < umbral_dfdt)
    indices = np.where(condicion)[0]
    return int(indices[0]) if len(indices) else 0


def _calcular_rocof(t_arr, f_arr, ventana_s=3.0):
    """ROCOF [Hz/s] por regresión lineal en [0, ventana_s] post-falla. Réplica (línea 1706)."""
    mask = (t_arr >= 0) & (t_arr <= ventana_s)
    if np.sum(mask) < 2:
        return float('nan')
    t_w, f_w = t_arr[mask], f_arr[mask]
    valid = np.isfinite(t_w) & np.isfinite(f_w)
    t_w, f_w = t_w[valid], f_w[valid]
    if len(t_w) < 2 or np.ptp(t_w) == 0:
        return float('nan')
    try:
        coeffs = np.polyfit(t_w - t_w[0], f_w, 1)
        return float(coeffs[0])
    except np.linalg.LinAlgError:
        return float('nan')


#  Carga de curvas 

def cargar_curva_real(ev_path):
    """Curva real del sistema: mediana entre unidades SCADA, alineada en t_trip=0.

    Retorna (t_aligned, f_arr, n_unidades, fuente).
    Cada unidad se alinea en su propio t0 detectado (absorbe desfases de reloj
    entre registradores) y luego se toma la mediana sobre una grilla común de 1 s.
    """
    candidatos = [
        ("Graficas Registro 1SEG COBEE", "SCADA COBEE (1SEG)"),
        ("Resultados_COBEE",             "EMF CNDC"),
    ]
    carpeta = fuente = None
    for sub, etiqueta in candidatos:
        d = os.path.join(ev_path, sub)
        if os.path.isdir(d):
            carpeta, fuente = d, etiqueta
            break
    if carpeta is None:
        sys.exit("[ERROR] No existe 'Graficas Registro 1SEG COBEE' ni 'Resultados_COBEE' "
                 "en el evento.\n        Corra primero el Bloque 03 (ExtractorResultadosCNDC).")

    archivos = [f for f in glob.glob(os.path.join(carpeta, "*.xlsx"))
                if not os.path.basename(f).lower().startswith("tabla")]
    if not archivos:
        sys.exit(f"[ERROR] Sin archivos .xlsx de unidades en: {carpeta}\n"
                 "        Corra primero el Bloque 03 (ExtractorResultadosCNDC).")

    curvas = []   # [(t_aligned, f_arr, nombre)]
    for fp in sorted(archivos):
        nombre = os.path.splitext(os.path.basename(fp))[0]
        try:
            df = pd.read_excel(fp).dropna()
            tr_raw  = _parse_to_seconds(df.iloc[:, 0])
            fr_cols = [c for c in df.columns
                       if any(kw in str(c).lower() for kw in ("frec", "hz", "freq"))]
            fr_col  = fr_cols[0] if fr_cols else df.columns[1]
            fr_arr  = pd.to_numeric(df[fr_col], errors="coerce").ffill().values
            tr_norm = (tr_raw - tr_raw.min()).values
            valid   = (np.isfinite(tr_norm) & np.isfinite(fr_arr)
                       & (fr_arr > F_VALID_MIN) & (fr_arr < F_VALID_MAX))
            tr_norm, fr_arr = tr_norm[valid], fr_arr[valid]
            if len(fr_arr) < 10:
                continue
            idx_f = _detectar_inicio_falla(fr_arr)
            if idx_f <= 0:
                print(f"  [WARN] {nombre}: t0 no detectado — unidad omitida")
                continue
            curvas.append((tr_norm - tr_norm[idx_f], fr_arr, nombre))
        except Exception as e:
            print(f"  [WARN] {nombre}: error de lectura ({e}) — omitida")

    if not curvas:
        sys.exit("[ERROR] Ninguna curva real con t0 detectable.")

    # Grilla común a 1 s; mediana entre unidades
    t_min = max(c[0].min() for c in curvas)
    t_max = min(c[0].max() for c in curvas)
    grid  = np.arange(np.ceil(t_min), np.floor(t_max) + 0.5, 1.0)
    apiladas = np.vstack([np.interp(grid, c[0], c[1]) for c in curvas])
    f_median = np.median(apiladas, axis=0)
    print(f"  Curva real: {len(curvas)} unidades ({fuente}): "
          + ", ".join(c[2] for c in curvas))
    return grid, f_median, len(curvas), fuente


def cargar_curva_sim(ev_path, sim_path=None):
    """Curva simulada: mediana de las barras de 'F. Barras SIN.xlsx', alineada en t_trip=0.

    Retorna (t_aligned, f_arr, archivo).
    """
    if sim_path is None:
        patrones = glob.glob(os.path.join(ev_path, "E*", "Datos Curvas", "F. Barras SIN.xlsx"))
        if not patrones:
            sys.exit("[ERROR] No se encontró 'F. Barras SIN.xlsx' en E*\\Datos Curvas.\n"
                     "        Exporte las frecuencias de barras desde PF (ShowData) o use --sim.")
        sim_path = sorted(patrones)[0]
    if not os.path.isfile(sim_path):
        sys.exit(f"[ERROR] No existe: {sim_path}")

    raw = pd.read_excel(sim_path, header=None)
    # Fila 0: nombres de barras | Fila 1: unidades | datos desde fila 2
    datos = raw.iloc[2:].apply(pd.to_numeric, errors="coerce").dropna(how="all")
    t_arr = datos.iloc[:, 0].values.astype(float)
    f_mat = datos.iloc[:, 1:].values.astype(float)
    # Normalizar pu→Hz si el export viene en por-unidad
    if np.nanmax(f_mat) < 2.0:
        f_mat = f_mat * 50.0
    f_arr = np.nanmedian(f_mat, axis=1)
    valid = np.isfinite(t_arr) & np.isfinite(f_arr)
    t_arr, f_arr = t_arr[valid], f_arr[valid]

    idx_f = _detectar_inicio_falla(f_arr)
    if idx_f > 0:
        t_trip = t_arr[idx_f]
    else:
        # Fallback: t_sim_falla guardado por la interfaz en event_config.json
        t_trip = 5.0
        cfg_path = os.path.join(ev_path, "event_config.json")
        if os.path.isfile(cfg_path):
            try:
                with open(cfg_path, encoding="utf-8") as fh:
                    t_trip = float(json.load(fh).get("t_sim_falla", 5.0))
            except Exception:
                pass
        print(f"  [WARN] t0 sim no detectado por pendiente — usando t_sim_falla={t_trip}s")

    barras = [str(b) for b in raw.iloc[0, 1:].dropna().tolist()]
    print(f"  Curva sim : {len(barras)} barras ({', '.join(barras)})  ←  {os.path.basename(sim_path)}")
    return t_arr - t_trip, f_arr, sim_path


#  Métricas 

def metricas_curva(t, f):
    """f0, ROCOF (1/2/3 s), nadir, t_nadir, f@+30s — todo relativo a t_trip=0."""
    idx0 = int(np.nanargmin(np.abs(t)))
    f0 = float(f[idx0])
    post = (t >= 0) & (t <= 60)
    t_p, f_p = t[post], f[post]
    i_nad = int(np.nanargmin(f_p))
    f30 = float(f_p[np.nanargmin(np.abs(t_p - 30.0))]) if np.any(t_p >= 30) else float('nan')
    return {
        "f0_hz":      f0,
        "rocof_1s":   _calcular_rocof(t, f, 1.0),
        "rocof_2s":   _calcular_rocof(t, f, 2.0),
        "rocof_3s":   _calcular_rocof(t, f, 3.0),
        "f_nadir_hz": float(f_p[i_nad]),
        "t_nadir_s":  float(t_p[i_nad]),
        "delta_f_hz": f0 - float(f_p[i_nad]),
        "f_30s_hz":   f30,
    }


def veredicto(m_real, m_sim):
    """Diagnóstico automático según ratio de ROCOF y de t_nadir."""
    lineas = []
    roc_r, roc_s = abs(m_real["rocof_2s"]), abs(m_sim["rocof_2s"])
    tn_r,  tn_s  = m_real["t_nadir_s"], m_sim["t_nadir_s"]
    ratio_roc = roc_r / roc_s if roc_s > 1e-6 else float('inf')
    ratio_tn  = tn_s / tn_r   if tn_r  > 1e-6 else float('inf')

    if ratio_roc > 1.3:
        lineas.append(
            f"INERCIA: el modelo cae {ratio_roc:.1f}x mas lento en los primeros 2 s.\n"
            f"  La pendiente inicial depende SOLO de la inercia → H_modelo ≈ {ratio_roc:.1f} x H_real.\n"
            f"  Revisar: unidades en servicio que debian estar paradas (aportan H con 0 MW)\n"
            f"  y valores H de los tipos TypSym. Ver diagnostico_inercia en el script de carga.")
    if ratio_roc <= 1.3 and ratio_tn > 1.5:
        lineas.append(
            f"GOBERNADORES/AMORTIGUAMIENTO: la pendiente inicial coincide "
            f"(ROCOF real {roc_r:.3f} vs sim {roc_s:.3f} Hz/s), pero el nadir simulado\n"
            f"  llega {ratio_tn:.1f}x mas tarde ({tn_s:.1f}s vs {tn_r:.1f}s). El retraso viene de la\n"
            f"  respuesta primaria: constantes de tiempo de los gobernadores DSL y/o kpf de cargas.")
    if ratio_roc > 1.3 and ratio_tn > 1.5:
        lineas.append(
            "CASO MIXTO: hay exceso de inercia Y respuesta primaria lenta. Corregir primero\n"
            "  la inercia (unidades en servicio) y re-evaluar el tiempo del nadir.")
    if not lineas:
        lineas.append(
            f"SIN DISCREPANCIA SIGNIFICATIVA: ROCOF real/sim = {ratio_roc:.2f}, "
            f"t_nadir sim/real = {ratio_tn:.2f}. Las curvas son comparables.")
    return "\n\n".join(lineas), ratio_roc, ratio_tn


#  Main 

def main():
    ap = argparse.ArgumentParser(description="Comparador ROCOF: curva real SCADA vs simulación RMS")
    ap.add_argument("--ev",  required=True, help="Carpeta del evento (ej: ...\\Evento 1)")
    ap.add_argument("--sim", help="Ruta a F. Barras SIN.xlsx (default: autodetectar en E*\\Datos Curvas)")
    args = ap.parse_args()

    ev_path = args.ev
    if not os.path.isdir(ev_path):
        sys.exit(f"[ERROR] Carpeta de evento no existe: {ev_path}")
    m = re.search(r"(\d+)\s*$", os.path.basename(ev_path.rstrip("\\/")))
    n_ev = m.group(1) if m else "X"

    print("=" * 66)
    print(f"  DIAGNOSTICO ROCOF — {os.path.basename(ev_path)}")
    print("=" * 66)

    t_r, f_r, n_unid, fuente = cargar_curva_real(ev_path)
    t_s, f_s, sim_path = cargar_curva_sim(ev_path, args.sim)

    m_r = metricas_curva(t_r, f_r)
    m_s = metricas_curva(t_s, f_s)

    filas = [
        ("f0 — inicio evento [Hz]",  f"{m_r['f0_hz']:.4f}",      f"{m_s['f0_hz']:.4f}"),
        ("ROCOF 1s [Hz/s]",          f"{m_r['rocof_1s']:.4f}",   f"{m_s['rocof_1s']:.4f}"),
        ("ROCOF 2s [Hz/s]",          f"{m_r['rocof_2s']:.4f}",   f"{m_s['rocof_2s']:.4f}"),
        ("ROCOF 3s [Hz/s]",          f"{m_r['rocof_3s']:.4f}",   f"{m_s['rocof_3s']:.4f}"),
        ("f_min — nadir [Hz]",       f"{m_r['f_nadir_hz']:.4f}", f"{m_s['f_nadir_hz']:.4f}"),
        ("t_nadir desde trip [s]",   f"{m_r['t_nadir_s']:.1f}",  f"{m_s['t_nadir_s']:.1f}"),
        ("Δf = f0 − f_min [Hz]",     f"{m_r['delta_f_hz']:.4f}", f"{m_s['delta_f_hz']:.4f}"),
        ("f @ +30 s [Hz]",           f"{m_r['f_30s_hz']:.4f}",   f"{m_s['f_30s_hz']:.4f}"),
    ]
    print()
    print(f"  {'Métrica':<28} {'REAL (' + fuente + ')':<22} {'SIMULACIÓN':<14}")
    print(f"  {'' * 28} {'' * 22} {'' * 14}")
    for nom, vr, vs in filas:
        print(f"  {nom:<28} {vr:<22} {vs:<14}")

    texto_v, ratio_roc, ratio_tn = veredicto(m_r, m_s)
    print()
    print("  " + "═" * 62)
    print("  VEREDICTO")
    print("  " + "═" * 62)
    for ln in texto_v.splitlines():
        print(f"  {ln}")
    print()

    #  PNG 
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt

        fig, ax = plt.subplots(figsize=(11, 6))
        ax.plot(t_r, f_r, color="#1a7a4a", lw=1.8, label=f"Real ({fuente}, {n_unid} uds)")
        ax.plot(t_s, f_s, color="#c0392b", lw=1.8, ls="--", label="Simulación RMS")
        ax.scatter([m_r["t_nadir_s"]], [m_r["f_nadir_hz"]], color="#1a7a4a", zorder=5)
        ax.scatter([m_s["t_nadir_s"]], [m_s["f_nadir_hz"]], color="#c0392b", zorder=5)
        ax.annotate(f"nadir real {m_r['f_nadir_hz']:.3f} Hz @ {m_r['t_nadir_s']:.0f}s",
                    (m_r["t_nadir_s"], m_r["f_nadir_hz"]), textcoords="offset points",
                    xytext=(8, -14), fontsize=9, color="#1a7a4a")
        ax.annotate(f"nadir sim {m_s['f_nadir_hz']:.3f} Hz @ {m_s['t_nadir_s']:.0f}s",
                    (m_s["t_nadir_s"], m_s["f_nadir_hz"]), textcoords="offset points",
                    xytext=(8, 10), fontsize=9, color="#c0392b")
        ax.axvline(0, color="gray", lw=0.8, ls=":")
        ax.set_xlim(-10, 60)
        ax.set_xlabel("Tiempo desde el disparo [s]")
        ax.set_ylabel("Frecuencia [Hz]")
        ax.set_title(f"Evento {n_ev} — Real vs Simulación (alineadas en el trip)\n"
                     f"ROCOF 2s: real {m_r['rocof_2s']:.3f} / sim {m_s['rocof_2s']:.3f} Hz/s "
                     f"(ratio {ratio_roc:.2f})")
        ax.grid(alpha=0.3)
        ax.legend()
        png_path = os.path.join(ev_path, f"diagnostico_rocof_Ev{n_ev}.png")
        fig.tight_layout()
        fig.savefig(png_path, dpi=130)
        plt.close(fig)
        print(f"  Gráfica : {png_path}")
    except ImportError:
        print("  [WARN] matplotlib no disponible — se omite el PNG")

    #  XLSX 
    xlsx_path = os.path.join(ev_path, f"diagnostico_rocof_Ev{n_ev}.xlsx")
    df_met = pd.DataFrame(
        [(nom, vr, vs) for nom, vr, vs in filas]
        + [("ratio ROCOF real/sim", f"{ratio_roc:.2f}", ""),
           ("ratio t_nadir sim/real", f"{ratio_tn:.2f}", ""),
           ("veredicto", texto_v.replace("\n", " "), "")],
        columns=["Métrica", "Real", "Simulación"],
    )
    grid = np.arange(-10.0, 60.5, 0.5)
    df_cur = pd.DataFrame({
        "t_s":     grid,
        "f_real":  np.interp(grid, t_r, f_r, left=np.nan, right=np.nan),
        "f_sim":   np.interp(grid, t_s, f_s, left=np.nan, right=np.nan),
    })
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as wr:
        df_met.to_excel(wr, sheet_name="metricas", index=False)
        df_cur.to_excel(wr, sheet_name="curvas_alineadas", index=False)
    print(f"  Resumen : {xlsx_path}")


if __name__ == "__main__":
    main()
