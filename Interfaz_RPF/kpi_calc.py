"""
kpi_calc.py
-----------
Funciones puras de cálculo de KPIs CNDC (f_min, ROCOF, ΔP%, droop, etc.) y de
lookup de Pmax/Rp por unidad, movidas desde interfaz_analisis_RPF.py para que
puedan importarse sin arrastrar la app Streamlit completa (que ejecuta código
de página a nivel de módulo y no es segura de importar).

Sin dependencias de Streamlit ni de st.session_state — solo pandas/numpy/json/os.
Usado por interfaz_analisis_RPF.py (Bloques 3/4/5) y por bloque_dsl_params.py
(KPIs automáticos de experimentos DSL).
"""

from __future__ import annotations
import glob as _glob
import json
import os

import numpy as np
import pandas as pd

# ── Rango válido de frecuencia de red (sistema 50 Hz) ─────────────────────────
FREQ_MIN_HZ, FREQ_MAX_HZ, FREQ_RANGE_MAX_HZ = 45.0, 55.0, 10.0


# ── Lookup de Pmax por unidad ──────────────────────────────────────────────────

def _load_tech_map(path):
    """Carga P_max y Tecnología por loc_name desde Detalle_PF de loc_names_gen.xlsx."""
    try:
        df = pd.read_excel(path, sheet_name="Detalle_PF", engine="calamine")
        pot_cols = [c for c in df.columns if any(kw in c.lower() for kw in ['p_max', 'p nom', 'potencia'])]
        col_p = pot_cols[0] if pot_cols else 'P nom. (MW)'
        tec_cols = [c for c in df.columns if any(kw in str(c).lower() for kw in ['tecnol', 'tipo', 'technology', 'tech'])]
        result = df.set_index('loc_name PF')[[col_p]].rename(columns={col_p: 'P_max (MW)'})
        if tec_cols:
            result['Tecnología'] = df.set_index('loc_name PF')[tec_cols[0]].values
        else:
            result['Tecnología'] = 'Hidroeléctrica'
        return result.to_dict('index')
    except Exception:
        return {}


def _load_pmax_cargado(ev_path, n_evento):
    """Lee Pmax_MW de pgini_GEN_FINAL del Excel de cargado PF.

    Siempre usa pgini_GEN_FINAL (la Pmax no cambia con el ajuste post-LF).
    Devuelve dict {loc_name_pf: pmax_mw}.
    """
    candidates = sorted(
        _glob.glob(os.path.join(ev_path, f"datos_cargados_Ev{n_evento}*.xlsx")),
        key=os.path.getmtime, reverse=True,
    )
    for path in candidates:
        try:
            xl = pd.ExcelFile(path, engine="calamine")
            if "pgini_GEN_FINAL" not in xl.sheet_names:
                continue
            df = xl.parse("pgini_GEN_FINAL")
            if "loc_name PF" not in df.columns or "Pmax_MW" not in df.columns:
                continue
            df["loc_name PF"] = df["loc_name PF"].astype(str).str.strip()
            result = {}
            for _, row in df.iterrows():
                try:
                    v = float(row["Pmax_MW"])
                    if v > 0:
                        result[row["loc_name PF"]] = v
                except (ValueError, TypeError):
                    pass
            if result:
                return result
        except Exception:
            continue
    return {}


def _resolver_unit_key(name, lookup_dict):
    """Versión optimizada de resolución de claves."""
    bare_name = os.path.splitext(name)[0].replace("sym_", "").upper()

    candidates = {bare_name, f"SYM_{bare_name}", bare_name.lower(), f"sym_{bare_name.lower()}"}
    for c in candidates:
        if c in lookup_dict:
            return c, True

    for key in lookup_dict:
        k_norm = key.replace("sym_", "").replace("SYM_", "").upper()
        if bare_name in k_norm or k_norm in bare_name:
            return key, True

    return bare_name, False


def _find_pmax_time(t_arr, pot_arr, t_max_eval, t_min_eval=0.0):
    """Tiempo y potencia máxima en [t_min_eval, t_max_eval] relativo a t₀.

    Por defecto busca solo en la ventana post-falla [0, t_max_eval].
    """
    t_arr   = np.asarray(t_arr,   dtype=float)
    pot_arr = np.asarray(pot_arr, dtype=float)
    mask = (t_arr >= t_min_eval) & (t_arr <= t_max_eval)
    if not np.any(mask):
        return None, None
    t_sub, p_sub = t_arr[mask], pot_arr[mask]
    valid = np.isfinite(p_sub)
    if not np.any(valid):
        return None, None
    idx = int(np.argmax(p_sub[valid]))
    return float(t_sub[valid][idx]), float(p_sub[valid][idx])


def _get_pmax(tdat):
    """Extrae P_max de un registro de tech_map."""
    v = tdat.get('P_max (MW)', 100.0)
    try:
        v = float(v)
        return v if v > 0 else 100.0
    except Exception:
        return 100.0


def _get_pmax_from_cargado(unit_name, pmax_cargado, tech_map, fallback=100.0):
    """Obtiene P_max buscando primero en datos_cargados y luego en tech_map."""
    tk, found = _resolver_unit_key(unit_name, pmax_cargado)
    if found:
        return pmax_cargado[tk], tk, "datos_cargados"
    tk, found = _resolver_unit_key(unit_name, tech_map)
    if found:
        return _get_pmax(tech_map[tk]), tk, "loc_names_gen"
    return fallback, os.path.splitext(unit_name)[0], None


# ── Lookup de Rp (estatismo/droop) por unidad ──────────────────────────────────

def _rp_cfg_path(loc_gen_path):
    return os.path.join(os.path.dirname(loc_gen_path), "estatismo_config.json")


def _load_rp_cfg(loc_gen_path):
    p = _rp_cfg_path(loc_gen_path)
    try:
        if os.path.isfile(p):
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _get_rp_default(loc_key, loc_gen_path, fallback=10.0):
    cfg = _load_rp_cfg(loc_gen_path)
    v = cfg.get(loc_key, cfg.get(loc_key.replace("sym_", ""), None))
    return float(v) if v is not None else fallback


# ── KPIs CNDC ───────────────────────────────────────────────────────────────────

def _cndc_kpis(t_arr, freq_arr, pot_arr, p_max, rp, delta_t, f_nom=50.0):
    """KPIs según metodología oficial CNDC RPF.

    Puntos: t₀ (f₀, P₀), nadir (f_min, t_min), t₀+Δt (f_Δt, P_Δt).
    ΔP = P_Δt − P₀ ; ΔP% = ΔP/P_max×100 ; Aporta si ΔP% ≥ 1.5 %.
    Droop = (Δf'/f_nom) / (ΔP/P_max) × 100, con banda muerta ±25 mHz.
    """
    if len(freq_arr) == 0 or p_max <= 0:
        return None
    _abs_t = np.abs(t_arr.astype(float))
    if np.all(np.isnan(_abs_t)):
        return None
    idx_t0 = int(np.nanargmin(_abs_t))
    f0 = float(freq_arr[idx_t0])
    p0 = float(pot_arr[idx_t0])
    mask_post = t_arr >= 0
    if not np.any(mask_post):
        return None
    t_post, f_post = t_arr[mask_post], freq_arr[mask_post]
    idx_nadir = np.argmin(f_post)
    f_min = float(f_post[idx_nadir])
    t_min = float(t_post[idx_nadir])
    delta_f = f0 - f_min
    idx_dt = np.argmin(np.abs(t_arr - delta_t))
    f_dt = float(freq_arr[idx_dt])
    p_dt = float(pot_arr[idx_dt])

    dp = p_dt - p0
    dp_pct = (dp / p_max) * 100
    r_inic = p_max - p0
    r_inic_pct = (r_inic / p_max) * 100

    f_ref = 49.975 if f0 > 49.975 else f0
    df_prime = f_ref - f_dt
    droop_calc = (df_prime / f_nom) / (dp / p_max) * 100 if abs(dp) > 0.001 else float('nan')

    return {
        'f0': round(f0, 4), 'p0': round(p0, 3),
        'f_min': round(f_min, 4), 't_min': round(t_min, 1), 'delta_f': round(delta_f, 4),
        'f_dt': round(f_dt, 4), 'p_dt': round(p_dt, 3), 't_dt': int(delta_t),
        'r_inic': round(r_inic, 3), 'r_inic_pct': round(r_inic_pct, 2),
        'dp': round(dp, 3), 'dp_pct': round(dp_pct, 2),
        'droop_calc': round(droop_calc, 2) if droop_calc == droop_calc else '—',
        'droop_nom': round(float(rp) * 100, 1),
        'aporta': dp_pct >= 1.5,
    }


def _calcular_rocof(t_arr, f_arr, ventana_s=3.0):
    """ROCOF [Hz/s] por regresión lineal en la ventana [0, ventana_s] post-falla."""
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
        return round(float(coeffs[0]), 4)
    except np.linalg.LinAlgError:
        return float('nan')


# ── Detección de columnas tiempo/frecuencia/potencia ───────────────────────────

def _is_frequency_column(col_name, series_data):
    if "frecuencia" in col_name.lower() or "freq" in col_name.lower() or "hz" in col_name.lower() or "m:f" in col_name.lower():
        return True
    if len(series_data) > 1:
        numeric = pd.to_numeric(series_data, errors='coerce').dropna()
        if len(numeric) > 1:
            min_val = numeric.min()
            max_val = numeric.max()
            if FREQ_MIN_HZ <= min_val <= FREQ_MAX_HZ and FREQ_MIN_HZ <= max_val <= FREQ_MAX_HZ and (max_val - min_val) < FREQ_RANGE_MAX_HZ:
                return True
    return False


def _robust_col_detect(df):
    """Detecta columnas de tiempo, frecuencia y potencia en DataFrames de simulación."""
    cols = df.columns.tolist()
    tc = cols[0]  # Usualmente la primera es el tiempo

    fc_cands = [c for c in cols[1:] if _is_frequency_column(c, df[c])]
    fc_col = fc_cands[0] if fc_cands else cols[1]

    pc_cands = [c for c in cols[1:] if c != fc_col]
    pc_col = pc_cands[0] if pc_cands else (cols[2] if len(cols) > 2 else fc_col)

    return tc, fc_col, pc_col
