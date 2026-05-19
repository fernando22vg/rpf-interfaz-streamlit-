#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ComparativaREAL_SIMU_RMS.py
---------------------------
Consolida y compara datos reales (SCADA COBEE) vs resultados de simulación (PowerFactory).
Aplica criterios de evaluación CNDC: f_min, Droop calculado y umbrales de aporte.
"""

import os
import re
import sys
import glob
import pandas as pd
import numpy as np
from datetime import datetime

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN Y RUTAS
# ─────────────────────────────────────────────────────────────
RAIZ = r"C:\Datos del CNDC\01_INFO CNDC_RPF"
LOC_GEN_PATH = (r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT"
                r"\Designacion de loc_name\loc_names_gen.xlsx")

# Parámetros del sistema
F_NOM = 50.0
BANDA_MUERTA_HZ = 0.025  # ±25 mHz (CDM)
UMBRAL_APORTE_PCT = 1.5  # 1.5% Pmax

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def elegir(opciones, titulo):
    print(f"\n{titulo}:")
    for i, op in enumerate(opciones, 1):
        print(f"  {i}. {op}")
    while True:
        try:
            sel = int(input("  Seleccionar número: "))
            if 1 <= sel <= len(opciones):
                return opciones[sel - 1]
        except ValueError: pass
        print("  Opción inválida.")

def _get_kpis(df, unit_col, is_sim=False):
    """Extrae f0, f_min, P0, P_max_val y calcula deltas según criterios CNDC."""
    # Detectar columnas de frecuencia
    f_col_cand = [c for c in df.columns if 'frec' in c.lower() or 'hz' in c.lower()]
    f_col = f_col_cand[0] if f_col_cand else df.columns[1]
    
    p_col = unit_col
    
    # Asegurar datos numéricos
    df[f_col] = pd.to_numeric(df[f_col], errors='coerce')
    df[p_col] = pd.to_numeric(df[p_col], errors='coerce')
    df = df.dropna(subset=[f_col, p_col])

    if df.empty: return None

    f = df[f_col].values
    p = df[p_col].values
    
    f0 = f[0]
    p0 = p[0]
    f_min = np.min(f)
    idx_nadir = np.argmin(f)
    p_nadir = p[idx_nadir]
    
    # Delta f' (Lógica CNDC para el escalón de frecuencia)
    ref_f = 49.975 if f0 > 49.975 else f0
    delta_f_prime = abs(ref_f - f_min)
    
    # Descontar banda muerta (±25 mHz)
    delta_f_eff = max(0, delta_f_prime - BANDA_MUERTA_HZ)
    delta_p = p_nadir - p0
    
    return {
        'f0': f0, 'f_min': f_min, 'p0': p0, 
        'p_nadir': p_nadir, 'delta_f_eff': delta_f_eff, 'delta_p': delta_p
    }

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():
    print("="*70)
    print("  COMPARATIVA REAL vs SIMULACIÓN RMS (CRITERIOS CNDC)")
    print("="*70)

    # 1. Selección de Semestre y Evento
    if not os.path.isdir(RAIZ):
        print(f"❌ No se encontró la ruta raíz: {RAIZ}")
        return

    semestres = sorted(d for d in os.listdir(RAIZ) if os.path.isdir(os.path.join(RAIZ, d)))
    semestre = elegir(semestres, "Selecciona Semestre")
    
    base_ev = os.path.join(RAIZ, semestre, "Análisis_todos_los_eventos")
    if not os.path.isdir(base_ev):
        base_ev = os.path.join(RAIZ, semestre, "Analisis_todos_los_eventos")

    eventos = sorted(d for d in os.listdir(base_ev) if os.path.isdir(os.path.join(base_ev, d)))
    evento = elegir(eventos, "Selecciona Evento")
    ev_path = os.path.join(base_ev, evento)
    
    m_ev = re.search(r"(\d+)$", evento.strip())
    n_ev = m_ev.group(1) if m_ev else "0"

    # 2. Selección de Simulaciones a comparar
    sim_disponibles = [d for d in os.listdir(ev_path) if d.startswith(f"E{n_ev}.")]
    if not sim_disponibles:
        print("❌ No se encontraron carpetas de simulación (E.0/E.1).")
        return
    
    comparar_con = elegir(sim_disponibles + ["Ambos (E.0 y E.1)"], "Tipo de simulación para comparar")
    sims_to_process = sim_disponibles if "Ambos" in comparar_con else [comparar_con]

    # 3. Cargar Parámetros Técnicos (Pmax, Rp)
    try:
        df_tech = pd.read_excel(LOC_GEN_PATH, sheet_name="Detalle_PF")
        tech_map = df_tech.set_index('loc_name PF')[['P nom. (MW)', 'Estatismo (%)']].to_dict('index')
    except Exception:
        tech_map = {}

    # 4. Procesamiento
    real_dir = os.path.join(ev_path, "Graficas Registro 1SEG COBEE")
    if not os.path.isdir(real_dir):
        print(f"❌ No existe carpeta de datos reales: {real_dir}")
        return

    resultados_finales = []

    for sim_id in sims_to_process:
        sim_dir = os.path.join(ev_path, sim_id, "Datos Curvas")
        if not os.path.isdir(sim_dir): continue
        
        real_files = {os.path.splitext(f)[0]: f for f in os.listdir(real_dir) if f.endswith('.xlsx') and not f.startswith('~$')}
        
        for unit, r_file in real_files.items():
            s_file_match = glob.glob(os.path.join(sim_dir, f"*{unit}*.xlsx"))
            if not s_file_match: continue
            
            try:
                df_r = pd.read_excel(os.path.join(real_dir, r_file))
                df_s = pd.read_excel(s_file_match[0])
                
                tech_key = unit if unit in tech_map else f"sym_{unit}"
                tech_data = tech_map.get(tech_key, {'P nom. (MW)': 100.0, 'Estatismo (%)': 5.0})
                p_max, rp = float(tech_data['P nom. (MW)']), float(tech_data['Estatismo (%)']) / 100.0

                k_r = _get_kpis(df_r, unit)
                p_col_s = [c for c in df_s.columns if unit in c or 'pot' in c.lower()][0]
                k_s = _get_kpis(df_s, p_col_s, is_sim=True)

                if not k_r or not k_s: continue

                droop_r = (k_r['delta_f_eff'] / F_NOM) / (k_r['delta_p'] / p_max) * 100 if abs(k_r['delta_p']) > 0.1 else 0
                droop_s = (k_s['delta_f_eff'] / F_NOM) / (k_s['delta_p'] / p_max) * 100 if abs(k_s['delta_p']) > 0.1 else 0
                aporte_req = (p_max / (rp if rp > 0 else 0.05)) * (k_r['delta_f_eff'] / F_NOM)

                resultados_finales.append({
                    'Unidad': unit, 'Simulación': sim_id,
                    'f_min Real (Hz)': round(k_r['f_min'], 3), 'f_min Sim (Hz)': round(k_s['f_min'], 3),
                    '[A] Validez f_min': "✅ OK" if abs(k_s['f_min'] - k_r['f_min']) <= 0.1 else "❌ FUERA",
                    'Droop Real (%)': round(abs(droop_r), 2), 'Droop Sim (%)': round(abs(droop_s), 2),
                    '[B] Error Droop (%)': round(abs(droop_s - droop_r), 2),
                    '[C] Cumple Aporte': "✅ SI" if ( (k_r['delta_p']/p_max)*100 >= 1.5 and k_r['delta_p'] >= aporte_req) else "⚠️ BAJO"
                })
            except Exception: pass

    if resultados_finales:
        out_path = os.path.join(ev_path, f"Comparativa_REAL_SIMU_Ev{n_ev}.xlsx")
        pd.DataFrame(resultados_finales).to_excel(out_path, index=False)
        print(f"\n✅ Archivo generado: {out_path}")

if __name__ == "__main__":
    main()