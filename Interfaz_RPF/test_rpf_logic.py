import pytest
import numpy as np
import pandas as pd
from interfaz_analisis_RPF import _detectar_inicio_falla, _cndc_kpis, _calcular_rocof
from CondInicialesPF import extraer_sti, buscar_p0, _norm_hora, _ajustar_plini_por_distribuidor
from ExtractorResultadosCNDC import _parse_time_s, _interp

# 
# TESTS: Análisis RPF (interfaz_analisis_RPF.py)
# 

def test_detectar_inicio_falla():
    # Frequency dropping starting at index 5
    freq = np.array([50.0, 50.0, 50.0, 50.0, 50.0, 49.9, 49.8, 49.7, 49.6])
    # Using window=1 to simplify detection for the test
    idx = _detectar_inicio_falla(freq, umbral_dfdt=-0.05, ventana_suavizado=1)
    assert idx == 5

def test_cndc_kpis_calculation():
    t = np.array([-1.0, 0.0, 1.0, 35.0])
    f = np.array([50.0, 50.0, 49.5, 49.6])
    p = np.array([10.0, 10.0, 12.0, 15.0])
    p_max = 100.0
    rp = 0.05 # 5%
    
    kpis = _cndc_kpis(t, f, p, p_max, rp, delta_t=35)
    
    assert kpis is not None
    assert kpis['f0'] == 50.0
    assert kpis['f_min'] == 49.5
    assert kpis['dp'] == 5.0 # 15MW - 10MW
    assert kpis['dp_pct'] == 5.0 # (5/100)*100
    assert kpis['aporta'] is True # 5% >= 1.5%

def test_calcular_rocof():
    t = np.array([0, 0.5, 1.0, 1.5])
    f = np.array([50.0, 49.95, 49.90, 49.85])
    # Linear drop of 0.1 Hz per second
    rocof = _calcular_rocof(t, f, ventana_s=2.0)
    assert pytest.approx(rocof, 0.001) == -0.1

# 
# TESTS: Condiciones Iniciales (CondInicialesPF.py)
# 

def test_extraer_sti():
    assert extraer_sti("sym_ZON01") == "ZON01"
    assert extraer_sti("WT_QOL01_EQ") == "QOL01"
    assert extraer_sti("sym_ZON01(1)") == "ZON01"
    assert extraer_sti("sta_GEN01_II") == "GEN01"

def test_buscar_p0():
    dict_p0 = {"TIQ": 10.5, "ZON01": 20.0}
    # Fallback to TIQ from TIQ01
    assert buscar_p0("TIQ01", dict_p0) == 10.5
    assert buscar_p0("ZON01", dict_p0) == 20.0
    assert buscar_p0("UNKNOWN", dict_p0) is None

def test_norm_hora():
    assert _norm_hora("6:21") == "06:21"
    assert _norm_hora("18:45") == "18:45"
    assert _norm_hora("09:00:00") == "09:00"

def test_ajustar_plini_por_distribuidor():
    # Test the deterministic largest remainder algorithm
    df_plini = pd.DataFrame({
        "Distribuidor": ["Dist_A", "Dist_A"],
        "P_nom_MW": [10.0, 10.0],
        "plini_base_MW": [2.5, 2.5]
    })
    # Objective is 5.01, so one charge should get 2.5050 rounded to 2.5051
    dict_plini_dist = {"Dist_A": 5.01}
    
    df_balance, max_residuo = _ajustar_plini_por_distribuidor(df_plini, dict_plini_dist)
    
    assert df_plini["plini_MW"].sum() == pytest.approx(5.01, 0.0001)
    assert max_residuo == 0
    assert 2.505 in df_plini["plini_MW"].values

# 
# TESTS: EMF Parser (ExtractorResultadosCNDC.py)
# 

def test_parse_time_s():
    # HH:MM:SS
    assert _parse_time_s("10:30:15") == 10 * 3600 + 30 * 60 + 15
    # HH:MM
    assert _parse_time_s("14:20") == 14 * 3600 + 20 * 60
    # Bad format
    assert _parse_time_s("Not a time") is None

def test_interp_basic():
    ticks = [(0, 10.0), (100, 20.0)]
    # Pixel 50 should be value 15.0
    assert _interp(50, ticks) == 15.0
    # Pixel 0 should be 10.0
    assert _interp(0, ticks) == 10.0

def test_get_kpis_comparativa():
    # Testing the version in ComparativaREAL_SIMU_RMS.py
    from ComparativaREAL_SIMU_RMS import _get_kpis
    
    df = pd.DataFrame({
        'frecuencia': [50.0, 49.8, 49.7, 49.9],
        'MW': [100.0, 105.0, 110.0, 102.0]
    })
    
    kpis = _get_kpis(df, 'MW')
    assert kpis['f0'] == 50.0
    assert kpis['f_min'] == 49.7
    assert kpis['p0'] == 100.0
    assert kpis['p_nadir'] == 110.0
    assert kpis['delta_p'] == 10.0