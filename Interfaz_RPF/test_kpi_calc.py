import numpy as np
import pytest

from kpi_calc import _cndc_kpis, _calcular_rocof, _is_frequency_column, _robust_col_detect
import pandas as pd

# 
# TESTS: kpi_calc.py — funciones puras, sin dependencia de Streamlit
# 


def test_cndc_kpis_calculation():
    t = np.array([-1.0, 0.0, 1.0, 35.0])
    f = np.array([50.0, 50.0, 49.5, 49.6])
    p = np.array([10.0, 10.0, 12.0, 15.0])
    p_max = 100.0
    rp = 0.05  # 5%

    kpis = _cndc_kpis(t, f, p, p_max, rp, delta_t=35)

    assert kpis is not None
    assert kpis['f0'] == 50.0
    assert kpis['f_min'] == 49.5
    assert kpis['dp'] == 5.0  # 15MW - 10MW
    assert kpis['dp_pct'] == 5.0  # (5/100)*100
    assert kpis['aporta'] is True  # 5% >= 1.5%


def test_calcular_rocof():
    t = np.array([0, 0.5, 1.0, 1.5])
    f = np.array([50.0, 49.95, 49.90, 49.85])
    # Linear drop of 0.1 Hz per second
    rocof = _calcular_rocof(t, f, ventana_s=2.0)
    assert pytest.approx(rocof, 0.001) == -0.1


def test_robust_col_detect():
    df = pd.DataFrame({
        "Tiempo (s)": [0.0, 0.5, 1.0],
        "Potencia (MW)": [10.0, 10.5, 11.0],
        "Frecuencia (Hz)": [50.0, 49.9, 49.8],
    })
    tc, fc, pc = _robust_col_detect(df)
    assert tc == "Tiempo (s)"
    assert fc == "Frecuencia (Hz)"
    assert pc == "Potencia (MW)"


def test_is_frequency_column_by_value_range():
    serie = pd.Series([49.8, 49.9, 50.0, 50.1])
    assert _is_frequency_column("col_sin_nombre", serie) is True


def test_is_frequency_column_rejects_power_range():
    serie = pd.Series([10.0, 50.0, 120.0])
    assert _is_frequency_column("col_sin_nombre", serie) is False
