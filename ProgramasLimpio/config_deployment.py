"""
config_deployment.py
Configuración adaptada para despliegue en Streamlit Cloud.
- Detecta automáticamente si está en nube o local
- Usa variables de entorno para rutas sensibles
- Proporciona fallbacks seguros
"""

import os
import sys
import json
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# DETECCIÓN DE ENTORNO
# ─────────────────────────────────────────────────────────────────────────────

IS_CLOUD = "STREAMLIT_CLOUD" in os.environ or os.getenv("ENVIRONMENT") == "cloud"
SCRIPT_DIR = Path(__file__).parent.absolute()

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES DE CARPETAS (Relativas al proyecto)
# ─────────────────────────────────────────────────────────────────────────────

CARPETA_COBEE_EMF = "Resultados_COBEE"
CARPETA_DATOS_CURVAS = "Datos Curvas"
CARPETA_COSTO_MARGINAL = "Costo Marginal STI"
CARPETA_DATOS_FRECUENCIAS = "Datos_Frecuencias"

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE RUTAS — LOCAL vs. CLOUD
# ─────────────────────────────────────────────────────────────────────────────

if IS_CLOUD:
    # En Streamlit Cloud, usar rutas relativas o variables de entorno
    RAIZ = os.getenv("RAIZ", str(SCRIPT_DIR))
    RAIZ_DATOS = os.getenv("RAIZ_DATOS", str(SCRIPT_DIR))
    PF_BASE = os.getenv("PF_BASE", "")
    LOC_NAMES_GEN_PATH = os.getenv("LOC_NAMES_GEN_PATH", "")
    LOC_CAR_PATH = os.getenv("LOC_CAR_PATH", "")
    LOC_XFO_PATH = os.getenv("LOC_XFO_PATH", "")
else:
    # Local: intentar cargar config_rutas.json
    config_path = SCRIPT_DIR / "config_rutas.json"
    
    defaults = {
        "RAIZ": r"C:\Datos del CNDC\01_INFO CNDC_RPF",
        "RAIZ_DATOS": r"C:\Datos del CNDC\02_DATOS CNDC_RPF",
        "PF_BASE": r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2",
        "LOC_NAMES_GEN_PATH": r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Designacion de loc_name\loc_names_gen.xlsx",
        "LOC_CAR_PATH": r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Designacion de loc_name\loc_name_cargas.xlsx",
        "LOC_XFO_PATH": r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Designacion de loc_name\loc_names_xfo.xlsx",
    }
    
    if config_path.exists():
        try:
            with open(config_path, encoding="utf-8") as f:
                config = json.load(f)
                RAIZ = config.get("RAIZ", defaults["RAIZ"])
                RAIZ_DATOS = config.get("RAIZ_DATOS", defaults["RAIZ_DATOS"])
                PF_BASE = config.get("PF_BASE", defaults["PF_BASE"])
                LOC_NAMES_GEN_PATH = config.get("LOC_NAMES_GEN_PATH", defaults["LOC_NAMES_GEN_PATH"])
                LOC_CAR_PATH = config.get("LOC_CAR_PATH", defaults["LOC_CAR_PATH"])
                LOC_XFO_PATH = config.get("LOC_XFO_PATH", defaults["LOC_XFO_PATH"])
        except Exception as e:
            print(f"⚠ Error cargando config_rutas.json: {e}")
            RAIZ = defaults["RAIZ"]
            RAIZ_DATOS = defaults["RAIZ_DATOS"]
            PF_BASE = defaults["PF_BASE"]
            LOC_NAMES_GEN_PATH = defaults["LOC_NAMES_GEN_PATH"]
            LOC_CAR_PATH = defaults["LOC_CAR_PATH"]
            LOC_XFO_PATH = defaults["LOC_XFO_PATH"]
    else:
        RAIZ = defaults["RAIZ"]
        RAIZ_DATOS = defaults["RAIZ_DATOS"]
        PF_BASE = defaults["PF_BASE"]
        LOC_NAMES_GEN_PATH = defaults["LOC_NAMES_GEN_PATH"]
        LOC_CAR_PATH = defaults["LOC_CAR_PATH"]
        LOC_XFO_PATH = defaults["LOC_XFO_PATH"]

# ─────────────────────────────────────────────────────────────────────────────
# VALIDACIÓN DE RUTAS (solo alertas, no fallos)
# ─────────────────────────────────────────────────────────────────────────────

def validar_rutas():
    """Retorna lista de rutas no disponibles (para debugging)."""
    no_disponibles = []
    
    if not IS_CLOUD:
        # Solo validar si es local
        if RAIZ and not os.path.isdir(RAIZ):
            no_disponibles.append(f"RAIZ: {RAIZ}")
        if RAIZ_DATOS and not os.path.isdir(RAIZ_DATOS):
            no_disponibles.append(f"RAIZ_DATOS: {RAIZ_DATOS}")
        if PF_BASE and not os.path.isdir(PF_BASE):
            no_disponibles.append(f"PF_BASE: {PF_BASE}")
    
    return no_disponibles

# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES AUXILIARES
# ─────────────────────────────────────────────────────────────────────────────

def get_carpeta_relativa(nombre_carpeta):
    """Obtiene ruta relativa a la carpeta del script."""
    return SCRIPT_DIR / nombre_carpeta

def es_cloud():
    """Retorna True si se ejecuta en Streamlit Cloud."""
    return IS_CLOUD

# ─────────────────────────────────────────────────────────────────────────────
# EXPORTAR CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────────────────

CONFIG = {
    "IS_CLOUD": IS_CLOUD,
    "SCRIPT_DIR": str(SCRIPT_DIR),
    "RAIZ": RAIZ,
    "RAIZ_DATOS": RAIZ_DATOS,
    "PF_BASE": PF_BASE,
    "LOC_NAMES_GEN_PATH": LOC_NAMES_GEN_PATH,
    "LOC_CAR_PATH": LOC_CAR_PATH,
    "LOC_XFO_PATH": LOC_XFO_PATH,
    "CARPETA_COBEE_EMF": CARPETA_COBEE_EMF,
    "CARPETA_DATOS_CURVAS": CARPETA_DATOS_CURVAS,
    "CARPETA_COSTO_MARGINAL": CARPETA_COSTO_MARGINAL,
    "CARPETA_DATOS_FRECUENCIAS": CARPETA_DATOS_FRECUENCIAS,
}

if __name__ == "__main__":
    import pprint
    print(f"Entorno: {'CLOUD' if IS_CLOUD else 'LOCAL'}")
    print("\nConfiguración actual:")
    pprint.pprint(CONFIG)
    
    print("\nRutas no disponibles:")
    no_disp = validar_rutas()
    if no_disp:
        for r in no_disp:
            print(f"  ⚠ {r}")
    else:
        print("  ✓ Todas las rutas están disponibles")
