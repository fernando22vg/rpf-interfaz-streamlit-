#!/usr/bin/env python
"""
verify_config_keys.py - Verificar que todas las claves de configuración estén presentes
"""
from graph_config import DEFAULT_GRAPH_CONFIG

# Claves esperadas (según el código de interfaz_analisis_RPF.py)
REQUIRED_KEYS = {
    # Colores reales
    "freq_color_real",
    "pot_color_real",
    
    # Colores simulados
    "freq_color_simulated",
    "pot_color_simulated",
    "freq_color_sim0",
    "pot_color_sim0",
    "freq_color_sim1",
    "pot_color_sim1",
    
    # Estilos
    "line_width",
    "marker_size",
    
    # Opciones de visualización
    "show_initial",
    "show_nadir",
    "show_dt_eval",
    "show_deadband",
    "show_grid",
    
    # Layout
    "template",
    "plot_height",
    "legend_position",
}

print("Verificando claves de configuración...")
print("=" * 60)

missing_keys = REQUIRED_KEYS - set(DEFAULT_GRAPH_CONFIG.keys())
if missing_keys:
    print(f"❌ Claves FALTANTES: {missing_keys}")
    exit(1)

extra_keys = set(DEFAULT_GRAPH_CONFIG.keys()) - REQUIRED_KEYS
if extra_keys:
    print(f"⚠️  Claves adicionales: {extra_keys}")

print(f"✅ Todas las claves requeridas están presentes")
print(f"✅ Total de claves: {len(DEFAULT_GRAPH_CONFIG)}")
print("\nDetalles:")
for key, value in sorted(DEFAULT_GRAPH_CONFIG.items()):
    print(f"  • {key}: {value}")

print("\n" + "=" * 60)
print("✅ Verificación completada exitosamente")
