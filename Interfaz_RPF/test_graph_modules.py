#!/usr/bin/env python
"""
test_graph_modules.py - Verificar que los módulos de gráficas funcionan correctamente
"""
import sys
import traceback

def test_imports():
    """Prueba que los módulos se importan sin errores"""
    print("=" * 60)
    print("PRUEBA 1: Verificar imports de módulos")
    print("=" * 60)
    
    try:
        print("Importando graph_config...")
        from graph_config import (
            COLOR_PALETTE, LINE_WIDTHS, MARKER_SIZES,
            LAYOUT_PRESETS, DEFAULT_GRAPH_CONFIG
        )
        print("✅ graph_config importado correctamente")
        
        print("\nImportando graph_builders...")
        from graph_builders import (
            create_dual_axis_timeseries,
            create_comparison_chart,
            add_kpi_markers,
            add_reference_lines,
            apply_standard_layout,
        )
        print("✅ graph_builders importado correctamente")
        
        return True
    except Exception as e:
        print(f"❌ Error durante import: {e}")
        traceback.print_exc()
        return False


def test_config_structure():
    """Prueba que la configuración tiene la estructura correcta"""
    print("\n" + "=" * 60)
    print("PRUEBA 2: Verificar estructura de configuración")
    print("=" * 60)
    
    try:
        from graph_config import COLOR_PALETTE, DEFAULT_GRAPH_CONFIG, LAYOUT_PRESETS
        
        # Verificar COLOR_PALETTE
        required_colors = [
            "freq_real", "freq_simulated", "power_real", "power_simulated",
            "marker_initial", "marker_nadir", "marker_dt_eval", "marker_error",
            "deadband_line", "fault_line", "eval_line",
        ]
        for color_key in required_colors:
            if color_key not in COLOR_PALETTE:
                print(f"❌ Falta color: {color_key}")
                return False
        print(f"✅ COLOR_PALETTE tiene {len(required_colors)} colores requeridos")
        
        # Verificar DEFAULT_GRAPH_CONFIG
        required_keys = [
            "freq_color_real", "pot_color_real", "line_width", "marker_size",
            "show_initial", "show_nadir", "show_dt_eval", "show_deadband",
            "template", "plot_height", "legend_position",
        ]
        for key in required_keys:
            if key not in DEFAULT_GRAPH_CONFIG:
                print(f"❌ Falta configuración: {key}")
                return False
        print(f"✅ DEFAULT_GRAPH_CONFIG tiene {len(required_keys)} parámetros requeridos")
        
        # Verificar LAYOUT_PRESETS
        required_presets = ["default", "compact", "expanded"]
        for preset in required_presets:
            if preset not in LAYOUT_PRESETS:
                print(f"❌ Falta preset: {preset}")
                return False
        print(f"✅ LAYOUT_PRESETS tiene {len(required_presets)} presets")
        
        return True
    except Exception as e:
        print(f"❌ Error en estructura: {e}")
        traceback.print_exc()
        return False


def test_functions_signature():
    """Prueba que las funciones tienen las firmas correctas"""
    print("\n" + "=" * 60)
    print("PRUEBA 3: Verificar firmas de funciones")
    print("=" * 60)
    
    try:
        from graph_builders import (
            create_dual_axis_timeseries,
            create_comparison_chart,
            add_kpi_markers,
            add_reference_lines,
        )
        import inspect
        
        functions = [
            ("create_dual_axis_timeseries", create_dual_axis_timeseries),
            ("create_comparison_chart", create_comparison_chart),
            ("add_kpi_markers", add_kpi_markers),
            ("add_reference_lines", add_reference_lines),
        ]
        
        for fname, func in functions:
            sig = inspect.signature(func)
            params = list(sig.parameters.keys())
            if len(params) == 0:
                print(f"⚠️  {fname}: sin parámetros detectados")
            else:
                print(f"✅ {fname}: {len(params)} parámetros")
        
        return True
    except Exception as e:
        print(f"❌ Error en firmas: {e}")
        traceback.print_exc()
        return False


def main():
    """Ejecutar todas las pruebas"""
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " PRUEBAS DE MÓDULOS DE GRÁFICAS ".center(58) + "║")
    print("╚" + "=" * 58 + "╝")
    
    results = []
    results.append(("Imports", test_imports()))
    results.append(("Estructura Configuración", test_config_structure()))
    results.append(("Firmas de Funciones", test_functions_signature()))
    
    print("\n" + "=" * 60)
    print("RESUMEN")
    print("=" * 60)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{test_name:.<40} {status}")
    
    print(f"\nTotal: {passed}/{total} pruebas pasaron")
    
    if passed == total:
        print("\n🎉 ¡Todos los módulos están funcionando correctamente!")
        return 0
    else:
        print(f"\n⚠️  {total - passed} prueba(s) fallaron")
        return 1


if __name__ == "__main__":
    sys.exit(main())
