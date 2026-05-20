"""
patch_for_streamlit_cloud.py

Script para parchear interfaz_analisis_RPF.py con cambios minimalistas 
que lo hacen compatible con Streamlit Cloud.

USO:
    python patch_for_streamlit_cloud.py

CAMBIOS APLICADOS:
1. Agregar sys.path en línea 13-14 (después del docstring)
2. Importar config_deployment en lugar de usar rutas hardcodeadas
3. Agregar fallback para archivos faltantes en Cloud
"""

import os
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
TARGET_FILE = SCRIPT_DIR / "interfaz_analisis_RPF.py"

# Inyección a agregar después del docstring (líneas 12-13)
PATCH_IMPORTS = '''
# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN PARA STREAMLIT CLOUD (Auto-detecta local/cloud)
# ─────────────────────────────────────────────────────────────────────────────
import sys as _sys_patch
import os as _os_patch
_SCRIPT_DIR_PATCH = _os_patch.path.dirname(_os_patch.path.abspath(__file__))
if _SCRIPT_DIR_PATCH not in _sys_patch.path:
    _sys_patch.path.insert(0, _SCRIPT_DIR_PATCH)
del _sys_patch, _os_patch, _SCRIPT_DIR_PATCH

try:
    from config_deployment import CONFIG, es_cloud, validar_rutas
    _IS_CLOUD = es_cloud()
except ImportError:
    _IS_CLOUD = False
    CONFIG = {}
    def validar_rutas():
        return []

if _IS_CLOUD:
    import streamlit as st
    st.warning("⚠️ Ejecutándose en Streamlit Cloud. Algunas funciones pueden estar limitadas.")
'''

def apply_patch():
    """Aplica el parche al archivo."""
    if not TARGET_FILE.exists():
        print(f"❌ Archivo no encontrado: {TARGET_FILE}")
        return False
    
    with open(TARGET_FILE, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Buscar fin del docstring (después de la línea que empieza con r""" o """)
    docstring_end = 0
    for i, line in enumerate(lines):
        if i > 0 and '"""' in line and not line.strip().startswith('#'):
            docstring_end = i + 1
            break
    
    if docstring_end == 0:
        print("❌ No se pudo encontrar el fin del docstring")
        return False
    
    # Inyectar después del docstring
    new_lines = lines[:docstring_end] + [PATCH_IMPORTS + '\n'] + lines[docstring_end:]
    
    # Crear backup
    backup_file = TARGET_FILE.with_suffix('.py.backup')
    if not backup_file.exists():
        with open(backup_file, 'w', encoding='utf-8') as f:
            f.writelines(lines)
        print(f"✓ Backup creado: {backup_file}")
    
    # Escribir archivo parchado
    with open(TARGET_FILE, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    
    print(f"✓ Archivo parchado: {TARGET_FILE}")
    print(f"  - Agregado sys.path para módulos locales")
    print(f"  - Importado config_deployment.py")
    print(f"  - Agregado aviso de Cloud (si aplica)")
    
    return True

def revert_patch():
    """Revierte el parche usando el backup."""
    backup_file = TARGET_FILE.with_suffix('.py.backup')
    if not backup_file.exists():
        print(f"❌ Backup no encontrado: {backup_file}")
        return False
    
    with open(backup_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    with open(TARGET_FILE, 'w', encoding='utf-8') as f:
        f.writelines(lines)
    
    print(f"✓ Parche revertido. Archivo restaurado desde: {backup_file}")
    return True

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Parchea interfaz_analisis_RPF.py para Streamlit Cloud")
    parser.add_argument("--revert", action="store_true", help="Revertir parche a versión original")
    
    args = parser.parse_args()
    
    if args.revert:
        if revert_patch():
            sys.exit(0)
        else:
            sys.exit(1)
    else:
        if apply_patch():
            sys.exit(0)
        else:
            sys.exit(1)
