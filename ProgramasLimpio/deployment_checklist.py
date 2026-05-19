#!/usr/bin/env python
"""
CHECKLIST DE DESPLIEGUE - interfaz_analisis_RPF.py
Script para validar que todo esté listo para Streamlit Cloud

Uso:
    python deployment_checklist.py
"""

import os
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent

def check_file_exists(path, description):
    """Verifica si un archivo existe."""
    exists = path.exists()
    status = "✓" if exists else "✗"
    print(f"  [{status}] {description}: {path.relative_to(SCRIPT_DIR) if exists or path.parent == SCRIPT_DIR else path}")
    return exists

def check_content(path, keywords, description):
    """Verifica si un archivo contiene palabras clave."""
    if not path.exists():
        print(f"  [✗] {description}: archivo no existe")
        return False
    
    try:
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        found_all = all(keyword in content for keyword in keywords)
        status = "✓" if found_all else "⚠"
        print(f"  [{status}] {description}")
        return found_all
    except Exception as e:
        print(f"  [✗] {description}: {e}")
        return False

def main():
    print("\n" + "="*70)
    print(" CHECKLIST DE DESPLIEGUE - interfaz_analisis_RPF.py → Streamlit Cloud")
    print("="*70 + "\n")
    
    results = {
        "archivos": [],
        "contenido": [],
        "git": [],
        "notas": []
    }
    
    # ─────────────────────────────────────────────────────────────────
    # SECCIÓN 1: ARCHIVOS REQUERIDOS
    # ─────────────────────────────────────────────────────────────────
    print("PARTE 1: ARCHIVOS REQUERIDOS")
    print("-" * 70)
    
    files_to_check = [
        (SCRIPT_DIR / "requirements.txt", "requirements.txt (Dependencias)"),
        (SCRIPT_DIR / ".gitignore", ".gitignore (Exclusiones Git)"),
        (SCRIPT_DIR / ".streamlit" / "config.toml", ".streamlit/config.toml (Config Streamlit)"),
        (SCRIPT_DIR / "README.md", "README.md (Documentación)"),
        (SCRIPT_DIR / "config_deployment.py", "config_deployment.py (Config auto-detectar)"),
        (SCRIPT_DIR / "interfaz_analisis_RPF.py", "interfaz_analisis_RPF.py (App principal)"),
        (SCRIPT_DIR / "graph_config.py", "graph_config.py (Config gráficas)"),
        (SCRIPT_DIR / "graph_builders.py", "graph_builders.py (Constructores)"),
    ]
    
    for path, desc in files_to_check:
        results["archivos"].append(check_file_exists(path, desc))
    
    # ─────────────────────────────────────────────────────────────────
    # SECCIÓN 2: CONTENIDO CRÍTICO
    # ─────────────────────────────────────────────────────────────────
    print("\nPARTE 2: CONTENIDO CRÍTICO")
    print("-" * 70)
    
    req_file = SCRIPT_DIR / "requirements.txt"
    results["contenido"].append(check_content(
        req_file,
        ["streamlit", "pandas", "plotly", "numpy"],
        "requirements.txt contiene dependencias clave"
    ))
    
    gitignore_file = SCRIPT_DIR / ".gitignore"
    results["contenido"].append(check_content(
        gitignore_file,
        ["__pycache__", "config_rutas.json", ".streamlit/secrets.toml"],
        ".gitignore excluye archivos sensibles"
    ))
    
    readme_file = SCRIPT_DIR / "README.md"
    results["contenido"].append(check_content(
        readme_file,
        ["Streamlit", "Instalación", "Despliegue"],
        "README.md documentación presente"
    ))
    
    # ─────────────────────────────────────────────────────────────────
    # SECCIÓN 3: ESTADO DE GIT
    # ─────────────────────────────────────────────────────────────────
    print("\nPARTE 3: ESTADO DE GIT")
    print("-" * 70)
    
    git_dir = SCRIPT_DIR / ".git"
    if git_dir.exists():
        print(f"  [✓] Repositorio Git inicializado")
        results["git"].append(True)
        
        # Verificar remoto
        try:
            import subprocess
            result = subprocess.run(
                ["git", "remote", "-v"],
                cwd=SCRIPT_DIR,
                capture_output=True,
                text=True,
                timeout=5
            )
            if "origin" in result.stdout:
                print(f"  [✓] Remoto 'origin' configurado")
                results["git"].append(True)
            else:
                print(f"  [⚠] Remoto 'origin' no encontrado - falta: git remote add origin <URL>")
                results["git"].append(False)
        except Exception as e:
            print(f"  [⚠] No se pudo verificar remoto: {e}")
            results["git"].append(False)
    else:
        print(f"  [✗] Repositorio Git NO inicializado")
        print(f"     Ejecutar: git init")
        print(f"     Luego: git remote add origin <URL>")
        results["git"].append(False)
    
    # ─────────────────────────────────────────────────────────────────
    # SECCIÓN 4: NOTAS IMPORTANTES
    # ─────────────────────────────────────────────────────────────────
    print("\nPARTE 4: PRÓXIMOS PASOS")
    print("-" * 70)
    
    pasos = []
    
    if not (SCRIPT_DIR / ".git").exists():
        pasos.append("1. Inicializar Git local:")
        pasos.append("   git init")
        pasos.append("   git add -A")
        pasos.append("   git commit -m 'Initial commit: RPF app ready for Streamlit Cloud'")
        pasos.append("")
    
    if not all(results["git"]):
        pasos.append("2. Crear repositorio en GitHub y conectarlo:")
        pasos.append("   git remote add origin https://github.com/TU_USUARIO/repo-nombre.git")
        pasos.append("   git branch -M main")
        pasos.append("   git push -u origin main")
        pasos.append("")
    
    pasos.extend([
        "3. Conectar en Streamlit Cloud:",
        "   • Ir a app.streamlit.io",
        "   • Click 'New app'",
        "   • Conectar con GitHub",
        "   • Seleccionar: repo, rama (main), archivo (interfaz_analisis_RPF.py)",
        "",
        "4. Configurar secretos (si es necesario):",
        "   • En Streamlit Cloud → Settings → Secrets",
        "   • Agregar variables de entorno sensibles",
        "",
        "5. Verificar despliegue:",
        "   • Acceder a la URL pública",
        "   • Validar funcionalidad",
    ])
    
    for paso in pasos:
        print(f"  {paso}")
    
    # ─────────────────────────────────────────────────────────────────
    # RESUMEN FINAL
    # ─────────────────────────────────────────────────────────────────
    print("\n" + "="*70)
    print(" RESUMEN")
    print("="*70 + "\n")
    
    archivos_ok = sum(results["archivos"])
    contenido_ok = sum(results["contenido"])
    git_ok = sum(results["git"])
    
    print(f"Archivos requeridos: {archivos_ok}/{len(results['archivos'])} ✓")
    print(f"Contenido crítico:   {contenido_ok}/{len(results['contenido'])} ✓")
    print(f"Git configurado:     {git_ok}/{len(results['git']) if results['git'] else 0} ✓")
    
    total_checks = len(results["archivos"]) + len(results["contenido"]) + len(results["git"])
    total_passed = archivos_ok + contenido_ok + git_ok
    
    print(f"\nTotal: {total_passed}/{total_checks} checks pasados\n")
    
    if total_passed == total_checks:
        print("✓ ¡LISTO PARA DESPLIEGUE! Continúa con los pasos de Git y Streamlit Cloud.")
        return 0
    else:
        print("⚠ Algunos checks no pasaron. Revisa el archivo de problemas:")
        print(f"  → {SCRIPT_DIR / 'DEPLOYMENT_ISSUES.md'}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
