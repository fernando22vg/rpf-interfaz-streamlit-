#!/usr/bin/env python
"""Script de inicialización para despliegue en Streamlit Cloud"""
import os
import sys

# Crear directorio .streamlit si no existe
streamlit_dir = os.path.join(os.path.dirname(__file__), '.streamlit')
os.makedirs(streamlit_dir, exist_ok=True)

# Crear config.toml
config_path = os.path.join(streamlit_dir, 'config.toml')
if not os.path.exists(config_path):
    config_content = """[theme]
primaryColor = "#2E4057"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#262730"
font = "sans serif"

[client]
showErrorDetails = true
toolbarMode = "viewer"

[logger]
level = "info"

[server]
port = 8501
headless = true
runOnSave = true
maxUploadSize = 200
enableCORS = false
enableXsrfProtection = true

[browser]
gatherUsageStats = false
serverAddress = "localhost"
"""
    with open(config_path, 'w') as f:
        f.write(config_content)
    print(f"✓ Creado: {config_path}")
else:
    print(f"✓ Ya existe: {config_path}")

# Crear .gitignore si no existe
gitignore_path = os.path.join(os.path.dirname(__file__), '.gitignore')
if not os.path.exists(gitignore_path):
    print(f"✓ Ya existe: {gitignore_path}")
else:
    print(f"✓ Ya existe: {gitignore_path}")

print("\n✓ Inicialización completada para despliegue.")
