📤 GUÍA: CÓMO CARGAR ARCHIVOS A SHAREPOINT PARA DESPLIEGUE
═══════════════════════════════════════════════════════════════════════════════

SHAREPOINT: https://cobee1-my.sharepoint.com/personal/angel_mariscal_cobee_com/Documents

OBJETIVO: Copiar TODOS los archivos necesarios desde tu PC a SharePoint

═══════════════════════════════════════════════════════════════════════════════


🎯 ARCHIVOS A CARGAR (CRÍTICOS - 13 ARCHIVOS)
═══════════════════════════════════════════════════════════════════════════════

ESTRUCTURA EN SHAREPOINT:
────────────────────────────────────────────────────────────────────────────

📁 32_REGULACIÓN_DE_FRECUENCIA/
├── 📁 interfaz_analisis_RPF/                    ← NUEVA CARPETA
│   ├── 📄 interfaz_analisis_RPF.py              ✅ CRÍTICO
│   ├── 📄 graph_config.py                       ✅ CRÍTICO
│   ├── 📄 graph_builders.py                     ✅ CRÍTICO
│   │
│   ├── 📋 CONFIGURACIÓN
│   ├── 📄 requirements.txt                      ✅ CRÍTICO
│   ├── 📄 .gitignore                            ✅ CRÍTICO
│   ├── 📄 config_deployment.py                  ✅ CRÍTICO
│   │
│   ├── 📋 DOCUMENTACIÓN
│   ├── 📄 README.md                             ✅ LEER PRIMERO
│   ├── 📄 00_START_HERE.md                      ✅ LEER PRIMERO
│   ├── 📄 QUICKSTART.txt                        ✅ GUÍA PASOS
│   ├── 📄 DEPLOYMENT_ISSUES.md                  📖 Problemas
│   ├── 📄 DEPLOYMENT_SUMMARY.md                 📖 Resumen
│   ├── 📄 DEPLOYMENT_INVENTORY.md               📖 Inventario
│   ├── 📄 REQUISITOS_USUARIO_FINAL.md           👥 Para usuarios
│   │
│   ├── 📋 HERRAMIENTAS
│   ├── 📄 setup_deployment.py                   🛠️ Setup
│   ├── 📄 deployment_checklist.py               ✓ Validación
│   ├── 📄 patch_for_streamlit_cloud.py          ⚙️ Patch
│   │
│   ├── 📁 app/                                  📦 Módulos
│   ├── 📁 core/
│   ├── 📁 runners/
│   │
│   └── 📁 .streamlit/                           ⚙️ Config
│       └── 📄 config.toml                       (Crear con setup)


📋 PASO A PASO: CÓMO CARGAR
═══════════════════════════════════════════════════════════════════════════════

PASO 1: Preparar archivos locales
─────────────────────────────────────────────────────────────────────────────

Ubicación local:
  c:\Programas Python\ProgramasLimpio\

Archivos a copiar:
  ✅ interfaz_analisis_RPF.py
  ✅ graph_config.py
  ✅ graph_builders.py
  ✅ requirements.txt
  ✅ .gitignore
  ✅ config_deployment.py
  ✅ setup_deployment.py
  ✅ deployment_checklist.py
  ✅ patch_for_streamlit_cloud.py
  ✅ README.md
  ✅ 00_START_HERE.md
  ✅ QUICKSTART.txt
  ✅ DEPLOYMENT_ISSUES.md
  ✅ DEPLOYMENT_SUMMARY.md
  ✅ DEPLOYMENT_INVENTORY.md
  ✅ REQUISITOS_USUARIO_FINAL.md
  ✅ carpetas: app/, core/, runners/

TOTAL: 16 archivos + 3 carpetas


PASO 2: Crear carpeta en SharePoint
─────────────────────────────────────────────────────────────────────────────

1. Ir a SharePoint:
   https://cobee1-my.sharepoint.com/personal/angel_mariscal_cobee_com/Documents

2. Navegar a:
   📁 Documents → 📁 32_REGULACIÓN_DE_FRECUENCIA

3. Click derecho → "New" → "Folder"

4. Nombre: "interfaz_analisis_RPF"

5. Crear


PASO 3A: Cargar vía navegador (RECOMENDADO)
─────────────────────────────────────────────────────────────────────────────

MÉTODO 1: Drag & Drop (Más fácil)

1. Abrir SharePoint en navegador
2. Navegar a: 32_REGULACIÓN_DE_FRECUENCIA/interfaz_analisis_RPF/
3. Abrir Explorador de archivos (File Explorer)
4. Navegar a: c:\Programas Python\ProgramasLimpio\
5. SELECCIONAR todos los archivos:
   - Usar Ctrl+A (selecciona todo)
   - O Ctrl+Click en cada archivo
6. ARRASTRAR archivos a ventana de SharePoint
7. SOLTAR en carpeta SharePoint
8. ✅ Esperar que cargue (~2-5 min)


MÉTODO 2: Upload (Alternativo)

1. Abrir carpeta en SharePoint
2. Click en "Upload"
3. Seleccionar archivos
4. Click "Open"
5. ✅ Cargando...


MÉTODO 3: Sincronización (Avanzado)

1. Click en "Sync" en SharePoint
2. Conectar con OneSync
3. Carpeta en PC se sincroniza automáticamente
4. Copiar archivos a la carpeta local
5. ✅ Se sincroniza solo


PASO 3B: Cargar vía línea de comandos (PowerShell)
─────────────────────────────────────────────────────────────────────────────

# Instalar PnP PowerShell si no lo tienes
Install-Module -Name PnP.PowerShell -Force

# Conectar a SharePoint
Connect-PnPOnline -Url "https://cobee1-my.sharepoint.com/personal/angel_mariscal_cobee_com" -Interactive

# Cargar archivos
$files = @(
  "c:\Programas Python\ProgramasLimpio\interfaz_analisis_RPF.py",
  "c:\Programas Python\ProgramasLimpio\requirements.txt",
  # ... otros archivos
)

foreach ($file in $files) {
  Add-PnPFile -Path $file -Folder "Documents/32_REGULACIÓN_DE_FRECUENCIA/interfaz_analisis_RPF"
}


PASO 4: Verificar carga
─────────────────────────────────────────────────────────────────────────────

Después de cargar:

☑ ¿Todos los .py están?
☑ ¿Están requirements.txt y .gitignore?
☑ ¿Están los .md de documentación?
☑ ¿Están las carpetas app/, core/, runners/?
☑ ¿El contador de archivos es correcto?

Si todo ✅ → Pasar al Paso 5


PASO 5: Crear .streamlit/config.toml
─────────────────────────────────────────────────────────────────────────────

1. En SharePoint, crear subcarpeta:
   📁 .streamlit

2. Crear archivo: config.toml

   Contenido:
   ─────────────────────────────────────
   [theme]
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
   ─────────────────────────────────────

3. Guardar en SharePoint


PASO 6: Descargar para usar localmente
─────────────────────────────────────────────────────────────────────────────

Si quieres usar la app:

1. En SharePoint: Select all → Download (genera ZIP)
2. Descomprimir ZIP
3. Copiar a c:\Programas Python\ProgramasLimpio\
4. Ejecutar: python setup_deployment.py
5. Ejecutar: python deployment_checklist.py
6. Ejecutar: streamlit run interfaz_analisis_RPF.py


═══════════════════════════════════════════════════════════════════════════════


📊 ESTRUCTURA FINAL EN SHAREPOINT
═══════════════════════════════════════════════════════════════════════════════

Documents/
└── 32_REGULACIÓN_DE_FRECUENCIA/
    └── interfaz_analisis_RPF/
        ├── 📄 interfaz_analisis_RPF.py
        ├── 📄 graph_config.py
        ├── 📄 graph_builders.py
        ├── 📄 requirements.txt
        ├── 📄 .gitignore
        ├── 📄 config_deployment.py
        ├── 📄 setup_deployment.py
        ├── 📄 deployment_checklist.py
        ├── 📄 patch_for_streamlit_cloud.py
        ├── 📄 README.md
        ├── 📄 00_START_HERE.md
        ├── 📄 QUICKSTART.txt
        ├── 📄 DEPLOYMENT_ISSUES.md
        ├── 📄 DEPLOYMENT_SUMMARY.md
        ├── 📄 DEPLOYMENT_INVENTORY.md
        ├── 📄 REQUISITOS_USUARIO_FINAL.md
        ├── 📁 app/
        │   ├── __init__.py
        │   ├── streamlit_app.py
        │   ├── pages/
        │   └── components/
        ├── 📁 core/
        │   ├── __init__.py
        │   ├── config.py
        │   └── data_extraction.py
        ├── 📁 runners/
        │   └── [11 runners .py]
        └── 📁 .streamlit/
            └── config.toml


✅ CHECKLIST DE CARGA
═══════════════════════════════════════════════════════════════════════════════

ANTES DE CARGAR:
[ ] Preparé archivos locales en c:\Programas Python\ProgramasLimpio\
[ ] Verifiqué que setup_deployment.py existe
[ ] Verifiqué que todos los .py están en la carpeta local
[ ] Tengo acceso a SharePoint

DURANTE LA CARGA:
[ ] Creé carpeta "interfaz_analisis_RPF" en SharePoint
[ ] Cargué todos los archivos principales
[ ] Cargué todas las documentaciones
[ ] Cargué las 3 carpetas (app, core, runners)
[ ] Creé carpeta .streamlit/
[ ] Creé archivo config.toml

DESPUÉS DE LA CARGA:
[ ] Verifiqué que todos los archivos están en SharePoint
[ ] Puede ver 16+ archivos en la carpeta
[ ] Las carpetas están visibles
[ ] El archivo config.toml está en .streamlit/

SI TODO ✅:
→ Descargue ZIP de SharePoint
→ Descomprima localmente
→ Ejecute setup_deployment.py
→ Ejecute deployment_checklist.py
→ ¡Listo para desplegar!


📞 PROBLEMAS COMUNES
═══════════════════════════════════════════════════════════════════════════════

"No puedo crear carpeta en SharePoint"
→ Verificar permisos de acceso
→ Contactar admin de SharePoint

"El archivo es muy grande"
→ Split en partes
→ O comprimir antes de cargar

"Se carga muy lento"
→ Usar Sync en lugar de Upload
→ O cargar pocos archivos a la vez

"No veo los archivos después de cargar"
→ Refrescar (F5)
→ Esperar 1-2 minutos (sincronización)

"Error 403 - Access Denied"
→ No tienes permisos
→ Contactar propietario de carpeta


═══════════════════════════════════════════════════════════════════════════════

RESUMEN:
1. Crear carpeta "interfaz_analisis_RPF" en SharePoint
2. Cargar 16 archivos + 3 carpetas
3. Crear config.toml en .streamlit/
4. ✅ Listo para usar

Tiempo estimado: 10-15 minutos

═══════════════════════════════════════════════════════════════════════════════
