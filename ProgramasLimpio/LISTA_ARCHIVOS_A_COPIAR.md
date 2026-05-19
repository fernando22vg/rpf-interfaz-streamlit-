📋 LISTA DE VERIFICACIÓN: ARCHIVOS A COPIAR A SHAREPOINT
═══════════════════════════════════════════════════════════════════════════════

ORIGEN: c:\Programas Python\ProgramasLimpio\
DESTINO: SharePoint → 32_REGULACIÓN_DE_FRECUENCIA/interfaz_analisis_RPF/

═══════════════════════════════════════════════════════════════════════════════


✅ ARCHIVOS PRINCIPALES - COPIAR TODO
═════════════════════════════════════════════════════════════════════════════

ARCHIVO PRINCIPAL (CRÍTICO):
  [ ] interfaz_analisis_RPF.py          ← LA APP PRINCIPAL

ARCHIVOS DE GRÁFICAS (CRÍTICO):
  [ ] graph_config.py                   ← Configuración de gráficas
  [ ] graph_builders.py                 ← Constructores de gráficas

ARCHIVOS DE CONFIGURACIÓN (CRÍTICO):
  [ ] requirements.txt                  ← Dependencias Python
  [ ] .gitignore                        ← Excluir de Git
  [ ] config_deployment.py              ← Config inteligente

ARCHIVOS DE SETUP (IMPORTANTE):
  [ ] setup_deployment.py               ← Script de inicialización
  [ ] deployment_checklist.py           ← Script de validación
  [ ] patch_for_streamlit_cloud.py      ← Script de parche (opcional)

ARCHIVOS DE DOCUMENTACIÓN (IMPORTANTE):
  [ ] README.md                         ← Documentación principal
  [ ] 00_START_HERE.md                  ← Guía de inicio
  [ ] QUICKSTART.txt                    ← Pasos rápidos
  [ ] DEPLOYMENT_ISSUES.md              ← Problemas técnicos
  [ ] DEPLOYMENT_SUMMARY.md             ← Resumen ejecutivo
  [ ] DEPLOYMENT_INVENTORY.md           ← Inventario
  [ ] REQUISITOS_USUARIO_FINAL.md       ← Requisitos para usuarios
  [ ] SHAREPOINT_UPLOAD_GUIDE.md        ← Esta guía

TOTAL ARCHIVOS: 18 archivos


✅ CARPETAS - COPIAR TODO (con contenido)
═════════════════════════════════════════════════════════════════════════════

CARPETA: app/
  [ ] app/__init__.py
  [ ] app/streamlit_app.py
  [ ] app/pages/                        (subdirectorio)
  [ ] app/components/                   (subdirectorio)

CARPETA: core/
  [ ] core/__init__.py
  [ ] core/config.py
  [ ] core/data_extraction.py

CARPETA: runners/
  [ ] runners/CargaCondIniciales_PF_run.py
  [ ] runners/CondInicialesPF_run.py
  [ ] runners/ExtFLujos2daO_run.py
  [ ] runners/ExtractorResultadosCNDC_run.py
  [ ] runners/InventarioShunts_PF_run.py
  [ ] runners/MapeoRetirosSTI_run.py
  [ ] runners/OrdenadorDatosEvento_run.py
  [ ] runners/loc_namesGEN_run.py
  [ ] runners/loc_namesLineas_run.py
  [ ] runners/loc_names_xfo_run.py
  [ ] runners/DatsoGENBUSLNE_run.py
  (Total: 11 archivos en runners/)

TOTAL CARPETAS: 3 carpetas + subdirectorios


❌ ARCHIVOS QUE NO COPIAR (excluir)
═════════════════════════════════════════════════════════════════════════════

NO COPIAR:
  ❌ config_rutas.json              (Rutas locales Windows)
  ❌ .streamlit/secrets.toml         (Secretos - crear manualmente)
  ❌ __pycache__/                    (Bytecode compilado)
  ❌ .pytest_cache/                  (Cache de tests)
  ❌ *.pyc                           (Compilados Python)
  ❌ test_*.py                       (Tests locales)
  ❌ verify_config_keys.py           (Script de verificación local)
  ❌ ComparativaREAL_SIMU_RMS.py     (Archivos antiguos)
  ❌ *.xlsx, *.csv                   (Datos locales - muy grandes)

RAZÓN: No necesarios en cloud, pueden confundir, muy grandes


✅ DESPUÉS DE COPIAR: CREAR MANUALMENTE
═════════════════════════════════════════════════════════════════════════════

CARPETA: .streamlit/
  [ ] Crear carpeta: .streamlit

ARCHIVO: .streamlit/config.toml
  [ ] Crear archivo manualmente con contenido:

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


📊 RESUMEN DE COPIAS
═════════════════════════════════════════════════════════════════════════════

TIPO                    CANTIDAD  TOTAL
─────────────────────────────────────────
Archivos .py principales    3      3
Archivos .py config         3      3
Archivos .py setup          3      3
Archivos .md docs           8      8
Archivos en app/            4      4
Archivos en core/           3      3
Archivos en runners/        11     11

TOTAL A COPIAR:                   38 archivos + 3+ carpetas


🎯 ORDEN RECOMENDADO DE COPIA
═════════════════════════════════════════════════════════════════════════════

PRIMERO (Crítico - la app funciona con estos):
1. interfaz_analisis_RPF.py
2. graph_config.py
3. graph_builders.py
4. requirements.txt
5. config_deployment.py

SEGUNDO (Setup):
6. setup_deployment.py
7. deployment_checklist.py

TERCERO (Documentación - guías):
8. README.md
9. 00_START_HERE.md
10. QUICKSTART.txt

CUARTO (Documentación - referencia):
11. DEPLOYMENT_ISSUES.md
12. DEPLOYMENT_SUMMARY.md
13. DEPLOYMENT_INVENTORY.md
14. REQUISITOS_USUARIO_FINAL.md
15. SHAREPOINT_UPLOAD_GUIDE.md

QUINTO (Módulos):
16. Carpeta: app/
17. Carpeta: core/
18. Carpeta: runners/

SEXTO (Config):
19. .gitignore
20. patch_for_streamlit_cloud.py
21. Crear: .streamlit/config.toml


✅ VALIDACIÓN POST-COPIA
═════════════════════════════════════════════════════════════════════════════

Después de copiar TODO, verificar:

[ ] ¿Están los 3 archivos .py principales? (interfaz, graph_config, graph_builders)
[ ] ¿Está requirements.txt?
[ ] ¿Está .gitignore?
[ ] ¿Está config_deployment.py?
[ ] ¿Están los 3 scripts de setup?
[ ] ¿Están todos los .md?
[ ] ¿Está la carpeta app/? ¿Con __init__.py?
[ ] ¿Está la carpeta core/? ¿Con __init__.py?
[ ] ¿Está la carpeta runners/? ¿Con 11 archivos?
[ ] ¿Está la carpeta .streamlit/?
[ ] ¿Está config.toml en .streamlit/?

SI TODOS ✅ → ¡COPIA EXITOSA!


📥 CÓMO VERIFICAR EN SHAREPOINT
═════════════════════════════════════════════════════════════════════════════

1. Abrir SharePoint
2. Ir a: Documents → 32_REGULACIÓN_DE_FRECUENCIA → interfaz_analisis_RPF/
3. Debe ver:
   - 18-20 archivos .py y .md
   - 3 carpetas: app/, core/, runners/
   - 1 carpeta: .streamlit/

4. Contar: Si ve 20+ items → ¡Todo está!


🔄 CÓMO DESCARGAR LUEGO DESDE SHAREPOINT
═════════════════════════════════════════════════════════════════════════════

Cuando necesites usar la app:

1. Abrir SharePoint
2. Ir a: interfaz_analisis_RPF/
3. Seleccionar todo: Ctrl+A
4. Click en "Download"
5. Se genera ZIP automáticamente
6. Descargar ZIP
7. Descomprimir en: c:\Programas Python\ProgramasLimpio\
8. Ejecutar: python setup_deployment.py
9. Ejecutar: streamlit run interfaz_analisis_RPF.py


═══════════════════════════════════════════════════════════════════════════════

RESUMEN FINAL:

✅ Copiar: 18 archivos .py/.md + 3 carpetas
❌ No copiar: config_rutas.json, __pycache__, datos locales
🆕 Crear: .streamlit/config.toml manualmente

Después → Descargar como ZIP → Descomprimir → ¡Usar!

═══════════════════════════════════════════════════════════════════════════════
