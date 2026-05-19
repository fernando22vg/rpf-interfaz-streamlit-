# Interfaz de Análisis RPF - Streamlit App

Aplicación Streamlit para análisis integrado de RPF (Respuesta de Frecuencia), con generación de condiciones iniciales y carga en PowerFactory.

## Requisitos

- Python 3.8+
- pip (gestor de paquetes)

## Instalación Local

```bash
# Clonar o descargar el repositorio
cd ProgramasLimpio

# Crear entorno virtual (opcional pero recomendado)
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar la aplicación
streamlit run interfaz_analisis_RPF.py
```

## Despliegue en Streamlit Cloud

### Requisitos previos:
- Cuenta en GitHub
- Repositorio Git con este código
- Cuenta en app.streamlit.io

### Pasos:
1. Hacer push del código a GitHub (rama `main`)
2. Ir a app.streamlit.io y conectar con GitHub
3. Seleccionar repositorio, rama y archivo (`interfaz_analisis_RPF.py`)
4. Streamlit Cloud desplegará automáticamente

### Configurar Secretos (si necesario):
En Streamlit Cloud UI → Settings → Secrets, agregar:
```toml
RAIZ = "ruta/a/datos"
RAIZ_DATOS = "ruta/a/datos2"
```

## Estructura del Proyecto

```
ProgramasLimpio/
├── interfaz_analisis_RPF.py      # 🎯 ARCHIVO PRINCIPAL
├── graph_config.py                # Configuración de gráficas
├── graph_builders.py              # Constructores de gráficas
├── requirements.txt               # Dependencias Python
├── .gitignore                     # Archivos a ignorar en Git
├── .streamlit/
│   └── config.toml               # Configuración de Streamlit
├── app/
│   ├── pages/                    # Páginas adicionales (si aplica)
│   ├── components/               # Componentes reutilizables
│   └── streamlit_app.py          # App auxiliar
├── core/
│   ├── config.py                 # Configuración
│   └── data_extraction.py        # Extracción de datos
└── runners/
    ├── ExtFLujos2daO_run.py      # Runner de extracción
    ├── CondInicialesPF_run.py    # Runner de condiciones iniciales
    └── ...otros runners...
```

## Documentos Críticos

| Archivo | Ubicación | Propósito |
|---------|-----------|----------|
| `requirements.txt` | `ProgramasLimpio/requirements.txt` | Dependencias Python |
| `config.toml` | `ProgramasLimpio/.streamlit/config.toml` | Configuración de Streamlit |
| `.gitignore` | `ProgramasLimpio/.gitignore` | Archivos a excluir de Git |
| `config_rutas.json` | `ProgramasLimpio/config_rutas.json` | Configuración de rutas locales (no se carga en nube) |

## Edición y Desarrollo

### Flujo Local:
```bash
# 1. Editar archivos localmente
nano interfaz_analisis_RPF.py

# 2. Testear localmente
streamlit run interfaz_analisis_RPF.py

# 3. Hacer commit
git add -A
git commit -m "Cambios: [descripción]"

# 4. Push a GitHub
git push origin main

# → Streamlit Cloud redeploy automático (~30 seg)
```

### Sincronización con Streamlit Cloud:
- Los cambios se detectan automáticamente
- Redeploy automático cada vez que se hace push a `main`
- No es necesario hacer deploy manual

## Problemas Comunes

### 1. **"ModuleNotFoundError: No module named 'X'"**
- **Causa:** Paquete faltante en `requirements.txt`
- **Solución:** Agregar paquete a `requirements.txt` y hacer push

### 2. **"FileNotFoundError: [Errno 2] No such file or directory"**
- **Causa:** Rutas hardcodeadas apuntando a directorios locales
- **Solución:** Usar rutas relativas o configurar en `secrets.toml`

### 3. **"Permission denied" o "Resource is read-only"**
- **Causa:** Intentar escribir en directorios de solo lectura en Streamlit Cloud
- **Solución:** Usar `st.session_state` para datos temporales, guardar en `/tmp` si es necesario

### 4. **App se ralentiza o timeout**
- **Causa:** Cálculos pesados sin caché
- **Solución:** Usar `@st.cache_data` para resultados costosos

### 5. **"Git fatal: not a git repository"**
- **Causa:** Directorio no inicializado con Git
- **Solución:** Ejecutar `git init` y configurar remoto

## Variables de Entorno

Streamlit Cloud soporta secretos. Para agregar:

1. En local: crear `~/.streamlit/secrets.toml`:
   ```toml
   RAIZ = "C:\ruta\a\datos"
   PASSWORD = "mi_contraseña"
   ```

2. En Streamlit Cloud: UI → Settings → Secrets

Acceder en código:
```python
import streamlit as st
raiz = st.secrets.get("RAIZ", "default_value")
```

## Licencia

Proyecto interno CNDC.

## Contacto

Para problemas o preguntas, contactar al equipo de desarrollo.
