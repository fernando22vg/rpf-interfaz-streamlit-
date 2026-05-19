# ================================================================
#   VERIFICADOR DE LIBRERÍAS PYTHON
#   Para proyectos con DIgSILENT PowerFactory
#
# ================================================================

librerias = [
    # Nombre de importación    # Nombre pip          # Categoría
    ("numpy",                  "numpy",               "Cálculo numérico"),
    ("matplotlib",             "matplotlib",          "Visualización"),
    ("scipy",                  "scipy",               "Cálculo científico"),
    ("control",                "control",             "Sistemas de control"),
    ("pandas",                 "pandas",              "Manejo de datos"),
    ("win32com",               "pywin32",             "Conexión Excel/PowerFactory"),
    ("seaborn",                "seaborn",             "Visualización estadística"),
    ("plotly",                 "plotly",              "Gráficas interactivas"),
    ("pandapower",             "pandapower",          "Redes eléctricas"),
    ("pvlib",                  "pvlib",               "Energía solar"),
    ("openpyxl",               "openpyxl",            "Archivos Excel"),
    ("seaborn",                "seaborn",             "visualiazación estadística"),
    ("xlrd",                   "xlrd",                "Archivos Excel antiguos"),
    ("sympy",                  "sympy",               "Álgebra simbólica"),
    ("pymoo",                  "pymoo",               "Optimización genética"),
    ("deap",                   "deap",                "Algoritmos evolutivos"),
    ("tensorflow",             "tensorflow",          "Redes neuronales ANN"),
    ("streamlit",              "streamlit",           "Aplicaciones"),
    ("sklearn",                "scikit-learn",        "Machine learning"),
    ("slycot",                 "slycot",              "Control avanzado"),
    ("spyder",                 "spyder",              "IDE Spyder"),
]

# ── Colores para la terminal ──────────────────────────────────
VERDE  = "\033[92m"
ROJO   = "\033[91m"
RESET  = "\033[0m"
AZUL   = "\033[94m"
AMARILLO = "\033[93m"

# ── Encabezado ────────────────────────────────────────────────
print("=" * 65)
print(f"{AZUL}   VERIFICADOR DE LIBRERÍAS PYTHON - DIGSILENT + TESIS{RESET}")
print("=" * 65)
print(f"{'Librería':<20} {'pip install':<20} {'Estado':<12} {'Categoría'}")
print("-" * 65)

instaladas   = []
faltantes    = []

for importar, pip_nombre, categoria in librerias:
    try:
        lib = __import__(importar)
        version = getattr(lib, "__version__", "OK")
        print(f"{importar:<20} {pip_nombre:<20} {VERDE}✅ {version:<10}{RESET} {categoria}")
        instaladas.append(importar)
    except ImportError:
        print(f"{importar:<20} {pip_nombre:<20} {ROJO}❌ FALTA{RESET}      {categoria}")
        faltantes.append(pip_nombre)

# ── Verificar PowerFactory por separado ──────────────────────
print("-" * 65)
try:
    import powerfactory
    print(f"{'powerfactory':<20} {'(DIgSILENT)':<20} {VERDE}✅ OK{RESET}         Conexión PowerFactory")
    instaladas.append("powerfactory")
except ImportError:
    print(f"{'powerfactory':<20} {'(DIgSILENT)':<20} {AMARILLO}⚠️  PATH{RESET}       Configura el path de DIgSILENT")

# ── Resumen final ─────────────────────────────────────────────
print("=" * 65)
print(f"{AZUL}RESUMEN:{RESET}")
print(f"  {VERDE}✅ Instaladas: {len(instaladas)}{RESET}")
print(f"  {ROJO}❌ Faltantes:  {len(faltantes)}{RESET}")

if faltantes:
    print("\n" + "=" * 65)
    print(f"{AMARILLO}Ejecuta este comando para instalar las faltantes:{RESET}")
    print(f"\npip install {' '.join(faltantes)}\n")
else:
    print(f"\n{VERDE}¡Todas las librerías están instaladas correctamente!{RESET}\n")

print("=" * 65)