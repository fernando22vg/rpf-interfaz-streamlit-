# Extraccion completa de datos de red PowerFactory
# Barras, Lineas, Generadores, Cargas, Escenarios de Operacion y Study Cases

import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2\Python\3.11")
import os
os.environ["PATH"] = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2" + os.pathsep + os.environ["PATH"]
import pandas as pd
import powerfactory as pf

# ---------------------------------------------------------------------------
# Conexion a PowerFactory
# ---------------------------------------------------------------------------
os.system("taskkill /f /im PowerFactory.exe >nul 2>&1")

app = pf.GetApplication()
if app is None:
    raise RuntimeError(
        "No se pudo obtener la aplicacion de PowerFactory. "
        "Cierre PowerFactory si esta abierto y vuelva a ejecutar el script."
    )

app.Show()
user = app.GetCurrentUser()
prj  = app.ActivateProject("Pruebas")
print(f"Proyecto activo: {prj}")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def ga(obj, attr):
    """Lee un atributo de un objeto PF. Retorna None si no existe."""
    try:
        return obj.GetAttribute(attr) if obj is not None else None
    except Exception:
        return None

def ga_round(obj, attr, dec=2):
    val = ga(obj, attr)
    try:
        return round(val, dec) if val is not None else None
    except Exception:
        return val

def terminal_de(elm):
    """Dado un elemento, retorna el ElmTerm conectado en bus1."""
    cubicle = ga(elm, 'bus1')
    if cubicle is None:
        return None
    term = ga(cubicle, 'cterm')
    return term

# ---------------------------------------------------------------------------
# Ejecutar flujo de carga BASE
# ---------------------------------------------------------------------------
ldf = app.GetFromStudyCase('ComLdf')
ldf.Execute()
print("Flujo de carga base ejecutado.")

# ---------------------------------------------------------------------------
# 1. BARRAS (ElmTerm)
# ---------------------------------------------------------------------------
buses = app.GetCalcRelevantObjects('*.ElmTerm')
bus_data = []
for bus in buses:
    uknom = ga(bus, 'uknom')
    u_pu  = ga(bus, 'm:u')
    u_kv  = round(u_pu * uknom, 4) if u_pu is not None and uknom is not None else None
    bus_data.append({
        "Nombre":             ga(bus, 'loc_name'),
        "Tension nom. (kV)":  uknom,
        "Tension (pu)":       u_pu,
        "Tension (kV)":       u_kv,
        "Angulo (deg)":       ga_round(bus, 'm:phiu', 4),
        "En servicio":        bool(ga(bus, 'outserv') == 0),
    })
df_buses = pd.DataFrame(bus_data)
print(f"Barras: {len(df_buses)}")

# ---------------------------------------------------------------------------
# 2. LINEAS (ElmLne)
# ---------------------------------------------------------------------------
lines = app.GetCalcRelevantObjects('*.ElmLne')
line_data = []
for ln in lines:
    typ     = ga(ln, 'typ_id')
    uline   = ga(typ, 'uline')  if typ is not None else None
    inom    = ga(typ, 'InomAC') if typ is not None else None
    loading = ga(ln, 'c:loading')
    line_data.append({
        "Nombre":             ga(ln, 'loc_name'),
        "Distancia (km)":     ga_round(ln, 'dline', 2),
        "Tension nom. (kV)":  uline,
        "Carga (%)":          round(loading, 2) if loading is not None else None,
        "P from (MW)":        ga_round(ln, 'm:P:bus1', 4),
        "Q from (Mvar)":      ga_round(ln, 'm:Q:bus1', 4),
        "P to (MW)":          ga_round(ln, 'm:P:bus2', 4),
        "Q to (Mvar)":        ga_round(ln, 'm:Q:bus2', 4),
        "En servicio":        bool(ga(ln, 'outserv') == 0),
    })
df_lines = pd.DataFrame(line_data)
print(f"Lineas: {len(df_lines)}")

# ---------------------------------------------------------------------------
# 3. GENERADORES
# Tension (pu) se lee desde la barra conectada, no del generador directamente
# ---------------------------------------------------------------------------
gen_data = []
for cls in ('*.ElmSym', '*.ElmGenstat', '*.ElmPvsys', '*.ElmWind'):
    for gen in app.GetCalcRelevantObjects(cls):
        term  = terminal_de(gen)
        u_pu  = ga(term, 'm:u') if term is not None else None
        gen_data.append({
            "Nombre":           ga(gen, 'loc_name'),
            "Tipo objeto":      gen.GetClassName(),
            "Barra conectada":  ga(term, 'loc_name'),
            "P nom. (MW)":      ga_round(gen, 'pgini', 4),
            "Q nom. (Mvar)":    ga_round(gen, 'qgini', 4),
            "P result. (MW)":   ga_round(gen, 'm:P:bus1', 4),
            "Q result. (Mvar)": ga_round(gen, 'm:Q:bus1', 4),
            "Tension (pu)":     round(u_pu, 4) if u_pu is not None else None,
            "Cos phi":          ga_round(gen, 'cosini', 4),
            "En servicio":      bool(ga(gen, 'outserv') == 0),
        })
df_gens = pd.DataFrame(gen_data)
print(f"Generadores: {len(df_gens)}")

# ---------------------------------------------------------------------------
# 4. CARGAS (ElmLod)
# ---------------------------------------------------------------------------
loads = app.GetCalcRelevantObjects('*.ElmLod')
load_data = []
for ld in loads:
    load_data.append({
        "Nombre":           ga(ld, 'loc_name'),
        "P nom. (MW)":      ga_round(ld, 'plini', 4),
        "Q nom. (Mvar)":    ga_round(ld, 'qlini', 4),
        "P result. (MW)":   ga_round(ld, 'm:P:bus1', 4),
        "Q result. (Mvar)": ga_round(ld, 'm:Q:bus1', 4),
        "Cos phi":          ga_round(ld, 'coslini', 4),
        "En servicio":      bool(ga(ld, 'outserv') == 0),
    })
df_loads = pd.DataFrame(load_data)
print(f"Cargas: {len(df_loads)}")

# ---------------------------------------------------------------------------
# 5. ESCENARIOS DE OPERACION — cambios respecto al caso base
#
# Estrategia: capturar estado base, luego activar cada escenario,
# comparar parametros de cada elemento y registrar diferencias.
# ---------------------------------------------------------------------------

def capturar_estado_red():
    """Devuelve un dict con los parametros operativos actuales de cada elemento."""
    estado = {}
    for gen in app.GetCalcRelevantObjects('*.ElmSym') + \
               app.GetCalcRelevantObjects('*.ElmGenstat') + \
               app.GetCalcRelevantObjects('*.ElmPvsys') + \
               app.GetCalcRelevantObjects('*.ElmWind'):
        nombre = ga(gen, 'loc_name')
        estado[f"GEN|{nombre}"] = {
            "En servicio": bool(ga(gen, 'outserv') == 0),
            "P (MW)":      ga_round(gen, 'pgini', 4),
            "Q (Mvar)":    ga_round(gen, 'qgini', 4),
        }
    for ld in app.GetCalcRelevantObjects('*.ElmLod'):
        nombre = ga(ld, 'loc_name')
        estado[f"CARGA|{nombre}"] = {
            "En servicio": bool(ga(ld, 'outserv') == 0),
            "P (MW)":      ga_round(ld, 'plini', 4),
            "Q (Mvar)":    ga_round(ld, 'qlini', 4),
        }
    for ln in app.GetCalcRelevantObjects('*.ElmLne'):
        nombre = ga(ln, 'loc_name')
        estado[f"LINEA|{nombre}"] = {
            "En servicio": bool(ga(ln, 'outserv') == 0),
            "P (MW)":      None,
            "Q (Mvar)":    None,
        }
    return estado

def comparar_estados(base, otro):
    """Retorna lista de cambios entre estado base y otro estado."""
    cambios = []
    for elem, vals_otro in otro.items():
        vals_base = base.get(elem, {})
        tipo, nombre = elem.split("|", 1)
        for param, v_otro in vals_otro.items():
            v_base = vals_base.get(param)
            if v_base != v_otro and v_otro is not None:
                cambios.append({
                    "Elemento":       nombre,
                    "Tipo":           tipo,
                    "Parametro":      param,
                    "Valor base":     v_base,
                    "Valor escenario":v_otro,
                })
    return cambios

# Capturar estado base (sin escenario activo)
estado_base = capturar_estado_red()

scen_list_data = []   # info basica de cada escenario
scen_changes   = []   # todos los cambios de todos los escenarios juntos

scenarios_folder = app.GetProjectFolder('scen')
if scenarios_folder is not None:
    all_scen = scenarios_folder.GetContents('*.IntScenario')
    print(f"Escenarios encontrados: {len(all_scen)}")
    for scen in all_scen:
        nombre_scen = ga(scen, 'loc_name')
        scen.Activate()
        estado_scen = capturar_estado_red()
        cambios     = comparar_estados(estado_base, estado_scen)
        scen.Deactivate()

        scen_list_data.append({
            "Nombre":           nombre_scen,
            "Descripcion":      ga(scen, 'desc') or "",
            "Nro de cambios":   len(cambios),
        })
        for c in cambios:
            c["Escenario"] = nombre_scen
            scen_changes.append(c)

df_scen      = pd.DataFrame(scen_list_data)
df_scen_chg  = pd.DataFrame(scen_changes,
                  columns=["Escenario","Elemento","Tipo","Parametro",
                           "Valor base","Valor escenario"])

# ---------------------------------------------------------------------------
# 6. STUDY CASES (IntCase)
# ---------------------------------------------------------------------------
cases_folder = app.GetProjectFolder('study')
case_data = []
if cases_folder is not None:
    for case in cases_folder.GetContents('*.IntCase'):
        # El study case referencia un escenario y una variacion de red
        scen_ref = ga(case, 'scenario')
        var_ref  = ga(case, 'netVar')
        case_data.append({
            "Nombre":             ga(case, 'loc_name'),
            "Descripcion":        ga(case, 'desc') or "",
            "Escenario activo":   ga(scen_ref, 'loc_name') if scen_ref else "",
            "Variacion de red":   ga(var_ref,  'loc_name') if var_ref  else "",
            "Activo":             ga(case, 'isactive'),
        })
df_cases = pd.DataFrame(case_data)
print(f"Study cases: {len(df_cases)}")

# ---------------------------------------------------------------------------
# Exportar a Excel
# ---------------------------------------------------------------------------
nombre_archivo = input("\nIngresa el nombre del archivo Excel a crear (sin extension): ").strip()
if not nombre_archivo:
    nombre_archivo = f"datos_red_PF_{prj}"

output_path = rf"C:\Users\jose.lozano\OneDrive - COBEE S.A\Escritorio\Programas Python\{nombre_archivo}.xlsx"

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_buses.to_excel(writer,    sheet_name='Barras',           index=False)
    df_lines.to_excel(writer,    sheet_name='Lineas',           index=False)
    df_gens.to_excel(writer,     sheet_name='Generadores',      index=False)
    df_loads.to_excel(writer,    sheet_name='Cargas',           index=False)
    if not df_scen.empty:
        df_scen.to_excel(writer,     sheet_name='Escenarios',       index=False)
    if not df_scen_chg.empty:
        df_scen_chg.to_excel(writer, sheet_name='Escenarios_Cambios', index=False)
    if not df_cases.empty:
        df_cases.to_excel(writer,    sheet_name='StudyCases',        index=False)

print(f"\nDatos exportados a: {output_path}")

input("\nScript finalizado. PowerFactory sigue abierto. Presiona Enter para cerrar...")
