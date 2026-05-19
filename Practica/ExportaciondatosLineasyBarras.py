#Prueba conexion con PF
import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2\Python\3.11")
import os
os.environ["PATH"] = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2" + os.pathsep + os.environ["PATH"]
#Comandos necesarios para llamar a PF en funcion de la ubicaion del instalador
import pandas as pd
import powerfactory as pf

# Mata cualquier instancia previa de PowerFactory antes de conectar
os.system("taskkill /f /im PowerFactory.exe >nul 2>&1")

app = pf.GetApplication()
if app is None:
    raise RuntimeError(
        "No se pudo obtener la aplicacion de PowerFactory. "
        "Cierre PowerFactory si esta abierto y vuelva a ejecutar el script."
    )
#abri PF en modo engine(a veces causa conflictos) solo para demostrar que funcionan la interfaz o conexion PF-Python
app.Show()
#para activar el proyecto

user=app.GetCurrentUser()
prj=app.ActivateProject("Pruebas")

print(f"Proyecto activo: {prj}")

#Diccionario de barras o BUS
bus_dict={}
buses=app.GetCalcRelevantObjects('*.ElmTerm')
for i in buses:
    bus_dict[i.loc_name]=i

print(f"Barras encontradas: {len(bus_dict)}")
for nombre in bus_dict:
    print(f"  - {nombre}")

line_dict={}
line=app.GetCalcRelevantObjects('*.ElmLne')
for i in line:
    line_dict[i.loc_name]=i
print(f"Lineas encontradas: {len(line_dict)}")

ldf = app.GetFromStudyCase('ComLdf')  # llama al comado load flow (ComLdf)
ldf.Execute()  # Ejecuta el comandoload flow 

# Iterate over the list of lines (con distancia a 2 decimales)
for i in line:
    name = i.GetAttribute('loc_name')
    distance = i.GetAttribute('dline')
    try:
        loading = i.GetAttribute('c:loading')
        if loading is None:
            app.PrintPlain(f'Loading of the line: "{name}" = Sin resultado (fuera de servicio).  - Distancia: {distance:.2f} km')
        else:
            app.PrintPlain(f'Loading of the line: "{name}" = {loading:.2f} %  - Distancia: {distance:.2f} km')
    except AttributeError:
        app.PrintPlain(f'Loading of the line: "{name}" = Sin resultado (fuera de servicio).  - Distancia: {distance:.2f} km')

for n in line_dict:
    print(f"  - {n} - Distancia: {line_dict[n].GetAttribute('dline'):.2f} km")

# --- Exportar datos a Excel ---

# Datos de barras
bus_data = []
for nombre, bus in bus_dict.items():
    uknom  = bus.GetAttribute('uknom')
    u_pu   = bus.GetAttribute('m:u')
    u_kv   = round(u_pu * uknom, 4) if u_pu is not None and uknom is not None else None
    bus_data.append({
        "Nombre":            nombre,
        "Tension nom. (kV)": uknom,
        "Tension (pu)":      u_pu,
        "Tension (kV)":      u_kv,
        "Angulo (deg)":      bus.GetAttribute('m:phiu'),
        "En servicio":       bool(bus.GetAttribute('outserv') == 0),
    })
df_buses = pd.DataFrame(bus_data)

# Helper para leer atributos sin romper el script si no existen
def get_attr(obj, attr):
    try:
        return obj.GetAttribute(attr)
    except AttributeError:
        return None

# Datos de lineas
line_data = []
for nombre, ln in line_dict.items():
    loading   = get_attr(ln, 'c:loading')
    typ       = get_attr(ln, 'typ_id')          # Tipo de linea (TypLne)
    uline     = get_attr(typ, 'uline') if typ is not None else None   # Tension nom. del tipo
    inom      = get_attr(typ, 'InomAC') if typ is not None else None  # Corriente nom. del tipo
    line_data.append({
        "Nombre":            nombre,
        "Distancia (km)":    round(get_attr(ln, 'dline'), 2),
        "Tension nom. (kV)": uline,
        "Corriente nom. (A)":inom,
        "Carga (%)":         round(loading, 2) if loading is not None else None,
        "En servicio":       bool(get_attr(ln, 'outserv') == 0),
    })
df_lines = pd.DataFrame(line_data)

# Pedir nombre del archivo Excel al usuario
nombre_archivo = input("\nIngresa el nombre del archivo Excel a crear (sin extension): ").strip()
if not nombre_archivo:
    nombre_archivo = "resultados_PF"

# Guardar en Excel con dos hojas
output_path = rf"C:\Users\jose.lozano\OneDrive - COBEE S.A\Escritorio\Programas Python\{nombre_archivo}.xlsx"
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_buses.to_excel(writer, sheet_name='Barras', index=False)
    df_lines.to_excel(writer, sheet_name='Lineas', index=False)

print(f"\nDatos exportados a: {output_path}")

input("\nScript finalizado. PowerFactory sigue abierto. Presiona Enter para cerrar...")

# Cierra PowerFactory correctamente para liberar el proceso engine
#app.Exit()
