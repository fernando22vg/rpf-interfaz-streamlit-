#Prueba conexion con PF
import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2\Python\3.11")
import os
os.environ["PATH"] = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2" + os.pathsep + os.environ["PATH"]
#Comandos necesarios para llamar a PF en funcion de la ubicaion del instalador
import powerfactory as pf

app=pf.GetApplication()
#abri PF en modo engine(a veces causa conflictos) solo para demostrar que funcionan la interfaz o conexion PF-Python
app.Show()
#para activar el proyecto

user=app.GetCurrentUser()
project=app.ActivateProject("Models")
prj=app.GetActiveProject()
#Diccionario de barras o BUS
bus_dict={}
buses=app.GetCalcRelevantObjects('*.ElmTerm')
for i in buses:
    bus_dict[i.loc_name]=i


# Mantiene el script vivo mientras usas PowerFactory
# Cierra PowerFactory primero y luego presiona Enter aqui
input("PowerFactory abierto. Presiona Enter para cerrar...")
# Al ejecutarlo la terminal se congela por lo que forzaremos la salida sin esperar limpieza de PF
os._exit(0)


ldf = app.GetFromStudyCase('ComLdf')  # Get the load flow command (ComLdf)
ldf.Execute()  # Execute the load flow command

# Get a list of calculation relevant lines contained in the active project
lines = app.GetCalcRelevantObjects('*.ElmLne')

# Iterate over the list of lines
for line in lines:
    name = line.GetAttribute('loc_name')  # Get name of the current line
    loading = line.GetAttribute('c:loading')  # Get loading value of the current line
    # Print results in the output window as plain text (using f-string for formatting)
    app.PrintPlain(f'Loading of the line: "{name}" = {loading:.2f} percent.')
