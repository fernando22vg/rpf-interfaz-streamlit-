#Prueba conexion con PF
import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2\Python\3.11")
import os
os.environ["PATH"] = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2" + os.pathsep + os.environ["PATH"]
#Comandos necesarios para llamar a PF en funcion de la ubicaion del instalador
import pandas as pd
import powerfactory as pf

app=pf.GetApplication()
#abri PF en modo engine(a veces causa conflictos) solo para demostrar que funcionan la interfaz o conexion PF-Python
#app.Show()
#para activar el proyecto

user=app.GetCurrentUser()
scr=uapp.GetCurrentScript()
project=app.ActivateProject()

app=PrintPlain("hello wordl")


