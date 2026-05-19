
#Prueba conexion con PF
import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2\Python\3.11")
import os
os.environ["PATH"] = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2" + os.pathsep + os.environ["PATH"]
#Comandos necesarios para llamar a PF en funcion de la ubicaion del instalador
import pandas as pd
import powerfactory as pf

app=pf.GetApplication()
#abri PF en modo engine(a v
user=app.GetCurrentUser()
scr=app.GetCurrentScript()
prj=app.GetActiveProject()
#app.PrintPlain('hello wordl')
lineas=app.GetCalcRelevantObjects('*.ElmLne')
app=PrintPlain(lineas[0])
Name	Value	Unit	Type	Description
Tm	0,020000000000000	[s]	d	Constante de Tiempo de Medición
T_VT	0,010000000000000	[s]	d	Constante de Tiempo de filtrage de la Tensión Terminal
VT_MIN	0,85000000000000	[pu]	d	Límite Mínimo de Tensión Terminal del SCL
BETA	0,050000000000000	[-]	d	Parámetro de la curva de temporización
DT	0,85000000000000	[s]	d	Dial de Tiempo
HAB_CURVA	0,0000000000000	[-]	d	Habilita Curva de temporización
REF_SCL_PICO	1000,0000000000	[pu]	d	Referencia Instatánea
REF_SCL_TERM	1,0500000000000	[pu]	d	Referencia Térmica (Temporizada)
T_IR	1,0000000000000	[s]	d	Constante de Tiempo de filtrage de la Corriente Activa
C_0_9	0,90000000000000	[-]	d	Constante  = 0,9
IX_MIN	0,050000000000000	[pu]	d	Mínima corriente reactiva del SCL
T_IX	0,020000000000000	[s]	d	Constante de Tiempo de filtrage de la Corriente Reactiva
Kp	0,050000000000000	[pu]	d	Ganancia Proporcional
M1	-1,0000000000000	[-]	d	Constante = -1
HAB	1,0000000000000	[-]	d	Habilita Limitador
IXHOFF	0,050000000000000	[pu]	d	Constante
T0_1	0,10000000000000	[s]	d	Constante  de Tiempo = 0,1s
C_1	1,0000000000000	[-]	d	Constante = 1
C_1_1	1,0000000000000	[-]	d	Constante = 1
ALFA	0,040000000000000	[-]	d	Parámetro de la curva de temporización
set	0,50000000000000	[-]	d	Parámetro del Flip-Flop
TEMPO_DEFINIDO	60,000000000000	[s]	d	Tiempo Definido de la Temporización
TEMPO_RESET	5,0000000000000	[s]	d	Tiempo de Reset de la Referencia
TOFF	10,000000000000	[s]	d	Tiempo para Deshabilitación
TEN_U	1,0000000000000e-05	[s]	d	Tiempo para Habilitación Subexcitación
TEN_O	1,0000000000000	[s]	d	Tiempo para Habilitación Sobreexcitación
C0	0,0000000000000	[-]	d	constante = 0
LimMinSCL	-1,0000000000000	[pu]	d	Límite Mínimo de la salida de control
INF	9999,0000000000	[-]	d	constante = 9999
C1	1,0000000000000	[-]	d	Constante = 1
LimMaxSCL	1,0000000000000	[pu]	d	Límite Máximo de la salida de control

