#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
OrdenadorDatosEvento.py
-----------------------
Extrae y organiza los datos de frecuencia y potencia del evento real
(archivo "1 seg" de COBEE/CNDC) en CSVs individuales por unidad.

Flujo:
  1. Seleccion de semestre y evento (terminal, igual que otros programas)
  2. Lee la fecha/hora del evento desde Tabla_Eventos_*.xlsx
  3. Busca la carpeta "FALLA DD.MM.YY HRS HH.MM" en 02_DATOS CNDC_RPF
  4. Lee el archivo "1 seg.DD.MM.YY_hrs.HH.MM.xls"
  5. Genera un CSV por unidad COBEE con columnas: Tiempo_s, Frecuencia_Hz, {Unidad}

Salida: CARPETA_SAL\{semestre}\{evento}\{unidad}.csv
"""

import os
import re
import sys
import glob
from datetime import datetime, date

import pandas as pd
import openpyxl

# ══════════════════════════════════════════════════════════════
RAIZ_RPF   = r"C:\Datos del CNDC\01_INFO CNDC_RPF"
RAIZ_DATOS = r"C:\Datos del CNDC\02_DATOS CNDC_RPF"
# ══════════════════════════════════════════════════════════════


# ─────────────────────────────────────────────────────────────
# Utilidades
# ─────────────────────────────────────────────────────────────
def separador(titulo="", ancho=62):
    if titulo:
        print(f"\n{'='*ancho}")
        print(f"  {titulo}")
        print(f"{'='*ancho}")
    else:
        print(f"{'='*ancho}")


def _sanear_nombre(txt):
    txt = str(txt or '').strip()
    for ch in r'\/:*?"<>|':
        txt = txt.replace(ch, '_')
    return txt.strip().strip('.') or 'sin_nombre'


def _float(val, default=0.0):
    try:
        v = float(val)
        return v if v == v else default   # NaN check
    except Exception:
        return default


# ─────────────────────────────────────────────────────────────
# Selección interactiva (terminal)
# ─────────────────────────────────────────────────────────────
def elegir(opciones, titulo):
    print(f"\n{titulo}:")
    for i, op in enumerate(opciones, 1):
        print(f"  {i}. {op}")
    while True:
        try:
            sel = int(input("  Seleccionar numero: "))
            if 1 <= sel <= len(opciones):
                return opciones[sel - 1]
        except ValueError:
            pass
        print("  Opcion invalida, intente de nuevo.")


def seleccionar_semestre_evento():
    semestres = sorted(d for d in os.listdir(RAIZ_RPF)
                       if os.path.isdir(os.path.join(RAIZ_RPF, d)))
    if not semestres:
        raise RuntimeError(f"No hay semestres en {RAIZ_RPF}")
    semestre = elegir(semestres, "Semestre de estudio")

    base_ev = os.path.join(RAIZ_RPF, semestre, "Análisis_todos_los_eventos")
    if not os.path.isdir(base_ev):
        raise RuntimeError(f"No existe: {base_ev}")
    eventos = sorted(d for d in os.listdir(base_ev)
                     if os.path.isdir(os.path.join(base_ev, d)))
    if not eventos:
        raise RuntimeError(f"No hay eventos en {base_ev}")
    evento = elegir(eventos, "Evento")

    m = re.search(r"(\d+)", evento)
    n_evento = int(m.group(1)) if m else None

    return semestre, evento, n_evento


# ─────────────────────────────────────────────────────────────
# Leer fecha/hora del evento desde Tabla_Eventos
# ─────────────────────────────────────────────────────────────
def _extraer_datetime_de_fila(fila):
    """
    Recorre las celdas de una fila buscando un valor que pueda ser
    interpretado como fecha (con o sin hora).
    Devuelve un datetime o None.
    """
    candidatos = []
    for val in fila:
        if val is None:
            continue
        if isinstance(val, datetime):
            candidatos.append(val)
        elif isinstance(val, date):
            candidatos.append(datetime(val.year, val.month, val.day))
        elif isinstance(val, str):
            val = val.strip()
            formatos = [
                '%d/%m/%Y %H:%M', '%d/%m/%Y %H:%M:%S',
                '%d-%m-%Y %H:%M', '%d-%m-%Y %H:%M:%S',
                '%Y-%m-%d %H:%M', '%Y-%m-%d %H:%M:%S',
                '%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d',
            ]
            for fmt in formatos:
                try:
                    candidatos.append(datetime.strptime(val, fmt))
                    break
                except ValueError:
                    pass

    # Preferir el candidato con hora != 00:00 (contiene hora de la falla)
    for dt in candidatos:
        if dt.hour != 0 or dt.minute != 0:
            return dt
    return candidatos[0] if candidatos else None


def leer_fecha_evento(semestre, n_evento):
    patron = os.path.join(RAIZ_RPF, semestre, "Tabla_Eventos_*.xlsx")
    archivos = glob.glob(patron)
    if not archivos:
        return None

    wb = openpyxl.load_workbook(archivos[0], data_only=True)
    ws = wb.active

    for fila in ws.iter_rows(min_row=3, values_only=True):
        if fila[0] is None:
            continue
        try:
            if int(fila[0]) == n_evento:
                return _extraer_datetime_de_fila(fila[1:])
        except (ValueError, TypeError):
            pass

    return None


# ─────────────────────────────────────────────────────────────
# Buscar carpeta FALLA en 02_DATOS CNDC_RPF
# ─────────────────────────────────────────────────────────────
# Soporta año de 2 o 4 dígitos y espacio opcional entre HRS y la hora:
#   FALLA 04.02.2025 HRS21.05
#   FALLA 02.03.24 HRS 00.56
_PAT_FALLA = re.compile(
    r'FALLA\s+(\d{2})\.(\d{2})\.(\d{2,4})\s+HRS\s*(\d{2})\.(\d{2})',
    re.IGNORECASE,
)


def buscar_carpeta_falla(fecha_evento):
    """
    Busca en RAIZ_DATOS/{year}/ una carpeta con formato
    'FALLA DD.MM.YYYY HRS HH.MM' (o YY) que coincida con fecha_evento.
    Si hay varias, devuelve la de menor diferencia horaria.
    """
    yr4 = fecha_evento.year
    year_dir = os.path.join(RAIZ_DATOS, str(yr4))

    # Si no existe la carpeta del año exacto, buscar en todos los años
    if not os.path.isdir(year_dir):
        for d in sorted(os.listdir(RAIZ_DATOS)):
            candidate = os.path.join(RAIZ_DATOS, d)
            if os.path.isdir(candidate) and d.isdigit():
                yr4 = int(d)
                year_dir = candidate
                break

    if not os.path.isdir(year_dir):
        return None

    yr2 = yr4 % 100          # año de 2 dígitos para comparar
    dd  = fecha_evento.day
    mm  = fecha_evento.month
    hh  = fecha_evento.hour
    mi  = fecha_evento.minute

    candidatos = []
    for folder in os.listdir(year_dir):
        full = os.path.join(year_dir, folder)
        if not os.path.isdir(full):
            continue
        m = _PAT_FALLA.match(folder)
        if not m:
            continue
        fd_s, fm_s, fy_s, fh_s, fmi_s = m.groups()
        fd, fm, fh, fmi = int(fd_s), int(fm_s), int(fh_s), int(fmi_s)
        fy = int(fy_s) % 100    # normalizar a 2 dígitos (25 o 2025 → 25)

        if fd == dd and fm == mm and fy == yr2:
            diff = abs(fh * 60 + fmi - hh * 60 - mi)
            candidatos.append((diff, full, folder))

    if not candidatos:
        return None

    candidatos.sort()
    _, best_path, best_name = candidatos[0]
    print(f"  Carpeta FALLA encontrada : {best_name}")
    return best_path


# ─────────────────────────────────────────────────────────────
# Buscar y leer el archivo "1 seg"
# ─────────────────────────────────────────────────────────────
def buscar_archivo_1seg(carpeta_falla):
    patrones = [
        '1 seg.*.xls', '1 seg.*.xlsx',
        '1seg*.xls',   '1seg*.xlsx',
        '1 Seg.*.xls', '1 Seg.*.xlsx',
        '1 SEG.*.xls', '1 SEG.*.xlsx',
    ]
    for p in patrones:
        hits = glob.glob(os.path.join(carpeta_falla, p))
        if hits:
            return hits[0]

    # Fallback: búsqueda case-insensitive sobre todos los archivos de la carpeta
    try:
        for fname in os.listdir(carpeta_falla):
            lower = fname.lower()
            if lower.startswith('1') and 'seg' in lower and lower.endswith(('.xls', '.xlsx')):
                return os.path.join(carpeta_falla, fname)
    except OSError:
        pass
    return None


def leer_datos_1seg(excel_path):
    """
    Estructura del archivo:
      Fila 3 (índice 2) → nombres de columnas / unidades COBEE
      Columna B (índice 1) → Tiempo [s]
      Columna C (índice 2) → Frecuencia [Hz]
      Columna D+ (índice 3+) → Potencia por unidad [MW]

    Devuelve: (tiempo_series, frecuencia_series, {nombre_unidad: potencia_series})
    """
    ext = os.path.splitext(excel_path)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'

    raw = pd.read_excel(excel_path, header=None, engine=engine)

    if raw.shape[0] < 4:
        raise ValueError("El archivo tiene menos de 4 filas")

    # Fila 3 (índice 2): encabezados / nombres de unidades
    header_row = raw.iloc[2]

    # Datos desde fila 4 (índice 3)
    data = raw.iloc[3:].reset_index(drop=True)

    tiempo     = data.iloc[:, 1]   # columna B
    frecuencia = data.iloc[:, 2]   # columna C

    unidades = {}
    for i in range(3, len(header_row)):
        nombre = str(header_row.iloc[i]).strip()
        if nombre and nombre.lower() not in ('nan', 'none', ''):
            unidades[nombre] = data.iloc[:, i]

    return tiempo, frecuencia, unidades


# ─────────────────────────────────────────────────────────────
# Exportar CSVs
# ─────────────────────────────────────────────────────────────
def exportar_csvs(carpeta_sal, tiempo, frecuencia, unidades):
    os.makedirs(carpeta_sal, exist_ok=True)

    # Máscara para filas con tiempo válido
    mask = tiempo.notna() & (tiempo != '')
    t = tiempo[mask].reset_index(drop=True)
    f = frecuencia[mask].reset_index(drop=True)

    exportados = []
    for nombre_unidad, potencia in unidades.items():
        p = potencia[mask].reset_index(drop=True)
        
        df_out = pd.DataFrame({
            'Tiempo_s':      t,
            'Frecuencia_Hz': f,
            nombre_unidad:   p,
        })

        nombre_arch_xlsx = _sanear_nombre(nombre_unidad) + '.xlsx'
        ruta_xlsx = os.path.join(carpeta_sal, nombre_arch_xlsx)
        df_out.to_excel(ruta_xlsx, index=False)
        exportados.append(nombre_arch_xlsx)

    return exportados


# ─────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────
def main():
    separador("ORDENADOR DE DATOS DE EVENTO REAL")

    # [1] Selección de semestre y evento
    semestre, evento, n_evento = seleccionar_semestre_evento()
    print(f"\n  Semestre : {semestre}")
    print(f"  Evento   : {evento}  (N={n_evento})")

    # [2] Fecha del evento desde Tabla_Eventos
    separador("LEYENDO TABLA DE EVENTOS")
    fecha_evento = leer_fecha_evento(semestre, n_evento)
    if fecha_evento is None:
        print(f"  [ERROR] No se encontro la fecha del evento {n_evento}.")
        tabla_dir = os.path.join(RAIZ_RPF, semestre)
        print(f"          Verificar Tabla_Eventos_*.xlsx en: {tabla_dir}")
        sys.exit(1)
    print(f"  Fecha del evento : {fecha_evento.strftime('%d/%m/%Y %H:%M')}")

    # [3] Buscar carpeta FALLA
    separador("BUSCANDO CARPETA DE FALLA")
    carpeta_falla = buscar_carpeta_falla(fecha_evento)
    if carpeta_falla is None:
        yr2 = fecha_evento.year % 100
        dd, mm = fecha_evento.day, fecha_evento.month
        print(f"  [ERROR] No se encontro carpeta FALLA {dd:02d}.{mm:02d}.{yr2:02d}*")
        print(f"          Verificar en: {os.path.join(RAIZ_DATOS, str(fecha_evento.year))}")
        sys.exit(1)

    # [4] Buscar y leer archivo "1 seg"
    separador("LEYENDO ARCHIVO DE DATOS")
    archivo_1seg = buscar_archivo_1seg(carpeta_falla)
    if archivo_1seg is None:
        print(f"  [ERROR] No se encontro archivo '1 seg.*' en:")
        print(f"          {carpeta_falla}")
        try:
            archivos = os.listdir(carpeta_falla)
            print(f"  Archivos en la carpeta ({len(archivos)}):")
            for _f in sorted(archivos):
                print(f"    - {_f}")
        except OSError:
            pass
        sys.exit(1)

    print(f"  Archivo : {os.path.basename(archivo_1seg)}")
    tiempo, frecuencia, unidades = leer_datos_1seg(archivo_1seg)

    n_registros = int(tiempo.notna().sum())
    print(f"  Registros        : {n_registros}")
    print(f"  Unidades COBEE   : {len(unidades)}")
    for u in unidades:
        print(f"    - {u}")

    # [5] Exportar CSVs
    carpeta_sal = os.path.join(
        RAIZ_RPF, semestre,
        "Análisis_todos_los_eventos", evento,
        "Graficas Registro 1SEG COBEE",
    )
    separador("EXPORTANDO EXCEL")
    print(f"  Carpeta salida : {carpeta_sal}")
    exportados = exportar_csvs(carpeta_sal, tiempo, frecuencia, unidades) # La función ahora exporta XLSX

    print(f"\n  Excel generados : {len(exportados)}")
    for f_xlsx in exportados:
        print(f"    ✓  {f_xlsx}")

    separador()
    print("  Proceso completado.")
    separador()


if __name__ == '__main__':
    main()
