#!/usr/bin/env python3
"""
diagnostico_disparo_cc.py

Detecta cargas a PowerFactory (datos_cargados_Ev{n}*.xlsx) afectadas por el
bug de emparejamiento del prefijo CNDC "CC" en el campo Disparo (ej.
"CCERI30" no emparejaba con el loc_name PF "sym_ERI30"), corregido en
CargaCondIniciales_PF.py.

Senal de la falla: Disparo empieza con "CC" + letra mayuscula, y
'Suma pgini disparo (MW)' quedo en 0 con p_desc > 0 (la unidad nunca se
verifico/corrigio contra el p_desc registrado del evento).
"""
import os
import re
import glob

import pandas as pd

RAIZ = r"C:\Datos del CNDC\01_INFO CNDC_RPF"


def revisar_archivo(path):
    try:
        xl = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        return None
    hoja = "Resumen_Cargado" if "Resumen_Cargado" in xl.sheet_names else (
        "Resumen_Ajustado" if "Resumen_Ajustado" in xl.sheet_names else None)
    if hoja is None:
        return None
    df = xl.parse(hoja, header=None)
    info = {}
    for _, row in df.iterrows():
        if len(row) < 2:
            continue
        info[str(row.iloc[0]).strip()] = row.iloc[1]
    disparo = str(info.get("Disparo", "")).strip()
    p_desc = info.get("p_desc registrado (MW)")
    suma = info.get("Suma pgini disparo (MW)")
    try:
        p_desc = float(p_desc)
        suma = float(suma)
    except (TypeError, ValueError):
        return None
    afectado = bool(re.match(r"^CC[A-Z]", disparo)) and p_desc > 0 and abs(suma) < 0.01
    return {
        "disparo": disparo, "p_desc": p_desc, "suma": suma, "afectado": afectado,
    }


def main():
    semestres = sorted(d for d in os.listdir(RAIZ) if os.path.isdir(os.path.join(RAIZ, d)))
    afectados = []
    for semestre in semestres:
        base_ev = os.path.join(RAIZ, semestre, "Análisis_todos_los_eventos")
        if not os.path.isdir(base_ev):
            continue
        eventos = sorted(
            (d for d in os.listdir(base_ev) if os.path.isdir(os.path.join(base_ev, d))),
            key=lambda d: int(m.group(1)) if (m := re.search(r"(\d+)$", d)) else -1)
        for ev in eventos:
            ev_path = os.path.join(base_ev, ev)
            archivos = sorted(glob.glob(os.path.join(ev_path, "datos_cargados_Ev*.xlsx")))
            for arch in archivos:
                res = revisar_archivo(arch)
                if res is None:
                    continue
                marca = "AFECTADO" if res["afectado"] else "ok"
                print(f"  [{marca:<9}] {semestre} / {ev:<10} {os.path.basename(arch):<35} "
                      f"Disparo={res['disparo']:<12} p_desc={res['p_desc']:>7.2f}  "
                      f"suma_disparo={res['suma']:>7.2f}")
                if res["afectado"]:
                    afectados.append((semestre, ev, os.path.basename(arch), res["disparo"], res["p_desc"]))

    print("\n" + "=" * 78)
    if not afectados:
        print("Ningun datos_cargados_Ev*.xlsx afectado por el bug del prefijo CC.")
    else:
        print(f"{len(afectados)} carga(s) a PowerFactory afectada(s) por el bug del prefijo CC:")
        for semestre, ev, arch, disparo, p_desc in afectados:
            print(f"    {semestre} / {ev}  ({arch})  Disparo={disparo}  p_desc={p_desc:.2f} MW"
                  f"  -> re-cargar con CargaCondIniciales_PF.py corregido")


if __name__ == "__main__":
    main()
