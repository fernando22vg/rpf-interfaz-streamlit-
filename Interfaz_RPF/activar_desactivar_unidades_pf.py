#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
activar_desactivar_unidades_pf.py
---------------------------------
Selecciona semestre y evento (dialogos tkinter, estilo DatosCurvas_v3) y aplica
AUTOMATICAMENTE al escenario de operacion ACTIVO el estado de servicio de las
unidades segun las condiciones iniciales del evento:

  · pgini_MW > 0 y Fuente != mantenimiento  ->  EN SERVICIO  (outserv=0)
  · pgini_MW = 0, sin_despacho o mantenimiento -> FUERA DE SERVICIO (outserv=1)
  · Generadores de PF no listados en el Excel  -> FUERA DE SERVICIO

Cada conmutacion incluye el plant model / composite model de la unidad
(mismo mecanismo c_pmod / c_pmod2 / c_stagen que CargaCondIniciales_PF).

NO modifica pgini ni parametros DSL — SOLO outserv. Pensado para replicar la
participacion de unidades del evento sobre el escenario que estes trabajando
(p.ej. mientras modificas modelos DSL).

Fuente de datos (en orden de preferencia):
  1. datos_cargados_Ev{N}.xlsx      (hoja pgini_GEN_FINAL — estado exacto cargado)
  2. condiciones_iniciales_*.xlsx   (hoja pgini_GEN)

Registro de cambios: {evento}\\cambios_unidades_Ev{N}.log
"""

import os
import re
import sys
import glob
from datetime import datetime

# ══════════════════════════════════════════════════════════════
RAIZ    = r"C:\Datos del CNDC\01_INFO CNDC_RPF"
PF_BASE = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2"
# ══════════════════════════════════════════════════════════════

# ── Conexion a PowerFactory (embebido o externo) ─────────────────────────────
try:
    import powerfactory as pf
except ImportError:
    _pf_py = os.path.join(PF_BASE, "Python",
                          f"{sys.version_info.major}.{sys.version_info.minor}")
    sys.path.append(_pf_py)
    import powerfactory as pf


# ─────────────────────────────────────────────────────────────
# Dialogo tkinter (mismo estilo que DatosCurvas_v3.elegir)
# ─────────────────────────────────────────────────────────────
def elegir(opciones, titulo):
    """Listbox tkinter de seleccion simple. Retorna la opcion o None si se cancela."""
    import tkinter as tk

    resultado = [None]

    root = tk.Tk()
    root.title(titulo)
    root.resizable(False, False)
    root.attributes('-topmost', True)

    tk.Label(root, text=titulo, font=('Arial', 10, 'bold')).pack(pady=(12, 4), padx=16)

    frame = tk.Frame(root)
    frame.pack(padx=16, pady=4)

    scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
    listbox = tk.Listbox(
        frame, selectmode=tk.SINGLE,
        width=55, height=min(len(opciones), 16),
        yscrollcommand=scrollbar.set,
        activestyle='dotbox',
    )
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH)

    for op in opciones:
        listbox.insert(tk.END, op)
    listbox.select_set(0)
    listbox.focus_set()

    def confirmar(event=None):
        sel = listbox.curselection()
        if sel:
            resultado[0] = opciones[sel[0]]
            root.destroy()

    def cancelar(event=None):
        root.destroy()

    tk.Button(root, text="Aceptar", command=confirmar, width=14).pack(pady=(6, 12))
    root.bind('<Return>', confirmar)
    root.bind('<Escape>', cancelar)
    listbox.bind('<Double-Button-1>', confirmar)
    root.protocol("WM_DELETE_WINDOW", cancelar)

    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    root.mainloop()
    return resultado[0]


# ─────────────────────────────────────────────────────────────
# Helpers PowerFactory
# ─────────────────────────────────────────────────────────────
def get_comp(gen):
    """Composite/plant model del generador (c_pmod ElmSym; c_pmod2/c_stagen ElmGenstat)."""
    for attr in ("c_pmod", "c_pmod2", "c_stagen"):
        try:
            c = gen.GetAttribute(attr)
            if c is not None:
                return c
        except Exception:
            pass
    return None


def set_unidad(gen, valor):
    """Conmuta maquina + composite model. valor: 0=en servicio, 1=fuera de servicio."""
    gen.outserv = valor
    comp = get_comp(gen)
    if comp is not None:
        try:
            comp.outserv = valor
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────
# Lectura del Excel de condiciones del evento
# ─────────────────────────────────────────────────────────────
def leer_unidades_evento(ev_path, n_ev):
    """Retorna (dict loc_name -> (en_servicio: bool, pgini, fuente), archivo_usado)."""
    # Preferir datos_cargados (estado exacto post-carga); fallback condiciones_iniciales
    candidatos = []
    dc = os.path.join(ev_path, f"datos_cargados_Ev{n_ev}.xlsx")
    if os.path.isfile(dc):
        candidatos.append((dc, "pgini_GEN_FINAL"))
    for ci in sorted(glob.glob(os.path.join(ev_path, "condiciones_iniciales_*.xlsx"))):
        candidatos.append((ci, "pgini_GEN"))

    if not candidatos:
        raise FileNotFoundError(
            "No se encontro datos_cargados_Ev{N}.xlsx ni condiciones_iniciales_*.xlsx "
            f"en: {ev_path}")

    import pandas as pd
    path, hoja = candidatos[0]
    df = pd.read_excel(path, sheet_name=hoja)

    unidades = {}
    for _, fila in df.iterrows():
        loc    = str(fila["loc_name PF"]).strip()
        pgini  = float(fila.get("pgini_MW", 0) or 0)
        fuente = str(fila.get("Fuente", "")).strip()
        en_serv = (pgini > 0) and (fuente.lower() != "mantenimiento")
        unidades[loc] = (en_serv, pgini, fuente)
    return unidades, os.path.basename(path)


# ─────────────────────────────────────────────────────────────
# Programa principal
# ─────────────────────────────────────────────────────────────
def main():
    app = pf.GetApplication()
    if app is None:
        print("[ERROR] No se pudo conectar a PowerFactory.")
        return
    app.ClearOutputWindow()

    proyecto = app.GetActiveProject()
    if proyecto is None:
        print("[ERROR] No hay proyecto activo en PowerFactory.")
        return

    escenario = app.GetActiveScenario()
    nombre_esc = escenario.loc_name.strip() if escenario else "(ninguno)"
    print(f"  Proyecto activo : {proyecto.loc_name.strip()}")
    print(f"  Escenario activo: {nombre_esc}")

    # ── Seleccion de semestre y evento ────────────────────────────────────────
    semestres = sorted(d for d in os.listdir(RAIZ)
                       if os.path.isdir(os.path.join(RAIZ, d)))
    semestre = elegir(semestres, "Semestre de estudio")
    if semestre is None:
        return
    base_ev = os.path.join(RAIZ, semestre, "Análisis_todos_los_eventos")
    if not os.path.isdir(base_ev):
        base_ev = os.path.join(RAIZ, semestre, "Analisis_todos_los_eventos")
    eventos = sorted(d for d in os.listdir(base_ev)
                     if os.path.isdir(os.path.join(base_ev, d)))
    evento = elegir(eventos, "Evento")
    if evento is None:
        return
    ev_path = os.path.join(base_ev, evento)
    m = re.search(r"(\d+)\s*$", evento)
    n_ev = m.group(1) if m else evento.split()[-1]

    # ── Leer participacion de unidades del Excel ──────────────────────────────
    try:
        unidades, archivo = leer_unidades_evento(ev_path, n_ev)
    except Exception as e:
        print(f"[ERROR] {e}")
        return
    n_on_excel = sum(1 for v in unidades.values() if v[0])
    print(f"\n  Fuente de datos : {archivo}")
    print(f"  Unidades listadas: {len(unidades)}  "
          f"(EN SERVICIO: {n_on_excel}  |  FUERA: {len(unidades) - n_on_excel})")

    # ── Aplicar al escenario activo ───────────────────────────────────────────
    log_path = os.path.join(ev_path, f"cambios_unidades_Ev{n_ev}.log")

    def log(msg):
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"  {msg}")
        try:
            app.PrintPlain(msg)          # resumen visible en el Output Window de PF
        except Exception:
            pass
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{stamp}] {msg}\n")
        except OSError:
            pass

    gens = (app.GetCalcRelevantObjects("*.ElmSym", 1)     or []) + \
           (app.GetCalcRelevantObjects("*.ElmGenstat", 1) or [])

    log(f"=== Aplicando estados de '{semestre} / {evento}' "
        f"sobre escenario '{nombre_esc}' ({archivo}) ===")

    def aplicar(gen, objetivo, motivo):
        """Lleva maquina Y plant model al estado objetivo. Verifica cada uno por
        separado: corrige el comp aunque la maquina ya este en el estado correcto
        (caso tipico: maquina fuera de servicio con plant model aun activo).
        Retorna (cambio_gen, cambio_comp)."""
        loc  = gen.loc_name.strip()
        comp = get_comp(gen)
        cambio_gen  = getattr(gen, "outserv", 0) != objetivo
        cambio_comp = comp is not None and getattr(comp, "outserv", 0) != objetivo
        if cambio_gen:
            gen.outserv = objetivo
        if cambio_comp:
            try:
                comp.outserv = objetivo
            except Exception as e:
                log(f"  [ERROR] {loc}: no se pudo conmutar plant model "
                    f"'{comp.loc_name.strip()}': {e}")
                cambio_comp = False
        tag = "ON " if objetivo == 0 else "OFF"
        if cambio_gen and cambio_comp:
            log(f"  [{tag}] {loc:<24} -> maquina + plant model  ({motivo})")
        elif cambio_gen:
            extra = "sin plant model" if comp is None else "plant model ya correcto"
            log(f"  [{tag}] {loc:<24} -> maquina  ({extra}) ({motivo})")
        elif cambio_comp:
            log(f"  [{tag}] {loc:<24} -> SOLO plant model '{comp.loc_name.strip()}' "
                f"(maquina ya estaba) ({motivo})")
        return cambio_gen, cambio_comp

    n_act = n_desact = n_sin_cambio = n_no_listados = n_comp_fix = 0
    for gen in sorted(gens, key=lambda g: g.loc_name.strip().upper()):
        loc = gen.loc_name.strip()
        info = unidades.get(loc)
        if info is None:
            # No participa en el evento -> fuera de servicio (igual que la carga)
            cg, cc = aplicar(gen, 1, "no listado en el evento")
            if cg:
                n_no_listados += 1
            elif cc:
                n_comp_fix += 1
            else:
                n_sin_cambio += 1
            continue

        en_serv, pgini, fuente = info
        objetivo = 0 if en_serv else 1
        cg, cc = aplicar(gen, objetivo, f"pgini={pgini:.2f} MW, {fuente}")
        if cg:
            if objetivo == 0:
                n_act += 1
            else:
                n_desact += 1
        elif cc:
            n_comp_fix += 1
        else:
            n_sin_cambio += 1

    log(f"=== Resumen: {n_act} activadas | {n_desact} desactivadas | "
        f"{n_no_listados} no-listadas apagadas | "
        f"{n_comp_fix} plant models corregidos (maquina ya correcta) | "
        f"{n_sin_cambio} ya correctas ===")

    # ── Guardar escenario (opcional) ──────────────────────────────────────────
    if escenario is not None:
        guardar = elegir(["Si — guardar escenario", "No — dejar sin guardar"],
                         f"¿Guardar escenario '{nombre_esc}'?")
        if guardar and guardar.startswith("Si"):
            try:
                escenario.Save()
                log(f"Escenario guardado: {nombre_esc}")
            except Exception as e:
                print(f"  [ERROR] No se pudo guardar el escenario: {e}")

    print(f"\n  Registro completo: {log_path}")


if __name__ == "__main__":
    main()
