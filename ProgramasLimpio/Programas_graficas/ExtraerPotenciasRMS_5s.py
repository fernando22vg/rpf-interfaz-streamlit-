#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Extrae la potencia activa de todas las unidades generadoras en t = 5 s
desde resultados RMS de DIgSILENT PowerFactory.

Estrategia:
1) toma el ElmRes activo del caso de estudio,
2) usa ComRes para exportar a CSV la serie temporal de tiempo + potencia
   de cada generador,
3) lee el CSV,
4) devuelve el valor exactamente en t=5 s o interpola linealmente si
   no existe una muestra exacta.

Ejecutar dentro de PowerFactory con el caso de estudio ya calculado.
"""

import csv
import os
from typing import Dict, List, Optional, Sequence, Tuple

import powerfactory as pf

# =============================================================================
# CONFIGURACION
# =============================================================================

TARGET_TIME_S = 5.0
TIME_TOL_S = 1e-6
INTERPOLATE_IF_NEEDED = True

# Orden de intentos para cada generador. Se exportan todos y luego se elige
# el primer candidato que produzca datos válidos.
POWER_VARIABLE_CANDIDATES = (
    "m:Psum:bus1",
    "m:P:bus1",
    "m:P1:bus1",
    "m:P",
)

GENERATOR_CLASSES = (
    "*.ElmSym",
    "*.ElmGenstat",
    "*.ElmPvsys",
    "*.ElmWind",
    "*.ElmAsm",
)

OUTPUT_BASENAME = "potencias_generadores_t_5s.csv"


# =============================================================================
# HELPERS BASICOS
# =============================================================================

def _get_loc_name(obj, default=""):
    if obj is None:
        return default
    try:
        return obj.GetAttribute("loc_name")
    except Exception:
        try:
            return obj.loc_name
        except Exception:
            return default or str(obj)


def _get_class_name(obj, default=""):
    if obj is None:
        return default
    try:
        return obj.GetClassName()
    except Exception:
        return default


def _get_attr_safe(obj, attr, default=None):
    try:
        return obj.GetAttribute(attr)
    except Exception:
        try:
            return getattr(obj, attr)
        except Exception:
            return default


def _obj_key(obj):
    if obj is None:
        return "None"
    class_name = _get_class_name(obj, "")
    loc_name = _get_loc_name(obj, "")
    for getter in ("GetFullName", "GetFullPath"):
        try:
            value = getattr(obj, getter)()
            if value:
                return f"{class_name}::{value}"
        except Exception:
            pass
    if loc_name or class_name:
        return f"{class_name}::{loc_name}"
    return str(obj)


def _float_or_none(value):
    if value is None:
        return None
    try:
        txt = str(value).strip().replace(",", ".")
        if txt == "":
            return None
        return float(txt)
    except Exception:
        return None


def _safe_float(value, default=None):
    try:
        val = float(value)
        if val != val:
            return default
        return val
    except Exception:
        return default


def _script_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def _ensure_dir(path):
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)


# =============================================================================
# POWERFACTORY / RESULTADOS
# =============================================================================

def _get_application():
    app = pf.GetApplication()
    if app is None:
        raise RuntimeError("No se pudo obtener la aplicacion de PowerFactory.")
    return app


def _get_active_result_file(sc):
    elmres_list = sc.GetContents("*.ElmRes", 1) or []
    if not elmres_list:
        return None

    for elmres in elmres_list:
        if _get_loc_name(elmres).strip().lower() == "all calculations":
            return elmres

    return elmres_list[0]


def _get_or_create_comres(sc):
    comres_list = sc.GetContents("*.ComRes", 1) or []
    if comres_list:
        return comres_list[0], False

    try:
        comres = sc.CreateObject("ComRes", "__tmp_rms_5s_export__")
        if comres is not None:
            return comres, True
    except Exception:
        pass

    return None, False


def _get_generators(app):
    found = {}
    for cls in GENERATOR_CLASSES:
        try:
            objs = app.GetCalcRelevantObjects(cls) or []
        except Exception:
            objs = []
        for obj in objs:
            key = _obj_key(obj)
            if key not in found:
                found[key] = obj
    return list(found.values())


def _build_export_series(generators, elmres):
    """
    Devuelve una lista de diccionarios en el mismo orden que saldrán las
    columnas del CSV. La primera columna siempre es el tiempo.
    """
    series = [{
        "kind": "time",
        "element": None,
        "variable": "b:tnow",
        "gen_key": None,
        "gen_name": "time_s",
        "candidate_index": None,
        "candidate_variable": "b:tnow",
        "resultobj": elmres,
    }]

    for gen in generators:
        gen_key = _obj_key(gen)
        gen_name = _get_loc_name(gen, gen_key)
        for idx, var in enumerate(POWER_VARIABLE_CANDIDATES):
            series.append({
                "kind": "power",
                "element": gen,
                "variable": var,
                "gen_key": gen_key,
                "gen_name": gen_name,
                "candidate_index": idx,
                "candidate_variable": var,
                "resultobj": elmres,
            })

    return series


def _capture_comres_state(comres):
    state = {}
    for attr in (
        "pResult", "f_name", "iopt_exp", "iopt_sep", "iopt_head",
        "iopt_honly", "iopt_csel", "resultobj", "element", "variable"
    ):
        try:
            state[attr] = getattr(comres, attr)
        except Exception:
            pass
    return state


def _restore_comres_state(comres, state):
    for attr, value in state.items():
        try:
            setattr(comres, attr, value)
        except Exception:
            pass


def _configure_comres(comres, elmres, series, csv_path):
    comres.pResult = elmres
    comres.f_name = csv_path
    comres.iopt_exp = 4      # exportar a CSV
    comres.iopt_sep = 3      # separador ';'
    comres.iopt_head = 1     # cabecera
    try:
        comres.iopt_honly = 0
    except Exception:
        pass
    try:
        comres.iopt_csel = 1
    except Exception:
        pass

    elements = []
    variables = []
    resultobj = []

    for item in series:
        elements.append(item["element"])
        variables.append(item["variable"])
        resultobj.append(item.get("resultobj", elmres))

    comres.element = elements
    comres.variable = variables
    comres.resultobj = resultobj


def _export_results_csv(app, sc, elmres, series, csv_path):
    comres, temp_comres = _get_or_create_comres(sc)
    if comres is None:
        raise RuntimeError("No se pudo obtener ni crear un ComRes.")

    state = _capture_comres_state(comres)

    try:
        if os.path.isfile(csv_path):
            try:
                os.remove(csv_path)
            except Exception:
                pass

        _configure_comres(comres, elmres, series, csv_path)
        ierr = comres.Execute()
        if ierr != 0:
            raise RuntimeError(f"ComRes devolvio ierr={ierr}")

        if not os.path.isfile(csv_path):
            raise RuntimeError("El CSV no fue creado por ComRes.")

        if os.path.getsize(csv_path) == 0:
            raise RuntimeError("El CSV creado esta vacio.")

        return csv_path

    finally:
        _restore_comres_state(comres, state)
        if temp_comres:
            try:
                comres.Delete()
            except Exception:
                pass


# =============================================================================
# LECTURA CSV Y EXTRACCION DE VALORES
# =============================================================================

def _read_csv_table(csv_path):
    # Detectar delimitador probando los más comunes que usa PowerFactory
    for delimiter in (";", ",", "\t"):
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f, delimiter=delimiter)
            rows = [row for row in reader if row]

        data_rows = []
        for row in rows:
            if not row:
                continue
            t = _float_or_none(row[0])
            if t is not None:
                data_rows.append(row)

        if data_rows:
            return data_rows

    # Ningun delimitador funciono: volcar primeras lineas para diagnostico
    try:
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            preview = [f.readline().rstrip("\n") for _ in range(6)]
    except Exception:
        preview = ["<no se pudo leer>"]
    preview_str = " | ".join(repr(l) for l in preview if l)
    raise RuntimeError(
        "No se encontraron filas de datos numericas en el CSV. "
        f"Primeras lineas del archivo: {preview_str}"
    )


def _column_values(data_rows, col_index):
    values = []
    for row in data_rows:
        if col_index < len(row):
            values.append(_float_or_none(row[col_index]))
        else:
            values.append(None)
    return values


def _sample_series_at_time(times, values, target_time, interpolate=True):
    pairs = [(t, v) for t, v in zip(times, values) if t is not None and v is not None]
    if not pairs:
        return None, None, "sin_datos"

    pairs.sort(key=lambda x: x[0])
    times_sorted = [p[0] for p in pairs]
    values_sorted = [p[1] for p in pairs]

    for t, v in zip(times_sorted, values_sorted):
        if abs(t - target_time) <= TIME_TOL_S:
            return v, t, "exacto"

    if len(times_sorted) == 1:
        return values_sorted[0], times_sorted[0], "unico_punto"

    if target_time <= times_sorted[0]:
        return values_sorted[0], times_sorted[0], "antes_del_primer_punto"

    if target_time >= times_sorted[-1]:
        return values_sorted[-1], times_sorted[-1], "despues_del_ultimo_punto"

    for i in range(len(times_sorted) - 1):
        t0 = times_sorted[i]
        t1 = times_sorted[i + 1]
        if t0 <= target_time <= t1:
            v0 = values_sorted[i]
            v1 = values_sorted[i + 1]
            if abs(t1 - t0) <= TIME_TOL_S:
                return v0, t0, "puntos_duplicados"

            if not interpolate:
                if abs(target_time - t0) <= abs(target_time - t1):
                    return v0, t0, "cercano_sin_interpolar"
                return v1, t1, "cercano_sin_interpolar"

            alpha = (target_time - t0) / (t1 - t0)
            value = v0 + alpha * (v1 - v0)
            return value, target_time, "interpolado"

    return None, None, "fuera_de_rango"


def _extract_power_at_target(csv_path, series, target_time=TARGET_TIME_S):
    data_rows = _read_csv_table(csv_path)

    times = [_float_or_none(row[0]) for row in data_rows]

    results = []
    idx = 1  # columna 0 = tiempo
    for item in series[1:]:
        col_values = _column_values(data_rows, idx)
        value, used_time, mode = _sample_series_at_time(
            times=times,
            values=col_values,
            target_time=target_time,
            interpolate=INTERPOLATE_IF_NEEDED,
        )

        results.append({
            "gen_key": item["gen_key"],
            "gen_name": item["gen_name"],
            "variable": item["variable"],
            "candidate_index": item["candidate_index"],
            "candidate_variable": item["candidate_variable"],
            "value": value,
            "used_time": used_time,
            "mode": mode,
        })
        idx += 1

    return results


def _pick_best_candidate_per_generator(results, generators):
    """
    Para cada generador, devuelve el primer candidato válido según el orden
    de POWER_VARIABLE_CANDIDATES.
    """
    grouped = {}
    for item in results:
        grouped.setdefault(item["gen_key"], []).append(item)

    picked = []
    for gen in generators:
        key = _obj_key(gen)
        candidates = grouped.get(key, [])
        chosen = None

        for cand in sorted(candidates, key=lambda x: x["candidate_index"]):
            if cand["value"] is not None:
                chosen = cand
                break

        if chosen is None and candidates:
            chosen = sorted(candidates, key=lambda x: x["candidate_index"])[0]

        picked.append((gen, chosen))

    return picked


# =============================================================================
# SALIDA
# =============================================================================

def _write_output_csv(output_path, rows):
    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
            "loc_name",
            "class_pf",
            "outserv",
            "variable_usada",
            "time_used_s",
            "modo",
            "P_MW",
        ])
        for row in rows:
            writer.writerow([
                row["loc_name"],
                row["class_pf"],
                row["outserv"],
                row["variable_usada"],
                "" if row["time_used_s"] is None else f"{row['time_used_s']:.6f}",
                row["modo"],
                "" if row["P_MW"] is None else f"{row['P_MW']:.6f}",
            ])


def _print_summary(app, rows, output_path, target_time):
    app.PrintInfo("")
    app.PrintInfo("=" * 78)
    app.PrintInfo("EXTRACCION RMS - POTENCIA DE GENERADORES EN t = 5 s")
    app.PrintInfo("=" * 78)
    app.PrintInfo(f"Salida      : {output_path}")
    app.PrintInfo(f"Objetivo    : {target_time:.6f} s")
    app.PrintInfo(f"Generadores : {len(rows)}")
    app.PrintInfo("")

    ok = 0
    for row in rows:
        if row["P_MW"] is not None:
            ok += 1
        p_txt = "None" if row["P_MW"] is None else f"{row['P_MW']:+.6f} MW"
        t_txt = "None" if row["time_used_s"] is None else f"{row['time_used_s']:.6f} s"
        app.PrintInfo(
            f"- {row['loc_name']:<30} [{row['class_pf']:<12}] "
            f"P={p_txt:<18}  t_usado={t_txt:<12}  var={row['variable_usada']:<14}  modo={row['modo']}"
        )

    app.PrintInfo("")
    app.PrintInfo(f"Valores validos : {ok}/{len(rows)}")
    app.PrintInfo("=" * 78)


# =============================================================================
# MAIN
# =============================================================================

def main():
    app = _get_application()

    try:
        app.ClearOutputWindow()
    except Exception:
        pass

    sc = app.GetActiveStudyCase()
    if sc is None:
        raise RuntimeError("No hay caso de estudio activo.")

    project = app.GetActiveProject()
    if project is None:
        raise RuntimeError("No hay proyecto activo.")

    elmres = _get_active_result_file(sc)
    if elmres is None:
        raise RuntimeError("No se encontro ningun ElmRes en el caso de estudio activo.")

    generators = _get_generators(app)
    if not generators:
        raise RuntimeError("No se encontraron generadores calculables en el proyecto activo.")

    app.PrintInfo(f"Proyecto   : {_get_loc_name(project)}")
    app.PrintInfo(f"Caso       : {_get_loc_name(sc)}")
    app.PrintInfo(f"ElmRes     : {_get_loc_name(elmres)}")
    app.PrintInfo(f"Generadores: {len(generators)}")
    app.PrintInfo(f"t objetivo : {TARGET_TIME_S:.6f} s")
    app.PrintInfo("")

    output_dir = os.path.join(_script_dir(), "salidas_rms_5s")
    _ensure_dir(output_dir)
    csv_path = os.path.join(output_dir, OUTPUT_BASENAME)

    series = _build_export_series(generators, elmres)
    exported_csv = _export_results_csv(app, sc, elmres, series, csv_path)
    results = _extract_power_at_target(exported_csv, series, TARGET_TIME_S)
    picked = _pick_best_candidate_per_generator(results, generators)

    rows = []
    for gen, chosen in picked:
        outserv = _safe_float(_get_attr_safe(gen, "outserv"), 0)
        outserv_txt = "fuera_servicio" if int(outserv or 0) else "en_servicio"

        row = {
            "loc_name": _get_loc_name(gen),
            "class_pf": _get_class_name(gen),
            "outserv": outserv_txt,
            "variable_usada": "" if chosen is None else chosen["variable"],
            "time_used_s": None if chosen is None else chosen["used_time"],
            "modo": "sin_datos" if chosen is None else chosen["mode"],
            "P_MW": None if chosen is None else chosen["value"],
        }
        rows.append(row)

    summary_path = os.path.join(output_dir, "resumen_potencias_generadores_t_5s.csv")
    _write_output_csv(summary_path, rows)
    _print_summary(app, rows, summary_path, TARGET_TIME_S)

    app.PrintInfo("")
    app.PrintInfo(f"CSV crudo   : {exported_csv}")
    app.PrintInfo(f"CSV resumen : {summary_path}")


main()
