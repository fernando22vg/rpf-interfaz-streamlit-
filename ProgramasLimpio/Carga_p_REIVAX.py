#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""InjectorModelo_General.py
--------------------------------
Inyector masivo (proceso inverso) para parámetros de modelos en PowerFactory.

Objetivo
- Permitir modificar en lote parámetros de todo el modelo (no solo Reivax).
- Basado en un archivo de entrada (Excel/CSV) con una tabla "larga".

Entrada (tabla larga)
Puede ser Excel (.xlsx) o CSV (.csv). Debe incluir como mínimo las columnas:

  - unidad_loc_name   : loc_name exacto del ElmSym (o un identificador de unidad)
  - destino_type      : uno de {"ElmDsl","BlkRef"}
  - destino_id        : identificador del destino dentro del composite/model
                         * Para ElmDsl: se compara contra el loc_name del ElmDsl destino.
                         * Para BlkRef: se compara contra el loc_name del BlkRef destino.
  - parametro         : nombre del atributo a setear
  - valor             : valor numérico o texto

Columnas opcionales:
  - dsl_nombre        : prefijo/slot name (si aplica) para debug
  - blkdef_nombre     : debug

Restricciones / reglas
- No intenta cargar arrays/tablas complejas: si el valor parece un array/tablas
  (por ejemplo "[ ... ]" o contiene ';' repetido), se registra en log y se omite.

- Para ElmDsl:
    * se valida atributo con GetAttributeNames() si está disponible.

- Para BlkRef:
    * se valida escribiendo solo si GetAttributeNames() existe, si no se intenta SetAttribute.

Modo UI
- El script pregunta al usuario si desea aplicar:
    (1) una sola unidad
    (2) todas las unidades del archivo
    (3) un subconjunto (lista)

Salida
- Genera Excel de log con resultados por fila.

Uso en PowerFactory
- "Run as Script" y seleccionar el archivo.

NOTA
- Este inyector requiere que el archivo de entrada sea coherente con los loc_name
  en PF. Para modelos Reivax específicos puedes reutilizar la misma idea usando
  el extractor para preparar destino_id.
"""

import os
import sys
import re
import json
from datetime import datetime

import pandas as pd

import powerfactory as pf


# -------------------------- Configuración --------------------------

DESTINOS_PERMITIDOS = {"ElmDsl", "BlkRef"}

# Si el valor es muy "estructurado", probablemente es array/tabla -> omitir.
_ARRAY_LIKE_RE = re.compile(r"\[.*\]|\{.*\}|\(.*\)|\\n|;|,.*,.+")


def _log_path(carpeta, prefix):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(carpeta, f"{prefix}_{ts}.xlsx")


def _sanitize_str(x):
    if x is None:
        return ""
    return str(x).strip()


def _looks_array_or_table(valor_str: str) -> bool:
    if valor_str is None:
        return False
    s = str(valor_str).strip()
    if not s:
        return False
    # Casos típicos (ajusta según tu formato)
    if _ARRAY_LIKE_RE.search(s):
        return True
    # Heurística adicional: si contiene separadores múltiples
    if s.count(",") >= 2:
        return True
    return False


def _to_typed_value(v):
    """Convierte valor a tipo apto para SetAttribute.

    - int si es entero
    - float si es numérico
    - bool si detecta true/false/0/1
    - str en otro caso
    """
    if v is None:
        return None

    if isinstance(v, (int, float)) and not isinstance(v, bool):
        # numpy types llegan a veces como float/int
        return v

    s = str(v).strip()
    if s == "":
        return ""

    low = s.lower()
    if low in {"true", "false"}:
        return low == "true"
    if low in {"0", "1"}:
        # Ojo: 0/1 podrían ser numéricos de parámetros. Lo tratamos como bool SOLO si
        # el destino_type sugiere un booleano, pero aquí no lo sabemos. Retornamos int.
        # Para evitar errores, convertimos como int.
        try:
            return int(s)
        except Exception:
            return s

    # Decimal con coma
    s2 = s.replace(",", ".") if s.count(",") == 1 and s.count(".") == 0 else s

    try:
        f = float(s2)
        if abs(f - int(round(f))) < 1e-12:
            return int(round(f))
        return f
    except Exception:
        return s


def _get_app_project():
    app = pf.GetApplication()
    if app is None:
        raise RuntimeError("No se pudo obtener PowerFactory application (pf.GetApplication() == None).")

    project = app.GetActiveProject()
    if project is None:
        raise RuntimeError("No hay proyecto activo en PowerFactory.")

    return app, project


def _elegir_carpeta():
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    carpeta = filedialog.askdirectory(
        title="Carpeta de salida — InjectorModelo_General",
        mustexist=True,
    )
    root.destroy()
    if not carpeta:
        raise RuntimeError("Selección cancelada.")
    return carpeta.replace("/", os.sep)


def _preguntar_unidades(unidades_existentes):
    unidades_existentes = sorted(unidades_existentes)
    print("\nUnidades disponibles en el archivo:")
    for u in unidades_existentes:
        print(f"  - {u}")

    print("\nSeleccion de modo de aplicacion:")
    print("  [1] Aplicar a UNA unidad")
    print("  [2] Aplicar a TODAS las unidades")
    print("  [3] Aplicar a VARIAS unidades (por lista) ")

    while True:
        opt = input("Opcion [1/2/3]: ").strip()
        if opt in {"1", "2", "3"}:
            break
        print("Opcion invalida")

    if opt == "2":
        return unidades_existentes

    if opt == "1":
        while True:
            u = input("loc_name unidad (unidad_loc_name) exacto: ").strip()
            if u in unidades_existentes:
                return [u]
            print("Unidad no encontrada en el archivo. Intente de nuevo.")

    # opt == 3
    while True:
        raw = input("Lista separada por coma (ej: sym_A,sym_B): ").strip()
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if not parts:
            print("Lista vacia")
            continue
        bad = [p for p in parts if p not in unidades_existentes]
        if bad:
            print("Unidades no encontradas:", bad)
            continue
        return parts


def _buscar_obj_unidad(app, unidad_loc_name: str):
    """Busca ElmSym (o genéricamente) por loc_name."""
    # Buscar por clase ElmSym primero; si no, se intenta ElmGenstat.
    for cls in ("*.ElmSym", "*.ElmGenstat"):
        objs = app.GetCalcRelevantObjects(cls) or []
        for o in objs:
            try:
                if str(o.loc_name).strip() == unidad_loc_name:
                    return o
            except Exception:
                pass
    return None


def _pelm_lista(cm):
    lista = None
    try:
        lista = cm.GetAttribute("pelm")
    except Exception:
        lista = None
    if isinstance(lista, (list, tuple)) and len(lista) > 0:
        return [x for x in lista if x is not None]

    res = []
    for i in range(60):
        try:
            ref = cm.GetAttribute(f"pelm[{i}]")
        except Exception:
            ref = None
        if ref is None:
            break
        res.append(ref)
    return res


def _cls(obj):
    try:
        return obj.GetClassName()
    except Exception:
        return ""


def _fullname(obj):
    try:
        return obj.GetFullName()
    except Exception:
        return None


def _buscar_composite_model(sym_elm, all_comps):
    # 1) c_pmod directo
    try:
        cm = sym_elm.GetAttribute("c_pmod")
        if cm is not None and _cls(cm) == "ElmComp":
            return cm, "c_pmod"
    except Exception:
        pass

    # 2) escaneo por pelm
    fn = _fullname(sym_elm)
    if fn is not None:
        for cm in all_comps:
            for ref in _pelm_lista(cm):
                if _fullname(ref) == fn:
                    return cm, "pelm_scan"

    # 3) carpeta padre fallback
    try:
        parent = sym_elm.GetParent()
        if parent is not None:
            comps = parent.GetContents("*.ElmComp", 0) or []
            if comps:
                return comps[0], "carpeta_padre"
    except Exception:
        pass

    return None, None


def _dsl_slots_de_composite(cm):
    return [ref for ref in _pelm_lista(cm) if _cls(ref) == "ElmDsl"]


def _blkrefs_in_composite(cm):
    """Recupera BlkRef dentro de los BlkDef instanciados.

    Estrategia: recorrer contenidos del composite y filtrar BlkRef.
    Si no hay acceso directo, se intenta escanear hijos.
    """
    blks = []
    try:
        # Algunos modelos permiten GetContents desde el composite.
        cont = cm.GetContents("*.BlkRef", 1) or []
        blks.extend(cont)
    except Exception:
        pass

    # Fallback: recorrer pelm -> ElmDsl -> typ_id -> GetContents('*')
    if not blks:
        for dsl in _dsl_slots_de_composite(cm):
            try:
                blkdef = dsl.GetAttribute("typ_id")
            except Exception:
                blkdef = None
            if blkdef is None:
                continue
            try:
                hijos = blkdef.GetContents("*.BlkRef", 1) or []
                blks.extend(hijos)
            except Exception:
                pass

    # Deduplicar por FullName si es posible
    uniq = {}
    for b in blks:
        k = _fullname(b) or id(b)
        uniq[k] = b
    return list(uniq.values())


def _get_attribute_names(obj):
    try:
        names = obj.GetAttributeNames()
        if isinstance(names, (list, tuple)):
            return set([str(x) for x in names])
        return set()
    except Exception:
        return set()


def _set_attribute(obj, attr, value):
    # Intentar asignacion directa y luego SetAttribute
    try:
        setattr(obj, attr, value)
        return True, "setattr"
    except Exception:
        pass

    try:
        obj.SetAttribute(attr, value)
        return True, "SetAttribute"
    except Exception as e:
        return False, f"SetAttribute failed: {e}"


def main():
    app, project = _get_app_project()
    app.ClearOutputWindow()

    print("\nInjectorModelo_General — inicio")
    print(f"Proyecto: {project.loc_name}")

    # Selección archivo de entrada
    # En PowerFactory podemos usar input() con ruta.
    inp = input("Ruta Excel/CSV de entrada (xlsx o csv): ").strip().strip('"')
    if not os.path.isfile(inp):
        raise RuntimeError(f"Archivo no encontrado: {inp}")

    # Selección carpeta salida
    out_dir = _elegir_carpeta()

    # Cargar tabla
    ext = os.path.splitext(inp)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(inp)
    else:
        df = pd.read_excel(inp)

    # Validar columnas requeridas
    required = {"unidad_loc_name", "destino_type", "destino_id", "parametro", "valor"}
    missing = required - set(df.columns)
    if missing:
        raise RuntimeError(f"Faltan columnas requeridas en el input: {sorted(missing)}\nColumnas presentes: {list(df.columns)}")

    # Limpieza base
    df["unidad_loc_name"] = df["unidad_loc_name"].astype(str).str.strip()
    df["destino_type"] = df["destino_type"].astype(str).str.strip()
    df["destino_id"] = df["destino_id"].astype(str).str.strip()
    df["parametro"] = df["parametro"].astype(str).str.strip()
    df["valor"] = df["valor"]

    # Filtrar destinos permitidos
    invalid_type = set(df[~df["destino_type"].isin(DESTINOS_PERMITIDOS)]["destino_type"].unique())
    if invalid_type:
        print("[WARN] destiny_type(s) no permitidos. Se omiten:", invalid_type)
        df = df[df["destino_type"].isin(DESTINOS_PERMITIDOS)]

    unidades_presentes = set(df["unidad_loc_name"].unique())
    unidades_aplicar = _preguntar_unidades(unidades_presentes)

    # Buscar composite models para performance
    all_comps = []
    try:
        all_comps = app.GetCalcRelevantObjects("*.ElmComp") or []
    except Exception:
        all_comps = []

    # Preparar log
    resultados = []

    # Cache de unidades->objetos composite
    cache_unidad = {}

    for loc_unit in unidades_aplicar:
        unidad_obj = _buscar_obj_unidad(app, loc_unit)
        if unidad_obj is None:
            # registrar todo como error
            sub = df[df["unidad_loc_name"] == loc_unit]
            for _, row in sub.iterrows():
                resultados.append({
                    "unidad_loc_name": loc_unit,
                    "destino_type": row["destino_type"],
                    "destino_id": row["destino_id"],
                    "parametro": row["parametro"],
                    "valor_input": row["valor"],
                    "set_ok": False,
                    "error": "unidad_obj no encontrada en PF",
                })
            continue

        cm, metodo = _buscar_composite_model(unidad_obj, all_comps)
        cache_unidad[loc_unit] = (unidad_obj, cm, metodo)

    # Aplicar fila por fila
    for i, row in df.iterrows():
        loc_unit = row["unidad_loc_name"]
        if loc_unit not in unidades_aplicar:
            continue

        destino_type = row["destino_type"]
        destino_id = row["destino_id"]
        parametro = row["parametro"]
        valor_raw = row["valor"]

        # Omite si parece array/tabla
        if _looks_array_or_table(valor_raw if isinstance(valor_raw, str) else str(valor_raw)):
            resultados.append({
                "unidad_loc_name": loc_unit,
                "destino_type": destino_type,
                "destino_id": destino_id,
                "parametro": parametro,
                "valor_input": valor_raw,
                "set_ok": False,
                "error": "omitido: valor parece array/tabla",
            })
            continue

        unidad_obj, cm, metodo = cache_unidad.get(loc_unit, (None, None, None))
        if cm is None:
            resultados.append({
                "unidad_loc_name": loc_unit,
                "destino_type": destino_type,
                "destino_id": destino_id,
                "parametro": parametro,
                "valor_input": valor_raw,
                "set_ok": False,
                "error": "sin ElmComp (composite model no encontrado)",
                "metodo": metodo,
            })
            continue

        # localizar destino dentro del composite
        dest_obj = None

        if destino_type == "ElmDsl":
            dsl_list = _dsl_slots_de_composite(cm)
            for dsl in dsl_list:
                try:
                    if str(dsl.loc_name).strip() == destino_id:
                        dest_obj = dsl
                        break
                except Exception:
                    pass

        elif destino_type == "BlkRef":
            blk_list = _blkrefs_in_composite(cm)
            for b in blk_list:
                try:
                    if str(b.loc_name).strip() == destino_id:
                        dest_obj = b
                        break
                except Exception:
                    pass

        if dest_obj is None:
            resultados.append({
                "unidad_loc_name": loc_unit,
                "destino_type": destino_type,
                "destino_id": destino_id,
                "parametro": parametro,
                "valor_input": valor_raw,
                "set_ok": False,
                "error": f"destino no encontrado ({destino_type})",
                "metodo": metodo,
            })
            continue

        # Validar atributo con GetAttributeNames si se puede
        attr_names = _get_attribute_names(dest_obj)
        if attr_names:
            if parametro not in attr_names:
                resultados.append({
                    "unidad_loc_name": loc_unit,
                    "destino_type": destino_type,
                    "destino_id": destino_id,
                    "parametro": parametro,
                    "valor_input": valor_raw,
                    "set_ok": False,
                    "error": "parametro no existe en GetAttributeNames()",
                })
                continue

        typed_val = _to_typed_value(valor_raw)

        ok, details = _set_attribute(dest_obj, parametro, typed_val)
        resultados.append({
            "unidad_loc_name": loc_unit,
            "destino_type": destino_type,
            "destino_id": destino_id,
            "parametro": parametro,
            "valor_input": valor_raw,
            "valor_typed": typed_val,
            "set_ok": bool(ok),
            "details": details,
        })

    df_res = pd.DataFrame(resultados)
    out_xlsx = _log_path(out_dir, "InjectorModeloLog")
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df_res.to_excel(writer, index=False, sheet_name="log")
        # Summary
        total = len(df_res)
        ok_count = int(df_res["set_ok"].sum()) if "set_ok" in df_res.columns else 0
        summ = pd.DataFrame([
            {"total_filas": total, "set_ok": ok_count, "set_fail": total - ok_count},
        ])
        summ.to_excel(writer, index=False, sheet_name="summary")

    print("\nInjectorModelo_General — FIN")
    print(f"Log generado: {out_xlsx}")


if __name__ == "__main__":
    main()

