#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DatosCurvas_v3.py
-----------------
Exporta los datos de las gráficas usando el equivalente a:
  Gráfico -> Export -> Show Data

Además incluye un MODO_DIAGNOSTICO para descubrir cómo está armada cada
página/plot en PowerFactory, especialmente cuando una página no expone
las curvas como objetos GrpCurve legibles.
"""

import os
from datetime import datetime
import powerfactory as pf

# ══════════════════════════════════════════════════════════════
RAIZ = r"C:\Datos del CNDC\01_INFO CNDC_RPF"

# Incluir también la columna de tiempo de simulación en cada exportación.
INCLUIR_TIEMPO_SIMULACION = True
VARIABLE_TIEMPO_SIMULACION = 'b:tnow'

# Modo productivo por defecto. Cambiar a True solo para depurar estructura.
MODO_DIAGNOSTICO = False

# Si la lista está vacía, se diagnostican TODAS las páginas encontradas.
PAGINAS_MODO_DIAGNOSTICO = []

# Si la lista está vacía, exporta todas las páginas encontradas.
PAGINAS_A_EXPORTAR = [
    'F. Barras SIN',
    'F.P. ZON01',
    'F.P. TIQ01',
    'F.P. BOT01',
    'F.P. BOT02',
    'F.P. BOT03',
    'F.P. BOT.ALL',
    'F.P. CUT01',
    'F.P. CUT02',
    'F.P. CUT03',
    'F.P. CUT04',
    'F.P. CUT05',
    'F.P. CUT.ALL',
    'F.P. SRO.ALL',
    'F.P. SRO02',
    'F.P. SRO01',
    'F.P. SAI01',
    'F.P. CHU.ALL',
    'F.P. CHU01',
    'F.P. CHU02',
    'F.P. CAH01',
    'F.P. CAH02',
    'F.P. CAH.ALL',
    'F.P. HAR.ALL',
    'F.P. HAR01',
    'F.P. HAR02',
    'F.P. HUA.ALL',
    'F.P. HUA01',
    'F.P. HUA02',
    'F.P. ANG03',
    'F.P. CRB',
    'F.slack',
    'Velocidades',
    'Ángulos',
]
# ══════════════════════════════════════════════════════════════


# ─────────────────────────────────────────────────────────────
# Selección interactiva (tkinter – compatible con PowerFactory)
# ─────────────────────────────────────────────────────────────
def elegir(opciones, titulo):
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

    tk.Button(root, text="Aceptar", command=confirmar, width=14).pack(pady=(6, 12))
    root.bind('<Return>', confirmar)
    listbox.bind('<Double-Button-1>', confirmar)

    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    root.mainloop()

    if resultado[0] is None:
        raise RuntimeError(f"Selección cancelada para: {titulo}")
    return resultado[0]


def _seleccionar_carpeta_salida():
    import re

    semestres = sorted(d for d in os.listdir(RAIZ)
                       if os.path.isdir(os.path.join(RAIZ, d)))
    if not semestres:
        raise RuntimeError(f"No se encontraron semestres en: {RAIZ}")
    semestre = elegir(semestres, "Semestre de estudio")

    base_ev = os.path.join(RAIZ, semestre, "Análisis_todos_los_eventos")
    if not os.path.isdir(base_ev):
        raise RuntimeError(f"No existe la carpeta de eventos: {base_ev}")
    eventos = sorted(d for d in os.listdir(base_ev)
                     if os.path.isdir(os.path.join(base_ev, d)))
    if not eventos:
        raise RuntimeError(f"No se encontraron eventos en: {base_ev}")
    evento = elegir(eventos, "Evento simulado")

    # Extraer número de evento para construir las opciones de simulación
    m = re.search(r"(\d+)", evento)
    n_ev = m.group(1) if m else "0"
    sim_opciones = [f"E{n_ev}.0", f"E{n_ev}.1"]
    simulacion = elegir(sim_opciones, "Tipo de simulación")

    carpeta_salida = os.path.join(base_ev, evento, simulacion, "Datos Curvas")
    os.makedirs(carpeta_salida, exist_ok=True)

    return carpeta_salida, semestre, evento, simulacion


# ─────────────────────────────────────────────────────────────
# Utilidades generales
# ─────────────────────────────────────────────────────────────
def _sanear_nombre_archivo(txt):
    txt = str(txt or '').strip()
    for ch in r'\\/:*?"<>|':
        txt = txt.replace(ch, '_')
    return txt.strip().strip('.') or 'sin_nombre'


def _get_attr_safe(obj, attr, default=None):
    try:
        return obj.GetAttribute(attr)
    except Exception:
        try:
            return getattr(obj, attr)
        except Exception:
            return default


def _get_loc_name(obj, default=''):
    if obj is None:
        return default
    try:
        return obj.GetAttribute('loc_name')
    except Exception:
        try:
            return obj.loc_name
        except Exception:
            return default or str(obj)


def _get_class_name(obj, default=''):
    if obj is None:
        return default
    try:
        return obj.GetClassName()
    except Exception:
        return default


def _obj_key(obj):
    if obj is None:
        return 'None'

    class_name = _get_class_name(obj, '')
    loc_name = _get_loc_name(obj, '')

    # Primero intentar una ruta/nombre completo, que suele ser la mejor clave.
    for getter in ('GetFullName', 'GetFullPath'):
        try:
            val = getattr(obj, getter)()
            if val:
                return f"{class_name}::{val}"
        except Exception:
            pass

    # Luego class + loc_name. Para este caso es mejor que usar IDs genéricos,
    # porque algunos wrappers de PF devuelven identificadores poco útiles o
    # repetidos para elementos distintos dentro de curveTableElement.
    if loc_name or class_name:
        return f"{class_name}::{loc_name}"

    for attr in ('fold_id', 'ID', 'id'):
        val = _get_attr_safe(obj, attr)
        if val not in (None, ''):
            return f"{class_name}::{val}"

    try:
        return str(obj)
    except Exception:
        return f"{class_name}::{loc_name}"


def _normalizar_a_lista(val):
    if val is None:
        return []
    if isinstance(val, list):
        return val
    if isinstance(val, tuple):
        return list(val)
    if isinstance(val, str):
        return [val]

    try:
        return list(val)
    except Exception:
        pass

    return [val]


def _resolver_variable_str(var_obj):
    """
    Convierte la referencia de variable al texto que espera ComRes.
    Ejemplos: 'm:Psum:bus1', 'n:fehz:bus1', 'b:tnow'
    """
    if var_obj is None:
        return None

    if isinstance(var_obj, str):
        txt = var_obj.strip()
        return txt or None

    candidatos = [
        'loc_name', 'f_name', 'u_name', 'for_name',
        'name', 'var_name', 'variable', 'parameter'
    ]
    for attr in candidatos:
        val = _get_attr_safe(var_obj, attr)
        if isinstance(val, str) and val.strip():
            return val.strip()

    try:
        txt = str(var_obj).strip()
        return txt or None
    except Exception:
        return None


def _mostrar_pagina(app, page):
    try:
        desktop = app.GetGraphicsBoard()
        if desktop is not None:
            try:
                desktop.Show(page)
                return True
            except Exception:
                pass
    except Exception:
        pass

    try:
        app.Show(page)
        return True
    except Exception:
        return False


def _dedupe_objetos(objetos):
    unicos = {}
    for obj in objetos or []:
        unicos[_obj_key(obj)] = obj
    return list(unicos.values())


def _fmt_obj(obj):
    if obj is None:
        return 'None'
    return f"[{_get_class_name(obj, '?')}] {_get_loc_name(obj, str(obj))}"


# ─────────────────────────────────────────────────────────────
# Descubrimiento de páginas y resultados
# ─────────────────────────────────────────────────────────────
def _buscar_paginas(app, sc):
    encontradas = {}

    try:
        for page in (sc.GetContents('*.GrpPage', 1) or []):
            encontradas[_obj_key(page)] = page
    except Exception:
        pass

    try:
        desktop = app.GetGraphicsBoard()
        if desktop is not None:
            for page in (desktop.GetContents('*.GrpPage', 1) or []):
                encontradas[_obj_key(page)] = page
    except Exception:
        pass

    paginas = list(encontradas.values())
    nombres_en_pf = {_get_loc_name(p) for p in paginas}

    if PAGINAS_A_EXPORTAR:
        orden = {nom: i for i, nom in enumerate(PAGINAS_A_EXPORTAR)}
        paginas = [p for p in paginas if _get_loc_name(p) in orden]
        paginas.sort(key=lambda p: orden.get(_get_loc_name(p), 999999))

        no_encontradas = [nom for nom in PAGINAS_A_EXPORTAR if nom not in nombres_en_pf]
        if no_encontradas:
            app.PrintWarn(
                f"ADVERTENCIA: {len(no_encontradas)} página(s) de PAGINAS_A_EXPORTAR "
                f"no existen en el proyecto activo:"
            )
            for nom in no_encontradas:
                app.PrintWarn(f"  - '{nom}'")
            app.PrintWarn(
                "  Verifique que los nombres coincidan exactamente con los de PowerFactory."
            )
    else:
        paginas.sort(key=lambda p: _get_loc_name(p).lower())
        no_encontradas = []

    return paginas, no_encontradas


def _obtener_plots_de_page(page):
    plots = []
    patrones = (
        '*.GrpPlot',
        '*.PltLinebarplot',
        '*.PltPlot',
        '*.VisPlot',
    )

    for patron in patrones:
        try:
            plots.extend(page.GetContents(patron, 1) or [])
        except Exception:
            pass

    return _dedupe_objetos(plots)


def _get_dataseries_de_plot(plot):
    try:
        ds = plot.GetDataSeries()
        if ds is not None:
            return ds
    except Exception:
        pass

    try:
        dss = plot.GetContents('*.PltDataseries', 1) or []
        if dss:
            return dss[0]
    except Exception:
        pass

    return None


def _get_elmres_de_page(page, sc=None, series=None):
    attrs_to_try = ('pResult', 'Results', 'e_pResult', 'result')

    for attr in attrs_to_try:
        val = _get_attr_safe(page, attr)
        if val is not None:
            return val

    for plot in _obtener_plots_de_page(page):
        for attr in attrs_to_try:
            val = _get_attr_safe(plot, attr)
            if val is not None:
                return val

        ds = _get_dataseries_de_plot(plot)
        if ds is not None:
            for attr in ('curveTableResultFile', 'pResult', 'Results', 'result'):
                val = _get_attr_safe(ds, attr)
                lista = _normalizar_a_lista(val)
                if lista:
                    for item in lista:
                        if item is not None:
                            return item
                elif val is not None:
                    return val

    if series:
        for item in series:
            res = item.get('resultobj')
            if res is not None:
                return res

    if sc is not None:
        for elmres in (sc.GetContents('*.ElmRes', 1) or []):
            if _get_loc_name(elmres) == 'All calculations':
                return elmres

        elmres_list = sc.GetContents('*.ElmRes', 1) or []
        if elmres_list:
            return elmres_list[0]

    return None


def _obtener_o_crear_comres(sc):
    # Siempre crear un ComRes temporal nuevo para garantizar configuración limpia.
    # Reutilizar el existente puede heredar iopt_honly=1 u otros estados previos.
    try:
        comres = sc.CreateObject('ComRes', '__tmp_export_show_data__')
        if comres is not None:
            return comres, True
    except Exception:
        pass

    # Fallback: usar el existente si no se pudo crear uno nuevo.
    comres_list = sc.GetContents('*.ComRes', 1) or []
    if comres_list:
        return comres_list[0], False

    return None, False


# ─────────────────────────────────────────────────────────────
# Extracción de curvas / dataseries
# ─────────────────────────────────────────────────────────────
def _agregar_serie(curvas, vistos, obj, var_str, res=None):
    if obj is None or not var_str:
        return False

    key = (_obj_key(obj), var_str, _obj_key(res))
    if key in vistos:
        return False
    vistos.add(key)

    curvas.append({
        'element': obj,
        'variable': var_str,
        'resultobj': res,
        'element_name': _get_loc_name(obj),
        'element_class': _get_class_name(obj),
    })
    return True


def _extraer_series_desde_objetos(page):
    curvas = []
    vistos = set()

    attrs_obj = (
        'p_obj', 'p_object', 'e_obj', 'e_object',
        'obj', 'object', 'element', 'p_data_obj', 'p_data_element'
    )
    attrs_var = (
        'p_var', 'p_variable', 'e_var', 'e_variable',
        'var', 'variable', 'parameter', 'param',
        'p_data_var', 'p_data_attribute'
    )
    attrs_res = ('pResult', 'Results', 'e_pResult', 'result')

    objetos_curva = []
    for patron in ('*.GrpCurve', '*.PltDataseries'):
        try:
            objetos_curva.extend(page.GetContents(patron, 1) or [])
        except Exception:
            pass

    for curve in _dedupe_objetos(objetos_curva):
        obj = None
        var_obj = None
        res = None

        for attr in attrs_obj:
            obj = _get_attr_safe(curve, attr)
            if obj is not None:
                break

        for attr in attrs_var:
            var_obj = _get_attr_safe(curve, attr)
            if var_obj is not None:
                break

        for attr in attrs_res:
            res = _get_attr_safe(curve, attr)
            if res is not None:
                break

        _agregar_serie(curvas, vistos, obj, _resolver_variable_str(var_obj), res)

    return curvas


def _extraer_series_desde_dataseries_table(plot):
    curvas = []
    vistos = set()

    ds = _get_dataseries_de_plot(plot)
    if ds is None:
        return curvas

    elementos = _normalizar_a_lista(_get_attr_safe(ds, 'curveTableElement'))
    variables = _normalizar_a_lista(_get_attr_safe(ds, 'curveTableVariable'))
    resultados = _normalizar_a_lista(_get_attr_safe(ds, 'curveTableResultFile'))

    n = max(len(elementos), len(variables), len(resultados))
    if n == 0:
        return curvas

    for i in range(n):
        obj = elementos[i] if i < len(elementos) else None
        var_obj = variables[i] if i < len(variables) else None
        res = resultados[i] if i < len(resultados) else None

        # La primera fila suele venir vacía o incompleta en algunos plots.
        if obj is None and _resolver_variable_str(var_obj) is None and res is None:
            continue

        _agregar_serie(curvas, vistos, obj, _resolver_variable_str(var_obj), res)

    return curvas


def _extraer_series_de_page(page, app):
    """
    Extrae las curvas configuradas en la página para exportar sólo esas
    variables con ComRes, emulando "Export Show Data".

    Intenta dos caminos:
    1) objetos tipo GrpCurve / PltDataseries con atributos objeto-variable,
    2) tablas internas del DataSeries del plot (curveTableElement/Variable).
    """
    curvas = []
    vistos = set()

    for item in _extraer_series_desde_objetos(page):
        _agregar_serie(curvas, vistos, item['element'], item['variable'], item.get('resultobj'))

    for plot in _obtener_plots_de_page(page):
        for item in _extraer_series_desde_dataseries_table(plot):
            _agregar_serie(curvas, vistos, item['element'], item['variable'], item.get('resultobj'))

    if not curvas:
        app.PrintWarn('  Sin curvas/series legibles en la página.')

    return curvas


# ─────────────────────────────────────────────────────────────
# Diagnóstico
# ─────────────────────────────────────────────────────────────
def _print_attrs_existentes(app, obj, titulo, attrs):
    app.PrintInfo(f"  {titulo}:")
    hubo = False
    for attr in attrs:
        try:
            val = obj.GetAttribute(attr)
            if val is None:
                continue

            hubo = True
            if isinstance(val, (list, tuple)):
                app.PrintInfo(f"    - {attr}: lista(len={len(val)})")
            elif hasattr(val, 'GetClassName'):
                app.PrintInfo(f"    - {attr}: {_fmt_obj(val)}")
            else:
                txt = str(val)
                if len(txt) > 140:
                    txt = txt[:140] + '...'
                app.PrintInfo(f"    - {attr}: {txt}")
        except Exception:
            pass

    if not hubo:
        app.PrintInfo('    (sin atributos legibles)')


def _preview_tabla_dataseries(app, ds, max_filas=12):
    elementos = _normalizar_a_lista(_get_attr_safe(ds, 'curveTableElement'))
    variables = _normalizar_a_lista(_get_attr_safe(ds, 'curveTableVariable'))
    resultados = _normalizar_a_lista(_get_attr_safe(ds, 'curveTableResultFile'))

    n = max(len(elementos), len(variables), len(resultados))
    app.PrintInfo(f"    curveTable rows: {n}")

    if n == 0:
        app.PrintInfo('    curveTable vacía o no accesible')
        return

    keys = []
    for i in range(min(n, max_filas)):
        obj = elementos[i] if i < len(elementos) else None
        var = variables[i] if i < len(variables) else None
        res = resultados[i] if i < len(resultados) else None
        k = _obj_key(obj)
        keys.append(k)
        app.PrintInfo(
            f"      [{i:02d}] obj={_fmt_obj(obj)} | "
            f"var={_resolver_variable_str(var)} | res={_fmt_obj(res)} | key={k}"
        )

    if len(keys) > 1:
        app.PrintInfo(f"    keys únicas preview: {len(set(keys))} / {len(keys)}")

    if n > max_filas:
        app.PrintInfo(f"      ... {n - max_filas} filas más")


def _diagnosticar_plot(app, plot):
    app.PrintInfo('')
    app.PrintInfo(f"  Plot: {_fmt_obj(plot)}")

    _print_attrs_existentes(
        app,
        plot,
        'Atributos de resultado del plot',
        ('pResult', 'Results', 'e_pResult', 'result')
    )

    for patron in ('*.GrpCurve', '*.PltDataseries', '*.SetCrvfilt', '*'):
        try:
            items = plot.GetContents(patron, 1) or []
            app.PrintInfo(f"  Contenido {patron:16s}: {len(items)}")
        except Exception:
            pass

    ds = _get_dataseries_de_plot(plot)
    if ds is None:
        app.PrintInfo('  DataSeries: no accesible')
        return

    app.PrintInfo(f"  DataSeries: {_fmt_obj(ds)}")

    _print_attrs_existentes(
        app,
        ds,
        'Atributos de DataSeries',
        (
            'useIndividualResults',
            'curveTableElement',
            'curveTableVariable',
            'curveTableResultFile',
            'curveTableLabel',
            'curveTableColour',
            'curveTableAxis',
            'pResult',
            'Results',
            'result',
        )
    )

    _preview_tabla_dataseries(app, ds)

    series_plot = _extraer_series_desde_dataseries_table(plot)
    app.PrintInfo(f"  Series extraídas desde tabla: {len(series_plot)}")
    for i, item in enumerate(series_plot[:10], 1):
        app.PrintInfo(
            f"    {i:02d}. {_fmt_obj(item['element'])} | "
            f"{item['variable']} | res={_fmt_obj(item.get('resultobj'))}"
        )
    if len(series_plot) > 10:
        app.PrintInfo(f"    ... {len(series_plot) - 10} series más")


def _diagnosticar_page(app, sc, page):
    app.PrintInfo('')
    app.PrintInfo('=' * 78)
    app.PrintInfo(f"DIAGNÓSTICO DE PÁGINA: {_get_loc_name(page)}")
    app.PrintInfo('=' * 78)
    app.PrintInfo(f"Página: {_fmt_obj(page)}")

    visible = _mostrar_pagina(app, page)
    app.PrintInfo(f"Mostrar en GUI: {'OK' if visible else 'NO'}")

    _print_attrs_existentes(
        app,
        page,
        'Atributos de resultado de la página',
        ('pResult', 'Results', 'e_pResult', 'result')
    )

    for patron in ('*.GrpPlot', '*.PltLinebarplot', '*.PltDataseries', '*.GrpCurve', '*.SetCrvfilt', '*'):
        try:
            items = page.GetContents(patron, 1) or []
            app.PrintInfo(f"Contenido {patron:16s}: {len(items)}")
            for item in items[:10]:
                app.PrintInfo(f"  - {_fmt_obj(item)}")
            if len(items) > 10:
                app.PrintInfo(f"  ... {len(items) - 10} objetos más")
        except Exception:
            pass

    plots = _obtener_plots_de_page(page)
    app.PrintInfo('')
    app.PrintInfo(f"Plots detectados en la página: {len(plots)}")
    for plot in plots:
        _diagnosticar_plot(app, plot)

    series_obj = _extraer_series_desde_objetos(page)
    app.PrintInfo('')
    app.PrintInfo(f"Series extraídas desde objetos directos: {len(series_obj)}")
    for i, item in enumerate(series_obj[:10], 1):
        app.PrintInfo(
            f"  {i:02d}. {_fmt_obj(item['element'])} | "
            f"{item['variable']} | res={_fmt_obj(item.get('resultobj'))}"
        )
    if len(series_obj) > 10:
        app.PrintInfo(f"  ... {len(series_obj) - 10} series más")

    series_totales = _extraer_series_de_page(page, app)
    app.PrintInfo('')
    app.PrintInfo(f"TOTAL series detectadas por extractor combinado: {len(series_totales)}")
    for i, item in enumerate(series_totales[:15], 1):
        app.PrintInfo(
            f"  {i:02d}. {_fmt_obj(item['element'])} | "
            f"{item['variable']} | res={_fmt_obj(item.get('resultobj'))}"
        )
    if len(series_totales) > 15:
        app.PrintInfo(f"  ... {len(series_totales) - 15} series más")

    elmres = _get_elmres_de_page(page, sc=sc, series=series_totales)
    app.PrintInfo(f"ElmRes resuelto: {_fmt_obj(elmres)}")


# ─────────────────────────────────────────────────────────────
# Exportación tipo "Show Data"
# ─────────────────────────────────────────────────────────────
def _capturar_estado_comres(comres):
    estado = {}
    for attr in (
        'pResult', 'f_name', 'iopt_exp', 'iopt_sep', 'iopt_head',
        'iopt_honly', 'iopt_csel', 'resultobj', 'element', 'variable'
    ):
        try:
            estado[attr] = getattr(comres, attr)
        except Exception:
            pass
    return estado


def _restaurar_estado_comres(comres, estado):
    for attr, valor in estado.items():
        try:
            setattr(comres, attr, valor)
        except Exception:
            pass


def _configurar_comres_show_data(comres, elmres, series, salida_csv,
                                 usar_resultobj_explicit=False, iopt_exp_val=4):
    comres.pResult = elmres
    comres.f_name = salida_csv
    comres.iopt_exp = iopt_exp_val
    # iopt_sep: 0=space, 1=semicolon, 2=comma, 3=tab — use semicolon for CSV
    comres.iopt_sep = 1
    comres.iopt_head = 1

    for attr, val in (('iopt_honly', 0), ('iopt_csel', 1)):
        try:
            comres.SetAttribute(attr, val)
        except Exception:
            try:
                setattr(comres, attr, val)
            except Exception:
                pass

    elements = []
    variables = []
    resultobj = []

    # b:tnow vive en el ElmRes, no en un elemento de red → usar elmres, no None.
    if INCLUIR_TIEMPO_SIMULACION:
        elements.append(elmres)
        variables.append(VARIABLE_TIEMPO_SIMULACION)
        resultobj.append(elmres)

    for s in series:
        elements.append(s['element'])
        variables.append(s['variable'])
        resultobj.append(s['resultobj'] if s.get('resultobj') is not None else elmres)

    comres.element = elements
    comres.variable = variables
    # Sólo asignar resultobj cuando se usa modo explícito; evitar lista de None.
    if usar_resultobj_explicit:
        comres.resultobj = resultobj


def _es_xlsx_valido(ruta):
    """True si el archivo es un ZIP/XLSX real (magic bytes PK)."""
    try:
        with open(ruta, 'rb') as f:
            return f.read(4) == b'PK\x03\x04'
    except Exception:
        return False


def _detectar_separador(ruta):
    """Detecta el separador del CSV probando ';', ',', '\t'."""
    try:
        with open(ruta, encoding='utf-8-sig', errors='replace') as f:
            primera = f.readline()
        for sep in (';', ',', '\t'):
            if sep in primera:
                return sep
    except Exception:
        pass
    return ';'


def _parsear_celda(valor):
    """Convierte una cadena a int o float si es posible; si no, devuelve el texto."""
    txt = valor.strip()
    if not txt:
        return txt
    # PowerFactory usa coma decimal en algunas configuraciones regionales
    normalizado = txt.replace(',', '.')
    try:
        entero = int(normalizado)
        return entero
    except ValueError:
        pass
    try:
        return float(normalizado)
    except ValueError:
        return txt


def _texto_a_excel(txt_path, xlsx_path):
    """
    Convierte cualquier archivo de texto plano (CSV, TXT) exportado por
    ComRes a un Excel real (.xlsx). Detecta separador y convierte números.
    """
    separador = _detectar_separador(txt_path)
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        with open(txt_path, encoding='utf-8-sig', errors='replace') as f:
            primera = True
            for linea in f:
                celdas_raw = linea.rstrip('\r\n').split(separador)
                if primera:
                    # Primera fila: encabezados como texto
                    ws.append(celdas_raw)
                    primera = False
                else:
                    ws.append([_parsear_celda(c) for c in celdas_raw])
        wb.save(xlsx_path)
        return True
    except Exception:
        pass
    try:
        import csv as _csv
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        with open(txt_path, encoding='utf-8-sig', errors='replace', newline='') as f:
            primera = True
            for fila in _csv.reader(f, delimiter=separador):
                if primera:
                    ws.append(fila)
                    primera = False
                else:
                    ws.append([_parsear_celda(c) for c in fila])
        wb.save(xlsx_path)
        return True
    except Exception:
        return False


def _csv_a_excel(csv_path, xlsx_path):
    """Alias de compatibilidad — delega en _texto_a_excel."""
    return _texto_a_excel(csv_path, xlsx_path)


def _exportar_show_data(app, page, elmres, comres, series, salida_path):
    """
    Intenta exportar con varios valores de iopt_exp (XLSX nativo primero, luego
    CSV con conversión a Excel).  Devuelve (ierr, modo, ruta_final).
    """
    # Limpiar archivos previos
    for ext in ('.xlsx', '.csv', '.txt'):
        p = salida_path if salida_path.endswith(ext) else os.path.splitext(salida_path)[0] + ext
        try:
            if os.path.isfile(p):
                os.remove(p)
        except Exception:
            pass

    _mostrar_pagina(app, page)

    base_path = os.path.splitext(salida_path)[0]

    # (iopt_exp, extensión_archivo, descripción)
    # Orden: formatos XLSX nativos de PF 2019-2024, luego CSV como último recurso.
    formatos_excel = [(4, '.xlsx'), (3, '.xlsx'), (5, '.xlsx'), (0, '.xlsx')]
    formatos_csv   = [(6, '.csv'), (2, '.csv'), (1, '.csv')]

    ultimo_ierr = None

    for usar_resultobj_explicit in (False, True):
        modo_tag = 'resultobj' if usar_resultobj_explicit else 'base'

        # ── Intentos de Excel nativo ──────────────────────────────────────────
        for iopt_val, ext in formatos_excel:
            ruta = base_path + ext
            try:
                if os.path.isfile(ruta):
                    os.remove(ruta)
                _configurar_comres_show_data(
                    comres=comres,
                    elmres=elmres,
                    series=series,
                    salida_csv=ruta,
                    usar_resultobj_explicit=usar_resultobj_explicit,
                    iopt_exp_val=iopt_val,
                )
                ierr = comres.Execute()
                ultimo_ierr = ierr
                if ierr == 0 and os.path.isfile(ruta) and os.path.getsize(ruta) > 0:
                    if _es_xlsx_valido(ruta):
                        return 0, f"{modo_tag}+iopt{iopt_val}_xlsx", ruta
                    # PF escribió texto con extensión .xlsx → convertir
                    tmp_txt = base_path + '_tmp.txt'
                    try:
                        os.rename(ruta, tmp_txt)
                        if _texto_a_excel(tmp_txt, ruta):
                            os.remove(tmp_txt)
                            return 0, f"{modo_tag}+iopt{iopt_val}_txt→xlsx", ruta
                        os.rename(tmp_txt, ruta)
                    except Exception:
                        pass
            except Exception:
                pass

        # ── Intentos CSV + conversión a Excel ─────────────────────────────────
        for iopt_val, ext in formatos_csv:
            ruta_csv  = base_path + ext
            ruta_xlsx = base_path + '.xlsx'
            try:
                if os.path.isfile(ruta_csv):
                    os.remove(ruta_csv)
                _configurar_comres_show_data(
                    comres=comres,
                    elmres=elmres,
                    series=series,
                    salida_csv=ruta_csv,
                    usar_resultobj_explicit=usar_resultobj_explicit,
                    iopt_exp_val=iopt_val,
                )
                ierr = comres.Execute()
                ultimo_ierr = ierr
                if ierr == 0 and os.path.isfile(ruta_csv) and os.path.getsize(ruta_csv) > 0:
                    convertido = _csv_a_excel(ruta_csv, ruta_xlsx)
                    try:
                        os.remove(ruta_csv)
                    except Exception:
                        pass
                    if convertido and os.path.isfile(ruta_xlsx):
                        return 0, f"{modo_tag}+iopt{iopt_val}_csv→xlsx", ruta_xlsx
                    # CSV disponible aunque no se convirtió
                    return 0, f"{modo_tag}+iopt{iopt_val}_csv", ruta_csv
            except Exception:
                pass

    return (ultimo_ierr if ultimo_ierr is not None else -1), 'all_failed', salida_path


def _escribir_resumen_exportacion(carpeta_salida, project_name, case_name, exportados_info, omitidos_info):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ruta = os.path.join(carpeta_salida, 'Resumen_Exportacion_ShowData.txt')

    lineas = []
    lineas.append('Resumen de exportación - DatosCurvas_v3')
    lineas.append('=' * 72)
    lineas.append(f'Fecha      : {timestamp}')
    lineas.append(f'Proyecto   : {project_name}')
    lineas.append(f'Caso       : {case_name}')
    lineas.append(f'Carpeta    : {carpeta_salida}')
    lineas.append(f'Exportados : {len(exportados_info)}')
    lineas.append(f'Omitidos   : {len(omitidos_info)}')
    lineas.append('')

    lineas.append('Exportados')
    lineas.append('-' * 72)
    if exportados_info:
        for item in exportados_info: # item['csv'] will be item['xlsx']
            lineas.append(
                f"- {item['page']} -> {item['xlsx']} | curvas={item['n_series']} | "
                f"elmres={item['elmres']} | modo={item['modo']} | "
                f"tiempo={'sí' if item.get('incluye_tiempo') else 'no'}"
            )
    else:
        lineas.append('(ninguno)')
    lineas.append('')

    lineas.append('Omitidos / errores')
    lineas.append('-' * 72)
    if omitidos_info:
        for item in omitidos_info:
            base = f"- {item['page']} -> {item['motivo']}"
            if 'ierr' in item:
                base += f" | ierr={item['ierr']}"
            lineas.append(base)
    else:
        lineas.append('(ninguno)')
    lineas.append('')

    try:
        with open(ruta, 'w', encoding='utf-8-sig') as f:
            f.write('\n'.join(lineas))
        return ruta
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────
# Programa principal
# ─────────────────────────────────────────────────────────────
def main():
    app = pf.GetApplication()
    if app is None:
        return

    try:
        app.ClearOutputWindow()
    except Exception:
        pass

    sep = '=' * 68
    app.PrintInfo(sep)
    app.PrintInfo('  DatosCurvas_v3  -  Exportación tipo "Export Show Data"')
    app.PrintInfo(sep)

    project = app.GetActiveProject()
    sc = app.GetActiveStudyCase()

    if project is None or sc is None:
        app.PrintError('No hay proyecto o caso de estudio activo.')
        return

    try:
        CARPETA_SALIDA, semestre, evento, simulacion = _seleccionar_carpeta_salida()
    except RuntimeError as e:
        app.PrintError(str(e))
        return

    app.PrintInfo(f"Proyecto   : {_get_loc_name(project)}")
    app.PrintInfo(f"Caso       : {_get_loc_name(sc)}")
    app.PrintInfo(f"Semestre   : {semestre}")
    app.PrintInfo(f"Evento     : {evento}")
    app.PrintInfo(f"Simulacion : {simulacion}")
    app.PrintInfo(f"Salida     : {CARPETA_SALIDA}")
    app.PrintInfo(f"Modo diag  : {'Sí' if MODO_DIAGNOSTICO else 'No'}")
    app.PrintInfo('')

    os.makedirs(CARPETA_SALIDA, exist_ok=True)

    paginas, paginas_no_encontradas = _buscar_paginas(app, sc)
    app.PrintInfo(f"Páginas encontradas para exportar: {len(paginas)}")
    if paginas_no_encontradas:
        app.PrintInfo(f"Páginas no encontradas en PF    : {len(paginas_no_encontradas)}")
    app.PrintInfo('')

    if not paginas:
        app.PrintWarn('No se encontraron páginas a exportar.')
        return

    if MODO_DIAGNOSTICO:
        if PAGINAS_MODO_DIAGNOSTICO:
            paginas_diag = [p for p in paginas if _get_loc_name(p) in set(PAGINAS_MODO_DIAGNOSTICO)]
        else:
            paginas_diag = paginas

        app.PrintInfo(f"Páginas en diagnóstico: {len(paginas_diag)}")
        if not PAGINAS_MODO_DIAGNOSTICO:
            app.PrintInfo('Diagnóstico configurado para TODAS las páginas encontradas.')
        for page in paginas_diag:
            _diagnosticar_page(app, sc, page)

        app.PrintInfo('')
        app.PrintInfo(sep)
        app.PrintInfo('  Fin del modo diagnóstico')
        app.PrintInfo(sep)
        return

    comres, comres_temporal = _obtener_o_crear_comres(sc)
    if comres is None:
        app.PrintError('No se pudo obtener ni crear un ComRes.')
        return

    estado_original = _capturar_estado_comres(comres)

    exportados = 0
    omitidos = 0
    usados = {}
    exportados_info = []
    omitidos_info = [
        {'page': nom, 'motivo': 'página no encontrada en el proyecto PowerFactory'}
        for nom in paginas_no_encontradas
    ]
    omitidos += len(paginas_no_encontradas)

    try:
        for page in paginas:
            nombre_page = _get_loc_name(page)
            app.PrintInfo(f"→ {nombre_page}")

            series = _extraer_series_de_page(page, app)
            if not series:
                app.PrintWarn('  Omitida: no se pudieron identificar curvas exportables.')
                omitidos_info.append({
                    'page': nombre_page,
                    'motivo': 'sin curvas exportables',
                })
                app.PrintInfo('')
                omitidos += 1
                continue

            elmres = _get_elmres_de_page(page, sc=sc, series=series)
            if elmres is None:
                app.PrintWarn('  Omitida: no se encontró ElmRes asociado.')
                omitidos_info.append({
                    'page': nombre_page,
                    'motivo': 'sin ElmRes asociado',
                })
                app.PrintInfo('')
                omitidos += 1
                continue

            app.PrintInfo(f"  Curvas detectadas : {len(series)}") # cite: 1
            app.PrintInfo(f"  Resultado base    : {_get_loc_name(elmres, 'ElmRes')}") # cite: 1
            app.PrintInfo(f"  Tiempo sim.       : {'Sí' if INCLUIR_TIEMPO_SIMULACION else 'No'}") # cite: 1

            base = _sanear_nombre_archivo(nombre_page)
            usados[base] = usados.get(base, 0) + 1
            sufijo = '' if usados[base] == 1 else f"_{usados[base]}"
            xlsx_path = os.path.join(CARPETA_SALIDA, f"{base}{sufijo}.xlsx")

            ierr, modo_export, ruta_final = _exportar_show_data(
                app, page, elmres, comres, series, xlsx_path
            )

            nombre_final = os.path.basename(ruta_final)
            if ierr == 0 and os.path.isfile(ruta_final):
                tam_kb = os.path.getsize(ruta_final) / 1024.0
                app.PrintInfo(f"  ✓ Exportado       : {nombre_final}")
                app.PrintInfo(f"  Modo usado        : {modo_export}")
                app.PrintInfo(f"  Tamaño aprox.     : {tam_kb:.1f} kB")
                exportados += 1
                exportados_info.append({
                    'page': nombre_page,
                    'xlsx': nombre_final,
                    'n_series': len(series),
                    'elmres': _get_loc_name(elmres, 'ElmRes'),
                    'modo': modo_export,
                    'incluye_tiempo': INCLUIR_TIEMPO_SIMULACION,
                })
            else:
                app.PrintWarn(f"  ✗ Error exportando (ierr={ierr})")
                omitidos += 1
                omitidos_info.append({
                    'page': nombre_page,
                    'motivo': 'error exportando',
                    'ierr': ierr,
                })

            app.PrintInfo('')

    finally:
        _restaurar_estado_comres(comres, estado_original)

        if comres_temporal:
            try:
                comres.Delete()
            except Exception:
                pass

    resumen_path = _escribir_resumen_exportacion(
        carpeta_salida=CARPETA_SALIDA,
        project_name=_get_loc_name(project),
        case_name=_get_loc_name(sc),
        exportados_info=exportados_info,
        omitidos_info=omitidos_info,
    )

    app.PrintInfo(sep)
    app.PrintInfo(f"  Exportados : {exportados}")
    app.PrintInfo(f"  Omitidos   : {omitidos}")
    app.PrintInfo(f"  Carpeta    : {CARPETA_SALIDA}")
    if resumen_path:
        app.PrintInfo(f"  Resumen    : {os.path.basename(resumen_path)}")
    app.PrintInfo(sep)


main()
