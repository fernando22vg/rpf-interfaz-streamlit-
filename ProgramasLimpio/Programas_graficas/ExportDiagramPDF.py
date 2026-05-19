#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ExportDiagramPDF.py
-------------------
Exporta las páginas gráficas de resultados RMS a PDF individuales.
Selección interactiva de semestre / evento / simulación.
Guarda los PDF en:
  RAIZ\{semestre}\Análisis_todos_los_eventos\{evento}\{E{N}.x}\Diagramas PDF\
"""

import os
from datetime import datetime
import powerfactory as pf

# ══════════════════════════════════════════════════════════════
RAIZ = r"C:\Datos del CNDC\01_INFO CNDC_RPF"

# True → solo ejecuta diagnóstico de métodos (no exporta nada).
MODO_DIAGNOSTICO = False

# Si la lista está vacía exporta todas las páginas encontradas.
PAGINAS_A_EXPORTAR = [
    'SIN',
    'STI_MED_PLAZO',
    'F.P. Zon',
    'F.P. TIQ',
    'F.P. BOT',
    'F.P. CUT',
    'F.P. CHU',
    'F.P. HAR',
    'P.F CAH',
    'F.P. HUA',
    'F.P. HUA2',
    'F.P. SRO',
    'F.P. SAI',
    'F.P. SAI1',
    'F.slack',
    'F. Barras SIN',
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

    m = re.search(r"(\d+)", evento)
    n_ev = m.group(1) if m else "0"
    sim_opciones = [f"E{n_ev}.0", f"E{n_ev}.1"]
    simulacion = elegir(sim_opciones, "Tipo de simulación")

    carpeta_salida = os.path.join(base_ev, evento, simulacion, "Diagramas PDF")
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
    for getter in ('GetFullName', 'GetFullPath'):
        try:
            val = getattr(obj, getter)()
            if val:
                return f"{class_name}::{val}"
        except Exception:
            pass
    return f"{class_name}::{loc_name}"


def _mostrar_pagina(app, page):
    try:
        desktop = app.GetGraphicsBoard()
        if desktop is not None:
            desktop.Show(page)
            return True
    except Exception:
        pass
    try:
        app.Show(page)
        return True
    except Exception:
        return False


# ─────────────────────────────────────────────────────────────
# Descubrimiento de páginas
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

    if PAGINAS_A_EXPORTAR:
        orden = {nom: i for i, nom in enumerate(PAGINAS_A_EXPORTAR)}
        paginas = [p for p in paginas if _get_loc_name(p) in orden]
        paginas.sort(key=lambda p: orden.get(_get_loc_name(p), 999999))
    else:
        paginas.sort(key=lambda p: _get_loc_name(p).lower())

    return paginas


# ─────────────────────────────────────────────────────────────
# Diagnóstico de métodos disponibles
# ─────────────────────────────────────────────────────────────
def _diagnosticar_metodos_export(app, page):
    """
    Imprime en el output window todos los métodos del objeto page
    y del graphics board que contengan palabras clave de exportación.
    Ejecutar con MODO_DIAGNOSTICO = True para descubrir el método correcto.
    """
    desktop = app.GetGraphicsBoard()
    palabras_clave = ('write', 'export', 'print', 'save', 'pdf',
                      'bmp', 'png', 'wmf', 'emf', 'image', 'plot')

    for obj, nombre in [(desktop, 'GraphicsBoard'), (page, 'GrpPage')]:
        if obj is None:
            app.PrintInfo(f"  {nombre}: None")
            continue
        app.PrintInfo('')
        app.PrintInfo(f"  ── Métodos de [{nombre}] ──")
        try:
            metodos = [m for m in dir(obj)
                       if any(k in m.lower() for k in palabras_clave)]
            if metodos:
                for m in sorted(metodos):
                    app.PrintInfo(f"    {m}")
            else:
                app.PrintInfo('    (ningún método con palabras clave de exportación)')
        except Exception as e:
            app.PrintInfo(f'    dir() falló: {e}')

        app.PrintInfo(f"  ── GetClassName: {_get_class_name(obj)} ──")
        app.PrintInfo(f"  ── loc_name   : {_get_loc_name(obj)} ──")


# ─────────────────────────────────────────────────────────────
# Exportación a PDF  (vía WriteWMF → conversión a PDF)
# ─────────────────────────────────────────────────────────────
def _archivo_valido(path):
    return os.path.isfile(path) and os.path.getsize(path) > 0


def _wmf_a_pdf_pillow(wmf_path, pdf_path):
    """Convierte WMF a PDF usando Pillow (requiere win32ui en Windows)."""
    from PIL import Image
    img = Image.open(wmf_path)
    img.load()
    img.convert('RGB').save(pdf_path, 'PDF', resolution=200)
    return _archivo_valido(pdf_path)


def _wmf_a_pdf_word(wmf_path, pdf_path):
    """Convierte WMF a PDF abriendo Word via COM y exportando."""
    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Add()
        doc.PageSetup.Orientation = 1          # wdOrientLandscape
        doc.PageSetup.LeftMargin   = 18        # márgenes pequeños (puntos)
        doc.PageSetup.RightMargin  = 18
        doc.PageSetup.TopMargin    = 18
        doc.PageSetup.BottomMargin = 18
        word.Selection.InlineShapes.AddPicture(
            os.path.abspath(wmf_path), False, True)
        doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17, False, 0)
        doc.Close(False)
    finally:
        word.Quit()
    return _archivo_valido(pdf_path)


def _exportar_pagina_pdf(app, page, output_path):
    """
    1. Muestra la página en el Graphics Board.
    2. Exporta a WMF con desktop.WriteWMF().
    3. Convierte el WMF a PDF (Pillow primero, luego Word via COM).
    Devuelve (ok: bool, metodo: str).
    """
    _mostrar_pagina(app, page)

    desktop = app.GetGraphicsBoard()
    if desktop is None:
        return False, 'sin_desktop'

    wmf_path = output_path.replace('.pdf', '.wmf')
    for p in (output_path, wmf_path):
        if os.path.isfile(p):
            try:
                os.remove(p)
            except Exception:
                pass

    # ── Exportar WMF desde PowerFactory ──────────────────────
    try:
        desktop.WriteWMF(wmf_path)
    except Exception:
        return False, 'WriteWMF_excepcion'

    if not _archivo_valido(wmf_path):
        return False, 'WriteWMF_vacio'

    # ── Convertir WMF → PDF ───────────────────────────────────
    for convertidor, nombre in [(_wmf_a_pdf_pillow, 'Pillow'),
                                 (_wmf_a_pdf_word,  'Word')]:
        try:
            if convertidor(wmf_path, output_path):
                try:
                    os.remove(wmf_path)
                except Exception:
                    pass
                return True, f'WriteWMF+{nombre}'
        except ImportError:
            pass
        except Exception:
            pass

    # Sin conversión: conservar WMF como resultado parcial
    return False, 'WMF_ok_conversion_fallida'


def _escribir_resumen(carpeta_salida, project_name, case_name, semestre,
                      evento, simulacion, exportados_info, omitidos_info):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ruta = os.path.join(carpeta_salida, 'Resumen_ExportPDF.txt')

    lineas = [
        'Resumen de exportación - ExportDiagramPDF',
        '=' * 72,
        f'Fecha      : {timestamp}',
        f'Proyecto   : {project_name}',
        f'Caso       : {case_name}',
        f'Semestre   : {semestre}',
        f'Evento     : {evento}',
        f'Simulacion : {simulacion}',
        f'Carpeta    : {carpeta_salida}',
        f'Exportados : {len(exportados_info)}',
        f'Omitidos   : {len(omitidos_info)}',
        '',
        'Exportados',
        '-' * 72,
    ]

    if exportados_info:
        for item in exportados_info:
            lineas.append(
                f"- {item['page']} -> {item['pdf']} | "
                f"metodo={item['metodo']} | {item['kb']:.1f} kB"
            )
    else:
        lineas.append('(ninguno)')

    lineas += ['', 'Omitidos / errores', '-' * 72]

    if omitidos_info:
        for item in omitidos_info:
            lineas.append(f"- {item['page']} -> {item['motivo']}")
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
    app.PrintInfo('  ExportDiagramPDF  -  Exportación de diagramas a PDF')
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
    app.PrintInfo('')

    paginas = _buscar_paginas(app, sc)
    app.PrintInfo(f"Páginas encontradas para exportar: {len(paginas)}")
    app.PrintInfo('')

    if not paginas:
        app.PrintWarn('No se encontraron páginas a exportar.')
        return

    if MODO_DIAGNOSTICO:
        app.PrintInfo('── MODO DIAGNÓSTICO ─────────────────────────────────────')
        _diagnosticar_metodos_export(app, paginas[0])
        app.PrintInfo('')
        app.PrintInfo(sep)
        app.PrintInfo('  Fin del modo diagnóstico. Revisar métodos en output window.')
        app.PrintInfo(sep)
        return

    exportados = 0
    omitidos = 0
    usados = {}
    exportados_info = []
    omitidos_info = []

    for page in paginas:
        nombre_page = _get_loc_name(page)
        app.PrintInfo(f"→ {nombre_page}")

        base = _sanear_nombre_archivo(nombre_page)
        usados[base] = usados.get(base, 0) + 1
        pdf_name = f"{base}.pdf" if usados[base] == 1 else f"{base}_{usados[base]}.pdf"
        pdf_path = os.path.join(CARPETA_SALIDA, pdf_name)

        ok, metodo = _exportar_pagina_pdf(app, page, pdf_path)

        if ok:
            kb = os.path.getsize(pdf_path) / 1024.0
            app.PrintInfo(f"  ✓ Exportado   : {pdf_name}")
            app.PrintInfo(f"  Método        : {metodo}")
            app.PrintInfo(f"  Tamaño aprox. : {kb:.1f} kB")
            exportados += 1
            exportados_info.append({
                'page': nombre_page,
                'pdf': pdf_name,
                'metodo': metodo,
                'kb': kb,
            })
        else:
            app.PrintWarn(f"  ✗ Error exportando")
            omitidos += 1
            omitidos_info.append({
                'page': nombre_page,
                'motivo': f'fallo metodo={metodo}',
            })

        app.PrintInfo('')

    resumen_path = _escribir_resumen(
        carpeta_salida=CARPETA_SALIDA,
        project_name=_get_loc_name(project),
        case_name=_get_loc_name(sc),
        semestre=semestre,
        evento=evento,
        simulacion=simulacion,
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
