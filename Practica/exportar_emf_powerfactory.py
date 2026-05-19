#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
exportar_emf_powerfactory.py
Exporta las gráficas de los casos RMS de PowerFactory como archivos EMF.

Debe ejecutarse DESDE DIgSILENT PowerFactory (menú Scripts o consola Python
integrada). No funciona como script externo independiente.

Estructura de salida generada:
  ROOT_SALIDA/{SEMESTRE}/Análisis_todos_los_eventos/Evento {N}/{EN.x}/Graficos EMF/

Detecta automáticamente los casos de estudio con nombre que siga el
patrón E{N}.{x}  (ej. E3.0, E3.1, E12.0, E12.1 …).
Si un caso no existe en el proyecto se omite y continúa con el siguiente.
"""

import os
import re
import sys

# ══════════════════════════════════════════════════════════════
# VERIFICAR ENTORNO POWERFACTORY
# ══════════════════════════════════════════════════════════════
try:
    import powerfactory as pf          # módulo solo disponible dentro de PF
    app = pf.GetApplication()
    if app is None:
        raise RuntimeError("GetApplication() retornó None")
except ImportError:
    print("ERROR: Este script debe ejecutarse desde DIgSILENT PowerFactory.")
    print("       Usa: Herramientas > Scripts Python, o la consola integrada.")
    sys.exit(1)

# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN  — ajustar según el proyecto
# ══════════════════════════════════════════════════════════════

# Ruta raíz donde se crearán las carpetas de salida
ROOT_SALIDA = r"C:\Datos del CNDC\01_INFO CNDC_RPF"

# Nombre del semestre (aparecerá en la ruta de salida)
SEMESTRE = "2025 sem1"

# Nombre de la carpeta de análisis dentro de cada semestre
CARPETA_ANALISIS = "Análisis_todos_los_eventos"

# Subcarpeta donde se guardan los EMF dentro de cada caso
CARPETA_EMF = "Graficos EMF"

# Patrón que deben tener los nombres de los casos de estudio.
# Acepta variantes como:
#   E3.0
#   E3_0
#   E3-0
#   E3.0 - Base
#   e5.1
PATRON_CASO = re.compile(r'E\s*(\d+)\s*[.\-_]\s*(\d+)', re.IGNORECASE)

# Nombre del panel de instrumentos virtual a buscar dentro del caso.
# Si es '', exporta TODOS los ComVnt que encuentre.
NOMBRE_PANEL = ''

# ══════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════

def msg(texto: str):
    """Imprime en la consola de PowerFactory y en stdout."""
    app.PrintPlain(str(texto))
    print(texto)


def warn(texto: str):
    app.PrintWarn(str(texto))
    print(f"[AVISO] {texto}")


def err(texto: str):
    app.PrintError(str(texto))
    print(f"[ERROR] {texto}")


def carpeta_salida(n_evento: int, sufijo: str) -> str:
    """Construye la ruta de salida para un caso dado."""
    return os.path.join(
        ROOT_SALIDA,
        SEMESTRE,
        CARPETA_ANALISIS,
        f"Evento {n_evento}",
        f"E{n_evento}.{sufijo}",
        CARPETA_EMF,
    )


# ══════════════════════════════════════════════════════════════
# BÚSQUEDA DE CASOS DE ESTUDIO
# ══════════════════════════════════════════════════════════════

def buscar_casos() -> list:
    """
    Retorna lista de (IntCase, n_evento:int, sufijo:str) para todos los
    casos cuyo nombre coincide con PATRON_CASO en el proyecto activo.
    """
    prj = app.GetActiveProject()
    if prj is None:
        err("No hay proyecto activo en PowerFactory.")
        return []

    todos = prj.GetContents('*.IntCase', 1)   # 1 = recursivo en todo el árbol
    encontrados = []

    for caso in todos:
        nombre = ""
        try:
            nombre = (caso.loc_name or "").strip()
        except Exception:
            nombre = ""

        m = PATRON_CASO.search(nombre)
        if m is None:
            try:
                # Fallback adicional por si el nombre útil aparece en la ruta completa
                m = PATRON_CASO.search(caso.GetFullName())
            except Exception:
                m = None

        if m:
            n_ev = int(m.group(1))
            sufijo = m.group(2)
            encontrados.append((caso, n_ev, sufijo))

    encontrados.sort(key=lambda x: (x[1], x[2]))
    return encontrados


# ══════════════════════════════════════════════════════════════
# EXPORTACIÓN DE PÁGINAS EMF
# ══════════════════════════════════════════════════════════════

def _obtener_paginas(caso) -> list:
    """
    Obtiene todas las SetPage de los paneles de instrumentos virtuales
    (ComVnt) del caso activo.
    Busca primero en el estudio activo; si no encuentra, busca en el proyecto.
    """
    paginas = []

    # 1. Buscar en el caso de estudio activo
    paneles = caso.GetContents('*.ComVnt', 1)

    # 2. Si no hay en el caso, buscar a nivel de proyecto (paneles compartidos)
    if not paneles:
        prj = app.GetActiveProject()
        if prj:
            paneles = prj.GetContents('*.ComVnt', 1)

    if not paneles:
        return paginas

    for panel in paneles:
        nombre_panel = panel.loc_name
        if NOMBRE_PANEL and nombre_panel != NOMBRE_PANEL:
            continue
        pages = panel.GetContents('*.SetPage')
        paginas.extend(pages)

    return paginas


def _exportar_pagina_emf(pagina, filepath: str) -> bool:
    """
    Intenta exportar una página como EMF usando los métodos disponibles.
    Prueba varios enfoques en orden de preferencia.
    Retorna True si tuvo éxito.
    """

    # ── Método 1: mostrar página y exportar vía tablero de gráficos ──────────
    try:
        pagina.Show()                          # activa la página en pantalla
        grfbrd = app.GetGraphicsBoard()
        if grfbrd is not None:
            # iopt_exp: 0=BMP, 1=WMF/EMF, 2=PNG, 3=JPG (según versión PF)
            ret = grfbrd.Export(filepath, 1)
            if ret == 0:
                return True
    except Exception as e:
        warn(f"      Método 1 falló ({e})")

    # ── Método 2: Export directo sobre la página ──────────────────────────────
    try:
        pagina.Show()
        ret = pagina.Export(filepath)
        if ret == 0:
            return True
    except Exception as e:
        warn(f"      Método 2 falló ({e})")

    # ── Método 3: ComGrfexp (objeto de exportación de gráficos) ──────────────
    try:
        grfexp = app.GetFromStudyCase('ComGrfexp')
        if grfexp is not None:
            pagina.Show()
            grfexp.pGrfbook = app.GetGraphicsBoard()
            grfexp.f_name   = filepath
            grfexp.iopt_typ = 1    # 1 = EMF/WMF
            ret = grfexp.Execute()
            if ret == 0:
                return True
    except Exception as e:
        warn(f"      Método 3 falló ({e})")

    # ── Método 4: ComPrint con formato EMF ───────────────────────────────────
    try:
        comprint = app.GetFromStudyCase('ComPrint')
        if comprint is not None:
            pagina.Show()
            comprint.iopt_typ = 4   # EMF (varía según versión PF)
            comprint.f_name   = filepath
            ret = comprint.Execute()
            if ret == 0:
                return True
    except Exception as e:
        warn(f"      Método 4 falló ({e})")

    return False


def exportar_caso(caso, n_evento: int, sufijo: str) -> int:
    """
    Activa el caso de estudio, obtiene sus páginas y exporta cada una como EMF.
    Retorna el número de archivos exportados correctamente.
    """
    nombre_caso = caso.loc_name
    ruta_salida = carpeta_salida(n_evento, sufijo)
    os.makedirs(ruta_salida, exist_ok=True)

    msg(f"\n  Activando caso: {nombre_caso}")

    try:
        ret_act = caso.Activate()
        if ret_act != 0:
            warn(f"  Activate() retornó {ret_act} para '{nombre_caso}' — se continúa de todas formas")
    except Exception as e:
        warn(f"  No se pudo activar '{nombre_caso}': {e}")
        return 0

    paginas = _obtener_paginas(caso)
    if not paginas:
        warn(f"  Sin páginas (SetPage) en el caso '{nombre_caso}'")
        return 0

    msg(f"  {len(paginas)} página(s) encontrada(s) → {ruta_salida}")

    exportados = 0
    for pagina in paginas:
        nombre_pag = (pagina.loc_name or f"pagina_{exportados + 1}") \
                     .replace('/', '_').replace('\\', '_').replace(':', '-')
        filepath = os.path.join(ruta_salida, f"{nombre_pag}.emf")

        msg(f"    -> {nombre_pag}.emf")
        if _exportar_pagina_emf(pagina, filepath):
            exportados += 1
            msg(f"       OK")
        else:
            err(f"       No se pudo exportar '{nombre_pag}' con ningún método")

    return exportados


# ══════════════════════════════════════════════════════════════
# FLUJO PRINCIPAL
# ══════════════════════════════════════════════════════════════

def main():
    sep = "=" * 62
    msg(sep)
    msg("  EXPORTADOR EMF  -  PowerFactory → Excel EMF")
    msg(f"  Salida raíz : {ROOT_SALIDA}")
    msg(f"  Semestre    : {SEMESTRE}")
    msg(sep)

    casos = buscar_casos()
    if not casos:
        err("No se encontraron casos de estudio que coincidan con el patrón esperado.")
        try:
            todos = app.GetActiveProject().GetContents('*.IntCase', 1) or []
            if todos:
                warn("Casos IntCase disponibles en el proyecto:")
                for caso in todos[:20]:
                    warn(f"  - {caso.loc_name}")
        except Exception:
            pass
        err("Verifica el nombre real de tus IntCase o ajusta PATRON_CASO.")
        return

    # Resumen de lo que se va a procesar
    msg(f"\nCasos encontrados ({len(casos)}):")
    for caso, n_ev, suf in casos:
        msg(f"  E{n_ev}.{suf}  →  {caso.loc_name}  [{caso.GetFullName()}]")

    msg(f"\n{'─' * 62}")

    # Guardar caso activo original para restaurar al final
    caso_original = app.GetActiveStudyCase()

    total_exportados = 0
    total_fallidos   = 0

    for caso, n_ev, suf in casos:
        try:
            n = exportar_caso(caso, n_ev, suf)
            total_exportados += n
            if n == 0:
                total_fallidos += 1
        except Exception as e:
            err(f"  Excepción procesando E{n_ev}.{suf}: {e}")
            total_fallidos += 1

    # Restaurar caso original
    if caso_original is not None:
        try:
            caso_original.Activate()
        except Exception:
            pass

    msg(f"\n{sep}")
    msg(f"  FIN  —  {total_exportados} archivos EMF exportados")
    if total_fallidos:
        warn(f"  {total_fallidos} caso(s) sin exportación exitosa")
    msg(sep)


if __name__ == '__main__':
    main()
