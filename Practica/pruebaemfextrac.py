#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ExtractorGraficosEMF.py
Extrae datos numéricos de archivos EMF generados por PowerFactory/CNDC.

Estructura esperada:
  ROOT/{semestre}/Análisis_todos_los_eventos/{Evento N}/{EN.0}/Graficos EMF/
  ROOT/{semestre}/Análisis_todos_los_eventos/{Evento N}/{EN.1}/Graficos EMF/

Salida: un archivo Excel por evento, con una hoja por cada gráfico EMF.
Cada hoja combina las curvas de todos los subdirectorios (E3.0, E3.1, …)
con el prefijo del subdirectorio en el nombre de columna.
"""

import os
import sys
import struct
import numpy as np
import pandas as pd

# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════
ROOT_DATOS       = r"C:\Datos del CNDC\01_INFO CNDC_RPF"
CARPETA_GRAFICOS = "Graficos EMF"
MIN_PUNTOS_CURVA = 20   # mínimo de puntos para considerar una curva de datos
DEBUG            = False

# ══════════════════════════════════════════════════════════════
# CONSTANTES EMF
# ══════════════════════════════════════════════════════════════
EMR_POLYBEZIER           = 2
EMR_POLYLINE             = 4
EMR_POLYPOLYLINE         = 7
EMR_SAVEDC               = 33
EMR_RESTOREDC            = 34
EMR_SETWORLDTRANSFORM    = 35
EMR_MODIFYWORLDTRANSFORM = 36
EMR_SELECTOBJECT         = 37
EMR_CREATEPEN            = 38
EMR_POLYBEZIER16         = 85
EMR_POLYLINE16           = 87
EMR_POLYPOLYLINE16       = 90
EMR_EXTTEXTOUTA          = 83
EMR_EXTTEXTOUTW          = 84
EMR_EXTCREATEPEN         = 95

MWT_IDENTITY       = 1
MWT_LEFTMULTIPLY   = 2
MWT_RIGHTMULTIPLY  = 3

_XFORM_IDENTITY = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


# ══════════════════════════════════════════════════════════════
# PARSER EMF
# ══════════════════════════════════════════════════════════════

def parse_emf(filepath: str) -> list:
    with open(filepath, 'rb') as f:
        raw_file = f.read()
    records = []
    offset = 0
    while offset + 8 <= len(raw_file):
        rtype = struct.unpack_from('<I', raw_file, offset)[0]
        rsize = struct.unpack_from('<I', raw_file, offset + 4)[0]
        if rsize < 8 or offset + rsize > len(raw_file):
            break
        records.append({'type': rtype, 'size': rsize,
                        'raw': raw_file[offset:offset + rsize]})
        offset += rsize
    return records


# ══════════════════════════════════════════════════════════════
# TRANSFORMACIONES AFÍN 2D
# ══════════════════════════════════════════════════════════════

def _xform_apply(xf, x, y):
    m11, m12, m21, m22, dx, dy = xf
    return (m11 * x + m21 * y + dx, m12 * x + m22 * y + dy)


def _xform_multiply(A, B):
    a11, a12, a21, a22, adx, ady = A
    b11, b12, b21, b22, bdx, bdy = B
    return (
        a11 * b11 + a12 * b21, a11 * b12 + a12 * b22,
        a21 * b11 + a22 * b21, a21 * b12 + a22 * b22,
        adx * b11 + ady * b21 + bdx, adx * b12 + ady * b22 + bdy,
    )


def _pts32(raw: bytes, offset: int, count: int) -> list:
    pts = []
    for i in range(count):
        o = offset + i * 8
        if o + 8 > len(raw):
            break
        pts.append(struct.unpack_from('<ii', raw, o))
    return pts


def _pts16(raw: bytes, offset: int, count: int) -> list:
    pts = []
    for i in range(count):
        o = offset + i * 4
        if o + 4 > len(raw):
            break
        pts.append(struct.unpack_from('<hh', raw, o))
    return pts


def extract_polylines(records: list) -> list:
    pen_table: dict = {}
    current_color  = (0, 0, 0)
    current_xform  = _XFORM_IDENTITY
    xform_stack    = []
    polylines      = []

    def _transform_pts(pts):
        return [_xform_apply(current_xform, x, y) for x, y in pts]

    for rec in records:
        rtype = rec['type']
        raw   = rec['raw']

        if rtype == EMR_SAVEDC:
            xform_stack.append(current_xform)
            continue
        elif rtype == EMR_RESTOREDC:
            if xform_stack:
                current_xform = xform_stack.pop()
            continue
        elif rtype == EMR_SETWORLDTRANSFORM:
            if len(raw) >= 32:
                current_xform = struct.unpack_from('<6f', raw, 8)
            continue
        elif rtype == EMR_MODIFYWORLDTRANSFORM:
            if len(raw) >= 36:
                xf   = struct.unpack_from('<6f', raw, 8)
                mode = struct.unpack_from('<I',  raw, 32)[0]
                if mode == MWT_IDENTITY:
                    current_xform = _XFORM_IDENTITY
                elif mode == MWT_LEFTMULTIPLY:
                    current_xform = _xform_multiply(current_xform, xf)
                elif mode == MWT_RIGHTMULTIPLY:
                    current_xform = _xform_multiply(xf, current_xform)
                else:
                    current_xform = xf
            continue

        if rtype == EMR_CREATEPEN:
            if len(raw) >= 28:
                ih = struct.unpack_from('<I', raw, 8)[0]
                cr = struct.unpack_from('<I', raw, 24)[0]
                pen_table[ih] = (cr & 0xFF, (cr >> 8) & 0xFF, (cr >> 16) & 0xFF)
        elif rtype == EMR_EXTCREATEPEN:
            if len(raw) >= 44:
                ih = struct.unpack_from('<I', raw, 8)[0]
                cr = struct.unpack_from('<I', raw, 28 + 4 + 4 + 4)[0]
                pen_table[ih] = (cr & 0xFF, (cr >> 8) & 0xFF, (cr >> 16) & 0xFF)
        elif rtype == EMR_SELECTOBJECT:
            if len(raw) >= 12:
                ih = struct.unpack_from('<I', raw, 8)[0]
                if ih < 0x80000000 and ih in pen_table:
                    current_color = pen_table[ih]
        elif rtype == EMR_POLYLINE:
            if len(raw) >= 28:
                count = struct.unpack_from('<I', raw, 24)[0]
                pts = _transform_pts(_pts32(raw, 28, count))
                if pts:
                    polylines.append({'points': pts, 'color': current_color})
        elif rtype == EMR_POLYLINE16:
            if len(raw) >= 28:
                count = struct.unpack_from('<I', raw, 24)[0]
                pts = _transform_pts(_pts16(raw, 28, count))
                if pts:
                    polylines.append({'points': pts, 'color': current_color})
        elif rtype == EMR_POLYPOLYLINE:
            if len(raw) >= 32:
                npolys = struct.unpack_from('<I', raw, 24)[0]
                total  = struct.unpack_from('<I', raw, 28)[0]
                counts = [struct.unpack_from('<I', raw, 32 + i * 4)[0] for i in range(npolys)]
                all_pts = _transform_pts(_pts32(raw, 32 + npolys * 4, total))
                idx = 0
                for cnt in counts:
                    if all_pts[idx:idx + cnt]:
                        polylines.append({'points': all_pts[idx:idx + cnt], 'color': current_color})
                    idx += cnt
        elif rtype == EMR_POLYPOLYLINE16:
            if len(raw) >= 32:
                npolys = struct.unpack_from('<I', raw, 24)[0]
                total  = struct.unpack_from('<I', raw, 28)[0]
                counts = [struct.unpack_from('<I', raw, 32 + i * 4)[0] for i in range(npolys)]
                all_pts = _transform_pts(_pts16(raw, 32 + npolys * 4, total))
                idx = 0
                for cnt in counts:
                    if all_pts[idx:idx + cnt]:
                        polylines.append({'points': all_pts[idx:idx + cnt], 'color': current_color})
                    idx += cnt
        elif rtype == EMR_POLYBEZIER:
            if len(raw) >= 28:
                count = struct.unpack_from('<I', raw, 24)[0]
                pts = _transform_pts(_pts32(raw, 28, count))
                if pts:
                    polylines.append({'points': pts, 'color': current_color, 'bezier': True})
        elif rtype == EMR_POLYBEZIER16:
            if len(raw) >= 28:
                count = struct.unpack_from('<I', raw, 24)[0]
                pts = _transform_pts(_pts16(raw, 28, count))
                if pts:
                    polylines.append({'points': pts, 'color': current_color, 'bezier': True})

    return polylines


def extract_texts(records: list) -> list:
    texts = []
    for rec in records:
        rtype = rec['type']
        raw   = rec['raw']
        if rtype in (EMR_EXTTEXTOUTW, EMR_EXTTEXTOUTA):
            if len(raw) < 76:
                continue
            x, y   = struct.unpack_from('<ii', raw, 36)
            nchars = struct.unpack_from('<I',  raw, 44)[0]
            offstr = struct.unpack_from('<I',  raw, 48)[0]
            if nchars == 0:
                continue
            try:
                if rtype == EMR_EXTTEXTOUTW:
                    end = offstr + nchars * 2
                    if end > len(raw):
                        continue
                    text = raw[offstr:end].decode('utf-16-le', errors='replace').strip()
                else:
                    end = offstr + nchars
                    if end > len(raw):
                        continue
                    text = raw[offstr:end].decode('latin-1', errors='replace').strip()
                if text:
                    texts.append({'x': x, 'y': y, 'text': text})
            except Exception:
                pass
    return texts


# ══════════════════════════════════════════════════════════════
# DETECCIÓN DE EJES
# ══════════════════════════════════════════════════════════════

def _to_float(s: str):
    try:
        return float(s.replace(',', '.'))
    except (ValueError, AttributeError):
        return None


def detect_axes(texts: list):
    numeric = [
        (t['x'], t['y'], _to_float(t['text']))
        for t in texts if _to_float(t['text']) is not None
    ]
    if len(numeric) < 4:
        return [], [], []

    xs_arr = np.array([n[0] for n in numeric], dtype=float)
    ys_arr = np.array([n[1] for n in numeric], dtype=float)
    coord_range = max(xs_arr.max() - xs_arr.min(), ys_arr.max() - ys_arr.min(), 1)
    thr = coord_range * 0.02

    best_x_grp, best_x_y = [], None
    for _, py, _ in numeric:
        grp = [(px, v) for px, ny, v in numeric if abs(ny - py) <= thr]
        if len(grp) > len(best_x_grp):
            best_x_grp, best_x_y = grp, py
    x_ticks = sorted(best_x_grp, key=lambda t: t[0]) if len(best_x_grp) >= 2 else []

    if best_x_y is not None:
        numeric_y = [(px, py, v) for px, py, v in numeric if abs(py - best_x_y) > thr]
    else:
        numeric_y = list(numeric)

    def _best_y_grp(pool):
        best, best_px = [], None
        for px, _, _ in pool:
            grp = [(py, v) for nx, py, v in pool if abs(nx - px) <= thr]
            if len(grp) > len(best):
                best, best_px = grp, px
        return (sorted(best, key=lambda t: t[0]) if len(best) >= 2 else []), best_px

    y_ticks_a, x_center_a = _best_y_grp(numeric_y)

    if x_center_a is not None:
        pool_b = [(px, py, v) for px, py, v in numeric_y if abs(px - x_center_a) > thr]
    else:
        pool_b = []
    y_ticks_b, _ = _best_y_grp(pool_b) if pool_b else ([], None)

    return x_ticks, y_ticks_a, y_ticks_b


def _interp(pixel: float, ticks: list) -> float:
    if not ticks or len(ticks) < 2:
        return float(pixel)
    p0, v0 = ticks[0]
    p1, v1 = ticks[-1]
    if p1 == p0:
        return float(v0)
    return v0 + (pixel - p0) * (v1 - v0) / (p1 - p0)


def find_axis_unit(texts: list, y_ticks: list) -> str:
    if not y_ticks:
        return ''
    tick_vals = {round(t[1], 4) for t in y_ticks}
    axis_xs = [
        t['x'] for t in texts
        if _to_float(t['text']) is not None
        and round(_to_float(t['text']), 4) in tick_vals
    ]
    if not axis_xs:
        return ''
    x_center = sum(axis_xs) / len(axis_xs)
    candidates = [
        t for t in texts
        if _to_float(t['text']) is None
        and ':' not in t['text']
        and 1 <= len(t['text']) <= 6
        and abs(t['x'] - x_center) < 200
    ]
    if not candidates:
        return ''
    return min(candidates, key=lambda t: abs(t['x'] - x_center))['text']


_KW_FRECUENCIA = ('frequency', 'freq')


def _es_frecuencia(signal_name: str) -> bool:
    sl = signal_name.lower()
    return any(kw in sl for kw in _KW_FRECUENCIA)


# ══════════════════════════════════════════════════════════════
# DETECCIÓN DE LEYENDA
# ══════════════════════════════════════════════════════════════

def _xaxis_label_y(texts: list, x_ticks: list):
    if not x_ticks:
        return None
    tick_vals = {t[1] for t in x_ticks}
    ys = [t['y'] for t in texts
          if _to_float(t['text']) is not None
          and _to_float(t['text']) in tick_vals]
    if not ys:
        return None
    from collections import Counter
    return Counter(ys).most_common(1)[0][0]


def detect_legend(texts: list, polylines: list, x_ticks: list) -> dict:
    label_y = _xaxis_label_y(texts, x_ticks)
    if label_y is None:
        return {}

    leyenda_texts = []
    for t in texts:
        if t['y'] <= label_y:
            continue
        if ':' not in t['text']:
            continue
        code = t['text'].split(':')[0].strip()
        if not code or ' ' in code:
            continue
        if t['text'] in ('Date:', 'Annex:'):
            continue
        leyenda_texts.append(t)

    if not leyenda_texts:
        return {}

    leyenda_markers = []
    for pl in polylines:
        pts = pl['points']
        if len(pts) < 2 or len(pts) > 6:
            continue
        ys_pl = [p[1] for p in pts]
        if min(ys_pl) <= label_y:
            continue
        color = pl.get('color', (0, 0, 0))
        if color == (0, 0, 0):
            continue
        xs_pl = [p[0] for p in pts]
        leyenda_markers.append({
            'color': color,
            'x_min': min(xs_pl),
            'x_max': max(xs_pl),
            'y_mid': sum(ys_pl) / len(ys_pl),
        })

    if not leyenda_markers:
        return {}

    color_label: dict = {}
    tol_y = 20

    for mk in leyenda_markers:
        candidatos = [
            t for t in leyenda_texts
            if t['x'] > mk['x_min']
            and t['y'] >= mk['y_mid']
            and abs(t['y'] - mk['y_mid']) <= tol_y
        ]
        if not candidatos:
            continue
        best = min(candidatos, key=lambda t: t['x'] - mk['x_max'])
        color_label[mk['color']] = best['text']

    return color_label


# ══════════════════════════════════════════════════════════════
# FILTRO Y CONVERSIÓN DE CURVAS
# ══════════════════════════════════════════════════════════════

def extract_data_curves(polylines: list, x_ticks: list,
                        y_ticks_a: list, y_ticks_b: list = None) -> list:
    y_all = list(y_ticks_a) + list(y_ticks_b or [])
    x_plot_min = min(t[0] for t in x_ticks) if x_ticks else None
    x_plot_max = max(t[0] for t in x_ticks) if x_ticks else None
    y_plot_min = min(t[0] for t in y_all) if y_all else None
    y_plot_max = max(t[0] for t in y_all) if y_all else None

    curves = []
    for pl in polylines:
        pts = pl['points']
        if len(pts) < MIN_PUNTOS_CURVA:
            continue
        px_vals = [p[0] for p in pts]
        py_vals = [p[1] for p in pts]
        if x_plot_min is not None:
            if max(px_vals) < x_plot_min or min(px_vals) > x_plot_max:
                continue
        if y_plot_min is not None:
            if max(py_vals) < y_plot_min or min(py_vals) > y_plot_max:
                continue

        t_vals   = [_interp(p[0], x_ticks) for p in pts]
        y_pixels = [p[1] for p in pts]
        r, g, b = pl.get('color', (0, 0, 0))
        curves.append({
            'tiempo':    t_vals,
            'y_pixels':  y_pixels,
            'color_hex': f'#{r:02X}{g:02X}{b:02X}',
            'n_puntos':  len(pts),
        })
    return curves


# ══════════════════════════════════════════════════════════════
# PROCESAMIENTO DE UN ARCHIVO EMF
# ══════════════════════════════════════════════════════════════

def procesar_emf(filepath: str) -> pd.DataFrame | None:
    """Procesa un archivo EMF y retorna un DataFrame con columnas tiempo_s | curva_1 | ..."""
    fname = os.path.basename(filepath)
    try:
        records  = parse_emf(filepath)
        polys    = extract_polylines(records)
        texts    = extract_texts(records)
        x_ticks, y_ticks_a, y_ticks_b = detect_axes(texts)

        if DEBUG:
            print(f"      [DBG] x_ticks: {len(x_ticks)}, y_a: {len(y_ticks_a)}, "
                  f"y_b: {len(y_ticks_b) if y_ticks_b else 0}, polys: {len(polys)}")

        if not x_ticks or not y_ticks_a:
            print(f"      [AVISO] Ejes no detectados en '{fname}'")

        curves = extract_data_curves(polys, x_ticks, y_ticks_a, y_ticks_b)

        if not curves:
            print(f"      [AVISO] Sin curvas en '{fname}' (polilíneas={len(polys)})")
            return None

        legend = detect_legend(texts, polys, x_ticks)

        if y_ticks_b:
            unit_a = find_axis_unit(texts, y_ticks_a)
            unit_b = find_axis_unit(texts, y_ticks_b)
            if 'hz' in unit_a.lower():
                ticks_hz, ticks_ot = y_ticks_a, y_ticks_b
            elif 'hz' in unit_b.lower():
                ticks_hz, ticks_ot = y_ticks_b, y_ticks_a
            else:
                ticks_hz, ticks_ot = y_ticks_a, y_ticks_b
        else:
            ticks_hz = ticks_ot = y_ticks_a

        base   = max(curves, key=lambda c: c['n_puntos'])
        t_base = np.array(base['tiempo'])

        name_count: dict = {}
        data = {'tiempo_s': t_base}
        for i, c in enumerate(curves, start=1):
            rgb = tuple(int(c['color_hex'][j:j+2], 16) for j in (1, 3, 5))
            full_signal = legend.get(rgb, '')
            if full_signal:
                base_name = full_signal.split(':')[0].strip()
                name_count[base_name] = name_count.get(base_name, 0) + 1
                n = name_count[base_name]
                col = base_name if n == 1 else f'{base_name}_{n}'
            else:
                col = f'curva_{i}_{c["color_hex"]}'

            if y_ticks_b and full_signal:
                y_ticks_use = ticks_hz if _es_frecuencia(full_signal) else ticks_ot
            else:
                y_ticks_use = y_ticks_a

            y_vals = [_interp(py, y_ticks_use) for py in c['y_pixels']]
            if len(c['tiempo']) == len(t_base):
                data[col] = y_vals
            else:
                data[col] = np.interp(t_base, c['tiempo'], y_vals)

        return pd.DataFrame(data)

    except Exception as e:
        print(f"      [ERROR] '{fname}': {e}")
        if DEBUG:
            import traceback
            traceback.print_exc()
        return None


# ══════════════════════════════════════════════════════════════
# PROCESAMIENTO POR EVENTO  (múltiples subdirectorios)
# ══════════════════════════════════════════════════════════════

def _combinar_dfs(dfs_por_sub: list) -> pd.DataFrame:
    """
    Combina DataFrames de distintos subdirectorios en uno solo.
    dfs_por_sub: [(prefijo, df), ...]
    Resultado: tiempo_s | {prefijo}_{col} | ...
    Se usa la grilla de tiempo más densa como referencia.
    Las demás series se interpolan a esa grilla.
    """
    if not dfs_por_sub:
        return pd.DataFrame()

    # Grilla de referencia: la más larga
    _, ref_df = max(dfs_por_sub, key=lambda x: len(x[1]))
    t_ref = ref_df['tiempo_s'].values

    combined = {'tiempo_s': t_ref}
    for prefijo, df in dfs_por_sub:
        t_src = df['tiempo_s'].values
        for col in df.columns:
            if col == 'tiempo_s':
                continue
            col_name = f'{prefijo}_{col}'
            vals = df[col].values
            if len(t_src) == len(t_ref) and np.allclose(t_src, t_ref, atol=1e-6):
                combined[col_name] = vals
            else:
                combined[col_name] = np.interp(t_ref, t_src, vals)

    return pd.DataFrame(combined)


def procesar_evento(sem: str, ev: str, ev_path: str,
                    subdirs: list):
    """
    Procesa un evento con múltiples subdirectorios.

    subdirs: [(nombre_subdir, ruta_graficos_emf), ...]
    Genera un Excel en ev_path con una hoja por cada archivo EMF único,
    combinando los datos de todos los subdirectorios.
    """
    # Recopilar todos los nombres de archivos EMF presentes en algún subdir
    emf_nombres: set = set()
    for _, gpath in subdirs:
        for f in os.listdir(gpath):
            if f.lower().endswith('.emf'):
                emf_nombres.add(f)

    if not emf_nombres:
        print(f"  Sin archivos .emf en ningún subdirectorio de: {ev_path}")
        return

    subs_str = ', '.join(s for s, _ in subdirs)
    print(f"\n  >> {sem} / {ev}  [{subs_str}]  ({len(emf_nombres)} gráficos EMF)")

    sem_c  = sem.replace(' ', '_')
    ev_c   = ev.replace(' ', '_')
    output = os.path.join(ev_path, f"{sem_c}_{ev_c}_graficos.xlsx")

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        exitos = 0
        for fname in sorted(emf_nombres):
            sheet_name = os.path.splitext(fname)[0][:31]
            print(f"    -> {fname}")

            dfs_por_sub = []
            for subdir_nombre, gpath in subdirs:
                fpath = os.path.join(gpath, fname)
                if not os.path.isfile(fpath):
                    print(f"       [AVISO] No encontrado en {subdir_nombre}")
                    continue
                df = procesar_emf(fpath)
                if df is not None and not df.empty:
                    dfs_por_sub.append((subdir_nombre, df))
                else:
                    print(f"       [AVISO] Sin datos en {subdir_nombre}")

            if dfs_por_sub:
                df_final = _combinar_dfs(dfs_por_sub)
                df_final.to_excel(writer, sheet_name=sheet_name, index=False)
                n_curvas = len(df_final.columns) - 1
                print(f"       OK {len(df_final)} puntos, {n_curvas} curva(s) "
                      f"({len(dfs_por_sub)}/{len(subdirs)} subdirs)")
                exitos += 1
            else:
                pd.DataFrame({
                    'aviso': [f'No se pudieron extraer datos de {fname}']
                }).to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"  Guardado ({exitos}/{len(emf_nombres)} hojas con datos): {output}")


# ══════════════════════════════════════════════════════════════
# DESCUBRIMIENTO DE EVENTOS
# ══════════════════════════════════════════════════════════════

def _buscar_disponibles(root: str) -> list:
    """
    Retorna lista de (sem, ev, ev_path, subdirs) donde
    subdirs = [(nombre_subdir, ruta_graficos_emf), ...].
    Detecta la estructura:
      root/{sem}/Análisis_todos_los_eventos/{ev}/{subdir}/Graficos EMF/
    """
    disponibles = []
    if not os.path.isdir(root):
        return disponibles

    for sem in sorted(os.listdir(root)):
        base_ev = None
        for cand in ('Análisis_todos_los_eventos', 'Analisis_todos_los_eventos'):
            p = os.path.join(root, sem, cand)
            if os.path.isdir(p):
                base_ev = p
                break
        if not base_ev:
            continue

        for ev in sorted(os.listdir(base_ev)):
            ev_path = os.path.join(base_ev, ev)
            if not os.path.isdir(ev_path):
                continue

            subdirs = []
            for sub in sorted(os.listdir(ev_path)):
                gpath = os.path.join(ev_path, sub, CARPETA_GRAFICOS)
                if os.path.isdir(gpath):
                    n = len([f for f in os.listdir(gpath) if f.lower().endswith('.emf')])
                    if n > 0:
                        subdirs.append((sub, gpath))

            if subdirs:
                disponibles.append((sem, ev, ev_path, subdirs))

    return disponibles


# ══════════════════════════════════════════════════════════════
# FUNCIONES DE NIVEL SUPERIOR
# ══════════════════════════════════════════════════════════════

def procesar_todo(root: str = ROOT_DATOS):
    disponibles = _buscar_disponibles(root)
    if not disponibles:
        print(f"No se encontraron eventos con '{CARPETA_GRAFICOS}' bajo:\n{root}")
        return
    for sem, ev, ev_path, subdirs in disponibles:
        procesar_evento(sem, ev, ev_path, subdirs)
    print(f"\nFIN - {len(disponibles)} evento(s) procesado(s).")


def procesar_semestre(sem: str, root: str = ROOT_DATOS):
    disponibles = [(s, ev, ep, sd) for s, ev, ep, sd
                   in _buscar_disponibles(root) if s == sem]
    if not disponibles:
        print(f"No hay eventos con '{CARPETA_GRAFICOS}' en el semestre '{sem}'.")
        return
    for _, ev, ev_path, subdirs in disponibles:
        procesar_evento(sem, ev, ev_path, subdirs)


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

def separador(titulo: str = "", ancho: int = 62):
    if titulo:
        print(f"\n{'=' * ancho}")
        print(f"  {titulo}")
        print(f"{'=' * ancho}")
    else:
        print(f"{'=' * ancho}")


def elegir(opciones: list, titulo: str) -> str:
    print(f"\n{titulo}:")
    for i, op in enumerate(opciones, 1):
        print(f"  {i}. {op}")
    while True:
        try:
            sel = int(input("  Seleccionar numero: "))
            if 1 <= sel <= len(opciones):
                return opciones[sel - 1]
        except (RuntimeError, EOFError):
            print("\n  ERROR: Este script requiere una terminal interactiva.")
            print("         Ejecutalo desde cmd, PowerShell o un IDE,")
            print("         NO desde DIgSILENT PowerFactory.")
            sys.exit(1)
        except (ValueError, KeyboardInterrupt):
            pass
        print("  Opcion invalida, intente de nuevo.")


if __name__ == '__main__':
    separador("EXTRACTOR DE GRAFICAS EMF  -  CNDC / PowerFactory")
    print(f"  Raiz de datos: {ROOT_DATOS}")

    # ── Paso 1: elegir semestre ────────────────────────────────
    separador("SELECCION DE SEMESTRE")

    semestres = []
    for d in sorted(os.listdir(ROOT_DATOS)):
        dp = os.path.join(ROOT_DATOS, d)
        if not os.path.isdir(dp):
            continue
        for cand in ('Análisis_todos_los_eventos', 'Analisis_todos_los_eventos'):
            if os.path.isdir(os.path.join(dp, cand)):
                semestres.append(d)
                break

    if not semestres:
        print(f"ERROR: No se encontraron semestres validos en:\n{ROOT_DATOS}")
        sys.exit(1)

    semestre = elegir(semestres, "Semestre de estudio")

    # ── Paso 2: listar eventos del semestre ────────────────────
    separador(f"SELECCION DE EVENTO  ({semestre})")

    base_ev = None
    for cand in ('Análisis_todos_los_eventos', 'Analisis_todos_los_eventos'):
        p = os.path.join(ROOT_DATOS, semestre, cand)
        if os.path.isdir(p):
            base_ev = p
            break

    eventos_raw = sorted(
        d for d in os.listdir(base_ev)
        if os.path.isdir(os.path.join(base_ev, d))
    )

    # Para cada evento detectar subdirectorios con Graficos EMF
    estados = []
    for ev in eventos_raw:
        ev_path = os.path.join(base_ev, ev)
        subdirs = []
        for sub in sorted(os.listdir(ev_path)):
            gpath = os.path.join(ev_path, sub, CARPETA_GRAFICOS)
            if os.path.isdir(gpath):
                n = len([f for f in os.listdir(gpath) if f.lower().endswith('.emf')])
                if n > 0:
                    subdirs.append((sub, gpath))
        estados.append((ev, ev_path, subdirs))

    n_con_emf = sum(1 for _, _, sd in estados if sd)

    print(f"\nEventos en '{semestre}':")
    for i, (ev, _, subdirs) in enumerate(estados, 1):
        if subdirs:
            subs_str = ', '.join(s for s, _ in subdirs)
            n_total  = sum(
                len([f for f in os.listdir(gp) if f.lower().endswith('.emf')])
                for _, gp in subdirs
            )
            etiqueta = f"[{subs_str}]  ({n_total} archivos EMF en total)"
        else:
            etiqueta = "[Sin subdirectorios con Graficos EMF]"
        print(f"  {i:2d}. {ev}   {etiqueta}")

    idx_todos = len(estados) + 1
    print(f"\n  {idx_todos:2d}. PROCESAR TODOS los eventos con Graficos EMF"
          f"  ({n_con_emf} de {len(estados)})")
    print()

    while True:
        try:
            sel = int(input("  Seleccionar numero: "))
            if 1 <= sel <= idx_todos:
                break
        except (ValueError, KeyboardInterrupt):
            pass
        print("  Opcion invalida, intente de nuevo.")

    if sel == idx_todos:
        procesar_semestre(semestre)
    else:
        ev_sel, ev_path_sel, subdirs_sel = estados[sel - 1]
        if not subdirs_sel:
            separador()
            print(f"  AVISO: {ev_sel} no tiene subdirectorios con Graficos EMF.")
            print(f"  Se espera estructura:")
            print(f"    {os.path.join(base_ev, ev_sel, 'EN.0', CARPETA_GRAFICOS)}")
            print(f"    {os.path.join(base_ev, ev_sel, 'EN.1', CARPETA_GRAFICOS)}")
            separador()
        else:
            procesar_evento(semestre, ev_sel, ev_path_sel, subdirs_sel)
