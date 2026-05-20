#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ExtractorResultadosCOBEE.py
Extrae tiempo, frecuencia y potencia de archivos EMF en Resultados_COBEE.

Cada EMF tiene 2 ejes: potencia de la unidad (nombre = nombre del archivo) y frecuencia.
Genera un Excel por archivo EMF con columnas: tiempo_s | frecuencia_hz | {nombre_unidad}
"""

import os
import sys
import glob
import struct
import re
import unicodedata
import numpy as np
import pandas as pd

# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════
RAIZ             = r"C:\Datos del CNDC\01_INFO CNDC_RPF"
CARPETA_COBEE    = "Resultados_COBEE"
MIN_PUNTOS_CURVA = 20
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

def parse_emf(filepath):
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


def _pts32(raw, offset, count):
    pts = []
    for i in range(count):
        o = offset + i * 8
        if o + 8 > len(raw):
            break
        pts.append(struct.unpack_from('<ii', raw, o))
    return pts


def _pts16(raw, offset, count):
    pts = []
    for i in range(count):
        o = offset + i * 4
        if o + 4 > len(raw):
            break
        pts.append(struct.unpack_from('<hh', raw, o))
    return pts


def extract_polylines(records):
    pen_table    = {}
    current_color = (0, 0, 0)
    current_xform = _XFORM_IDENTITY
    xform_stack   = []
    polylines     = []

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


def extract_texts(records):
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

def _to_float(s):
    try:
        return float(s.replace(',', '.'))
    except (ValueError, AttributeError):
        return None

def _parse_time_s(text):
    """Parsea HH:MM:SS o HH:MM dentro de un texto y devuelve segundos del día."""
    if text is None:
        return None

    normalized = unicodedata.normalize('NFKC', str(text)).strip()
    if not normalized:
        return None

    normalized = normalized.replace('\u200b', '').replace('\ufeff', '')
    normalized = normalized.replace('：', ':')

    match = re.search(r'(?<!\d)(\d{1,2})[:\s\-_/\\|.hHmMsS]+(\d{2})(?:[:\s\-_/\\|.hHmMsS]+(\d{2}))?(?!\d)', normalized)
    if not match:
        digits = re.sub(r'\D', '', normalized)
        if len(digits) == 6:
            match = re.match(r'(\d{2})(\d{2})(\d{2})$', digits)
        elif len(digits) == 4:
            match = re.match(r'(\d{2})(\d{2})$', digits)

    if not match:
        return None

    h = int(match.group(1))
    m = int(match.group(2))
    s = int(match.group(3) or 0)

    if 0 <= h <= 23 and 0 <= m <= 59 and 0 <= s <= 59:
        return h * 3600 + m * 60 + s
    return None


def _tick_value_match(value, ticks, tol=1e-6):
    return any(abs(value - tick) <= tol for tick in ticks)


def detect_axes(texts):
    items = []
    for t in texts:
        text = t['text']
        f = _to_float(text)
        if f is not None:
            items.append({'x': t['x'], 'y': t['y'], 'value': f, 'is_time': False})
            continue
        ts = _parse_time_s(text)
        if ts is not None:
            items.append({'x': t['x'], 'y': t['y'], 'value': ts, 'is_time': True})

    if len(items) < 4:
        return [], [], []

    xs_arr = np.array([n['x'] for n in items], dtype=float)
    ys_arr = np.array([n['y'] for n in items], dtype=float)
    coord_range = max(xs_arr.max() - xs_arr.min(), ys_arr.max() - ys_arr.min(), 1)
    thr = coord_range * 0.02
    time_thr = max(thr, 24.0)

    time_items = [item for item in items if item['is_time']]

    def _best_y_cluster(pool, current_thr):
        if not pool:
            return [], None

        best, best_y = [], None
        anchors = [item['y'] for item in pool]
        median_y = float(np.median(anchors))

        for anchor_y in anchors:
            grp = [item for item in pool if abs(item['y'] - anchor_y) <= current_thr]
            if len(grp) > len(best):
                best = grp
                best_y = float(anchor_y)

        if len(best) < 2:
            grp = [item for item in pool if abs(item['y'] - median_y) <= current_thr]
            if grp:
                best = grp
                best_y = median_y

        return best, best_y

    best_x_grp, best_x_y = [], None

    if time_items:
        best_x_grp, best_x_y = _best_y_cluster(time_items, time_thr)
        if len(best_x_grp) < 2 and len(time_items) >= 2:
            best_x_grp = time_items
            best_x_y = float(np.median([item['y'] for item in time_items]))

    if not best_x_grp:
        anchor_values = list(ys_arr)
        anchor_ref = float(np.median(anchor_values)) if anchor_values else None
        best_score = None

        for anchor_y in anchor_values:
            grp = [item for item in items if abs(item['y'] - anchor_y) <= thr]
            if len(grp) < 2:
                continue
            time_count = sum(1 for item in grp if item['is_time'])
            if time_items and time_count == 0:
                continue
            score = (
                time_count,
                len(grp),
                -abs(float(anchor_y) - anchor_ref),
            ) if anchor_ref is not None else (time_count, len(grp), 0.0)
            if best_score is None or score > best_score:
                best_x_grp = grp
                best_x_y = float(anchor_y)
                best_score = score

    x_ticks = sorted([(item['x'], item['value']) for item in best_x_grp], key=lambda t: t[0]) if best_x_grp else []

    if best_x_y is not None:
        numeric_y = [
            (item['x'], item['y'], item['value'])
            for item in items
            if not item['is_time'] and abs(item['y'] - best_x_y) > thr
        ]
    else:
        numeric_y = [(item['x'], item['y'], item['value']) for item in items if not item['is_time']]

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


def _interp(pixel, ticks):
    if not ticks or len(ticks) < 2:
        return float(pixel)
    p0, v0 = ticks[0]
    p1, v1 = ticks[-1]
    if p1 == p0:
        return float(v0)
    return v0 + (pixel - p0) * (v1 - v0) / (p1 - p0)


def _format_time_hhmmss(seconds):
    if seconds is None:
        return ''
    try:
        total = float(seconds)
    except (TypeError, ValueError):
        return ''
    if not np.isfinite(total):
        return ''
    whole = int(np.floor(total)) % 86400
    frac = total - np.floor(total)
    hh = whole // 3600
    mm = (whole % 3600) // 60
    ss = whole % 60
    if abs(frac) < 1e-6:
        return f"{hh:02d}:{mm:02d}:{ss:02d}"
    return f"{hh:02d}:{mm:02d}:{ss + frac:06.3f}"


def find_axis_unit(texts, y_ticks):
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


_KW_FRECUENCIA = ('frequency', 'freq', 'frec', 'hz')


def _es_frecuencia(signal_name):
    sl = signal_name.lower()
    return any(kw in sl for kw in _KW_FRECUENCIA)


# ══════════════════════════════════════════════════════════════
# DETECCIÓN DE LEYENDA
# ══════════════════════════════════════════════════════════════

def _xaxis_label_y(texts, x_ticks):
    if not x_ticks:
        return None
    tick_vals = {t[1] for t in x_ticks}
    ys = []
    for t in texts:
        val = _to_float(t['text'])
        if val is None:
            val = _parse_time_s(t['text'])
        if val is not None and val in tick_vals:
            ys.append(t['y'])
    if not ys:
        return None
    from collections import Counter
    return Counter(ys).most_common(1)[0][0]


def detect_legend(texts, polylines, x_ticks):
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

    color_label = {}
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

def extract_data_curves(polylines, x_ticks, y_ticks_a, y_ticks_b=None):
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
        r, g, b  = pl.get('color', (0, 0, 0))
        curves.append({
            'tiempo':    t_vals,
            'y_pixels':  y_pixels,
            'color_hex': f'#{r:02X}{g:02X}{b:02X}',
            'n_puntos':  len(pts),
        })
    return curves


# ══════════════════════════════════════════════════════════════
# IDENTIFICACIÓN POR RANGO DE VALORES (fallback sin leyenda)
# ══════════════════════════════════════════════════════════════

def _identificar_por_rango(curves, y_ticks_a, y_ticks_b):
    FREQ_MIN, FREQ_MAX = 45.0, 65.0

    candidatos = []
    for c in curves:
        for ticks, label in [(y_ticks_a, 'a')] + ([(y_ticks_b, 'b')] if y_ticks_b else []):
            vals   = np.array([_interp(py, ticks) for py in c['y_pixels']])
            mean_v = float(np.mean(vals))
            candidatos.append({
                'curve': c, 'ticks': ticks, 'axis': label,
                'mean': mean_v, 'is_freq': FREQ_MIN <= mean_v <= FREQ_MAX,
            })

    freq_cands = [c for c in candidatos if c['is_freq']]
    if not freq_cands:
        return None, None, None, None

    best_freq  = min(freq_cands, key=lambda c: abs(c['mean'] - 50.0))
    freq_curve = best_freq['curve']
    ticks_freq = best_freq['ticks']

    pow_candidates = sorted(
        [c for c in curves if c is not freq_curve],
        key=lambda c: c['n_puntos'], reverse=True
    )
    if not pow_candidates:
        return freq_curve, ticks_freq, None, None

    pow_curve = pow_candidates[0]
    if y_ticks_b:
        ticks_pow = y_ticks_b if best_freq['axis'] == 'a' else y_ticks_a
    else:
        ticks_pow = y_ticks_a

    return freq_curve, ticks_freq, pow_curve, ticks_pow


# ══════════════════════════════════════════════════════════════
# PROCESAMIENTO DE UN ARCHIVO EMF
# ══════════════════════════════════════════════════════════════

def procesar_emf_cobee(filepath, unit_name):
    """Retorna DataFrame: tiempo_s | frecuencia_hz | {unit_name}"""
    fname = os.path.basename(filepath)
    try:
        records  = parse_emf(filepath)
        polys    = extract_polylines(records)
        texts    = extract_texts(records)
        x_ticks, y_ticks_a, y_ticks_b = detect_axes(texts)

        if DEBUG:
            print(f"    [DBG] x={len(x_ticks)}, y_a={len(y_ticks_a)}, "
                  f"y_b={len(y_ticks_b) if y_ticks_b else 0}, polys={len(polys)}")

        if not x_ticks or not y_ticks_a:
            print(f"  [AVISO] Ejes no detectados en '{fname}'")
            return None

        curves = extract_data_curves(polys, x_ticks, y_ticks_a, y_ticks_b)
        if not curves:
            print(f"  [AVISO] Sin curvas en '{fname}' (polilíneas={len(polys)})")
            return None

        legend = detect_legend(texts, polys, x_ticks)

        # Determinar ejes por unidad detectada
        if y_ticks_b:
            unit_a = find_axis_unit(texts, y_ticks_a)
            unit_b = find_axis_unit(texts, y_ticks_b)
            if DEBUG:
                print(f"    [DBG] eje_a='{unit_a}', eje_b='{unit_b}'")
            if 'hz' in unit_a.lower():
                ticks_freq_def, ticks_pow_def = y_ticks_a, y_ticks_b
            elif 'hz' in unit_b.lower():
                ticks_freq_def, ticks_pow_def = y_ticks_b, y_ticks_a
            else:
                ticks_freq_def, ticks_pow_def = y_ticks_b, y_ticks_a
        else:
            ticks_freq_def = ticks_pow_def = y_ticks_a

        # Asignar curvas via leyenda
        freq_curve = None
        freq_ticks = ticks_freq_def
        pow_curve  = None
        pow_ticks  = ticks_pow_def

        for c in curves:
            rgb    = tuple(int(c['color_hex'][j:j + 2], 16) for j in (1, 3, 5))
            signal = legend.get(rgb, '')
            if DEBUG and signal:
                print(f"    [DBG] señal '{signal}' ({c['color_hex']})")
            if signal and _es_frecuencia(signal):
                freq_curve = c
            elif signal and unit_name.upper() in signal.upper():
                pow_curve  = c

        # Fallback por rango de valores
        if freq_curve is None or pow_curve is None:
            fc, ft, pc, pt = _identificar_por_rango(curves, y_ticks_a, y_ticks_b)
            if freq_curve is None and fc is not None:
                freq_curve, freq_ticks = fc, ft
            if pow_curve is None and pc is not None:
                pow_curve, pow_ticks = pc, pt

        # Último recurso: asignar por orden
        sorted_c = sorted(curves, key=lambda c: c['n_puntos'], reverse=True)
        if freq_curve is None and pow_curve is None:
            if len(sorted_c) >= 2:
                freq_curve, pow_curve = sorted_c[0], sorted_c[1]
            elif sorted_c:
                freq_curve = sorted_c[0]
        elif freq_curve is None:
            rem = [c for c in sorted_c if c is not pow_curve]
            freq_curve = rem[0] if rem else None
        elif pow_curve is None:
            rem = [c for c in sorted_c if c is not freq_curve]
            pow_curve = rem[0] if rem else None

        # Construir DataFrame
        ref    = max(curves, key=lambda c: c['n_puntos'])
        t_base = np.array(ref['tiempo'])
        data   = {
            'tiempo_s': t_base,
            'hora': [_format_time_hhmmss(t) for t in t_base],
        }

        if freq_curve is not None:
            y_vals = [_interp(py, freq_ticks) for py in freq_curve['y_pixels']]
            data['frecuencia_hz'] = (y_vals if len(freq_curve['tiempo']) == len(t_base)
                                     else np.interp(t_base, freq_curve['tiempo'], y_vals))

        if pow_curve is not None:
            y_vals = [_interp(py, pow_ticks) for py in pow_curve['y_pixels']]
            data[unit_name] = (y_vals if len(pow_curve['tiempo']) == len(t_base)
                               else np.interp(t_base, pow_curve['tiempo'], y_vals))

        df = pd.DataFrame(data)
        return df if not df.empty else None

    except Exception as e:
        print(f"  [ERROR] '{fname}': {e}")
        if DEBUG:
            import traceback
            traceback.print_exc()
        return None


# ══════════════════════════════════════════════════════════════
# PROCESAMIENTO DE LA CARPETA
# ══════════════════════════════════════════════════════════════

def procesar_carpeta(carpeta):
    emf_files = sorted(f for f in os.listdir(carpeta) if f.lower().endswith('.emf'))
    if not emf_files:
        print(f"  No se encontraron archivos .emf en:\n  {carpeta}")
        return

    print(f"\n  Carpeta : {carpeta}")
    print(f"  Archivos: {len(emf_files)} EMF\n")

    exitos = 0
    for fname in emf_files:
        unit_name = os.path.splitext(fname)[0]
        fpath     = os.path.join(carpeta, fname)
        xlsx_path = os.path.join(carpeta, f"{unit_name}.xlsx")

        print(f"  -> {fname}")
        df = procesar_emf_cobee(fpath, unit_name)

        if df is not None and not df.empty:
            df.to_excel(xlsx_path, index=False)
            cols = [c for c in df.columns if c != 'tiempo_s']
            print(f"     OK  {len(df)} puntos | columnas: {', '.join(cols)}")
            exitos += 1
        else:
            print(f"     [AVISO] Sin datos — no se generó Excel")

    print(f"\n  FIN — {exitos}/{len(emf_files)} Excel generados en:\n  {carpeta}")


# ══════════════════════════════════════════════════════════════
# SELECCIÓN INTERACTIVA
# ══════════════════════════════════════════════════════════════

def elegir_idx(n, titulo):
    print(titulo)
    while True:
        try:
            idx = int(input("Selecciona numero: ")) - 1
            if 0 <= idx < n:
                return idx
        except (ValueError, KeyboardInterrupt):
            pass
        except (RuntimeError, EOFError):
            print("\n  ERROR: Este script requiere una terminal interactiva.")
            sys.exit(1)
        print("  Opcion invalida, intenta de nuevo.")


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print("\n" + "=" * 62)
    print("  EXTRACTOR RESULTADOS COBEE  —  EMF → CSV")
    print("=" * 62)

    # 1. Seleccionar semestre
    semestres = sorted(d for d in os.listdir(RAIZ) if os.path.isdir(os.path.join(RAIZ, d)))
    if not semestres:
        print(f"ERROR: No se encontraron semestres en:\n  {RAIZ}")
        sys.exit(1)

    print("\nSelecciona el semestre:")
    for i, s in enumerate(semestres, 1):
        print(f"  {i}. {s}")
    semestre = semestres[elegir_idx(len(semestres), "")]

    # 2. Seleccionar evento
    eventos_raiz = os.path.join(RAIZ, semestre, "Análisis_todos_los_eventos")
    if not os.path.isdir(eventos_raiz):
        eventos_raiz = os.path.join(RAIZ, semestre, "Analisis_todos_los_eventos")
    if not os.path.isdir(eventos_raiz):
        print(f"ERROR: No se encontró la carpeta de eventos en '{semestre}'.")
        sys.exit(1)

    eventos = sorted(
        d for d in os.listdir(eventos_raiz)
        if os.path.isdir(os.path.join(eventos_raiz, d))
    )
    if not eventos:
        print(f"ERROR: No hay eventos en:\n  {eventos_raiz}")
        sys.exit(1)

    print(f"\nEventos del semestre '{semestre}':")
    for i, ev in enumerate(eventos, 1):
        cobee_path = os.path.join(eventos_raiz, ev, CARPETA_COBEE)
        n_emf = len([f for f in os.listdir(cobee_path)
                     if f.lower().endswith('.emf')]) if os.path.isdir(cobee_path) else 0
        estado = f"{n_emf} EMF" if n_emf else "sin Resultados_COBEE"
        print(f"  {i:2d}. {ev:<20}  [{estado}]")

    evento = eventos[elegir_idx(len(eventos), "")]

    # 3. Verificar carpeta Resultados_COBEE
    cobee_path = os.path.join(eventos_raiz, evento, CARPETA_COBEE)
    if not os.path.isdir(cobee_path):
        print(f"\nERROR: No se encontró la carpeta:\n  {cobee_path}")
        sys.exit(1)

    print(f"\n  Semestre : {semestre}")
    print(f"  Evento   : {evento}")
    print(f"  Carpeta  : {cobee_path}")

    # 4. Procesar
    print("\n" + "=" * 62)
    procesar_carpeta(cobee_path)
    print("=" * 62)

    input("\nPresiona Enter para cerrar...")
