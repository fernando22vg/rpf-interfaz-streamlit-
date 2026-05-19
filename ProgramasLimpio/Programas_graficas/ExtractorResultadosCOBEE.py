#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ExtractorResultadosCOBEE.py
Extrae tiempo, frecuencia y potencia de archivos EMF en Resultados_COBEE.

Estructura fija de cada EMF:
  Eje X  -> tiempo
  Eje Y1 -> frecuencia [Hz]
  Eje Y2 -> potencia de la unidad [MW]

Salida: un CSV por archivo EMF  (mismo nombre, misma carpeta)
  columnas: tiempo_s | frecuencia_hz | {nombre_unidad}
"""

import os
import sys
import struct
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
EMR_POLYLINE             = 4
EMR_POLYPOLYLINE         = 7
EMR_POLYBEZIER           = 2
EMR_POLYLINE16           = 87
EMR_POLYPOLYLINE16       = 90
EMR_POLYBEZIER16         = 85
EMR_SAVEDC               = 33
EMR_RESTOREDC            = 34
EMR_SETWORLDTRANSFORM    = 35
EMR_MODIFYWORLDTRANSFORM = 36
EMR_SELECTOBJECT         = 37
EMR_CREATEPEN            = 38
EMR_EXTCREATEPEN         = 95
EMR_EXTTEXTOUTA          = 83
EMR_EXTTEXTOUTW          = 84

MWT_IDENTITY      = 1
MWT_LEFTMULTIPLY  = 2
MWT_RIGHTMULTIPLY = 3
_XFORM_IDENTITY   = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


# ══════════════════════════════════════════════════════════════
# PARSER EMF
# ══════════════════════════════════════════════════════════════

def parse_emf(filepath):
    with open(filepath, 'rb') as f:
        raw = f.read()
    records, offset = [], 0
    while offset + 8 <= len(raw):
        rtype = struct.unpack_from('<I', raw, offset)[0]
        rsize = struct.unpack_from('<I', raw, offset + 4)[0]
        if rsize < 8 or offset + rsize > len(raw):
            break
        records.append({'type': rtype, 'raw': raw[offset:offset + rsize]})
        offset += rsize
    return records


# ══════════════════════════════════════════════════════════════
# TRANSFORMACIONES AFÍN 2D
# ══════════════════════════════════════════════════════════════

def _xform_apply(xf, x, y):
    m11, m12, m21, m22, dx, dy = xf
    return (m11 * x + m21 * y + dx, m12 * x + m22 * y + dy)


def _xform_mul(A, B):
    a11, a12, a21, a22, adx, ady = A
    b11, b12, b21, b22, bdx, bdy = B
    return (a11*b11+a12*b21, a11*b12+a12*b22,
            a21*b11+a22*b21, a21*b12+a22*b22,
            adx*b11+ady*b21+bdx, adx*b12+ady*b22+bdy)


def _pts32(raw, off, n):
    return [struct.unpack_from('<ii', raw, off + i*8) for i in range(n)
            if off + i*8 + 8 <= len(raw)]


def _pts16(raw, off, n):
    return [struct.unpack_from('<hh', raw, off + i*4) for i in range(n)
            if off + i*4 + 4 <= len(raw)]


def extract_polylines(records):
    pen_table, cur_color = {}, (0, 0, 0)
    cur_xf, xf_stack, polylines = _XFORM_IDENTITY, [], []

    def tf(pts):
        return [_xform_apply(cur_xf, x, y) for x, y in pts]

    for rec in records:
        t, raw = rec['type'], rec['raw']

        if t == EMR_SAVEDC:
            xf_stack.append(cur_xf); continue
        if t == EMR_RESTOREDC:
            if xf_stack: cur_xf = xf_stack.pop()
            continue
        if t == EMR_SETWORLDTRANSFORM:
            if len(raw) >= 32: cur_xf = struct.unpack_from('<6f', raw, 8)
            continue
        if t == EMR_MODIFYWORLDTRANSFORM:
            if len(raw) >= 36:
                xf   = struct.unpack_from('<6f', raw, 8)
                mode = struct.unpack_from('<I',  raw, 32)[0]
                if   mode == MWT_IDENTITY:     cur_xf = _XFORM_IDENTITY
                elif mode == MWT_LEFTMULTIPLY: cur_xf = _xform_mul(cur_xf, xf)
                elif mode == MWT_RIGHTMULTIPLY:cur_xf = _xform_mul(xf, cur_xf)
                else:                          cur_xf = xf
            continue
        if t == EMR_CREATEPEN and len(raw) >= 28:
            ih = struct.unpack_from('<I', raw, 8)[0]
            cr = struct.unpack_from('<I', raw, 24)[0]
            pen_table[ih] = (cr&0xFF, (cr>>8)&0xFF, (cr>>16)&0xFF)
        elif t == EMR_EXTCREATEPEN and len(raw) >= 44:
            ih = struct.unpack_from('<I', raw, 8)[0]
            cr = struct.unpack_from('<I', raw, 40)[0]
            pen_table[ih] = (cr&0xFF, (cr>>8)&0xFF, (cr>>16)&0xFF)
        elif t == EMR_SELECTOBJECT and len(raw) >= 12:
            ih = struct.unpack_from('<I', raw, 8)[0]
            if ih < 0x80000000 and ih in pen_table:
                cur_color = pen_table[ih]
        elif t in (EMR_POLYLINE, EMR_POLYBEZIER) and len(raw) >= 28:
            n   = struct.unpack_from('<I', raw, 24)[0]
            pts = tf(_pts32(raw, 28, n))
            if pts: polylines.append({'points': pts, 'color': cur_color})
        elif t in (EMR_POLYLINE16, EMR_POLYBEZIER16) and len(raw) >= 28:
            n   = struct.unpack_from('<I', raw, 24)[0]
            pts = tf(_pts16(raw, 28, n))
            if pts: polylines.append({'points': pts, 'color': cur_color})
        elif t in (EMR_POLYPOLYLINE, EMR_POLYPOLYLINE16) and len(raw) >= 32:
            npolys = struct.unpack_from('<I', raw, 24)[0]
            total  = struct.unpack_from('<I', raw, 28)[0]
            counts = [struct.unpack_from('<I', raw, 32 + i*4)[0] for i in range(npolys)]
            read   = _pts32 if t == EMR_POLYPOLYLINE else _pts16
            all_p  = tf(read(raw, 32 + npolys*4, total))
            idx = 0
            for cnt in counts:
                if all_p[idx:idx+cnt]:
                    polylines.append({'points': all_p[idx:idx+cnt], 'color': cur_color})
                idx += cnt

    return polylines


def extract_texts(records):
    texts = []
    for rec in records:
        t, raw = rec['type'], rec['raw']
        if t not in (EMR_EXTTEXTOUTW, EMR_EXTTEXTOUTA) or len(raw) < 76:
            continue
        x, y   = struct.unpack_from('<ii', raw, 36)
        nchars = struct.unpack_from('<I',  raw, 44)[0]
        offstr = struct.unpack_from('<I',  raw, 48)[0]
        if not nchars:
            continue
        try:
            if t == EMR_EXTTEXTOUTW:
                end  = offstr + nchars * 2
                text = raw[offstr:end].decode('utf-16-le', errors='replace').strip()
            else:
                end  = offstr + nchars
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
        return float(str(s).replace(',', '.'))
    except (ValueError, AttributeError):
        return None


def _parse_time_s(text):
    """Parsea HH:MM:SS o HH:MM → segundos del día. Retorna None si no es hora válida."""
    parts = text.strip().split(':')
    try:
        if len(parts) == 3:
            h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
            if 0 <= h <= 23 and 0 <= m <= 59 and 0 <= s <= 59:
                return h * 3600 + m * 60 + s
        elif len(parts) == 2:
            h, m = int(parts[0]), int(parts[1])
            if 0 <= h <= 23 and 0 <= m <= 59:
                return h * 3600 + m * 60
    except (ValueError, IndexError):
        pass
    return None


def detect_axes(texts):
    # Recopilar valores numéricos (float) y de tiempo (HH:MM:SS → segundos)
    # is_time=True indica que el valor es tiempo y no debe usarse para ejes Y
    all_vals = []
    for t in texts:
        f = _to_float(t['text'])
        if f is not None:
            all_vals.append((t['x'], t['y'], f, False))
            continue
        ts = _parse_time_s(t['text'])
        if ts is not None:
            all_vals.append((t['x'], t['y'], ts, True))

    if len(all_vals) < 4:
        return [], [], []

    xs  = np.array([n[0] for n in all_vals], dtype=float)
    ys  = np.array([n[1] for n in all_vals], dtype=float)
    thr = max(xs.max()-xs.min(), ys.max()-ys.min(), 1) * 0.02

    # Eje X: fila horizontal con más ticks (time o float)
    best_xg, best_xy = [], None
    for _, py, _, _ in all_vals:
        grp = [(px, v) for px, ny, v, _ in all_vals if abs(ny - py) <= thr]
        if len(grp) > len(best_xg):
            best_xg, best_xy = grp, py
    x_ticks = sorted(best_xg, key=lambda t: t[0]) if len(best_xg) >= 2 else []

    # Ejes Y: solo valores float, excluyendo la fila del eje X
    float_pool = [(px, py, v) for px, py, v, it in all_vals
                  if not it and (best_xy is None or abs(py - best_xy) > thr)]

    def best_col(pool):
        best, bx = [], None
        for px, _, _ in pool:
            grp = [(py, v) for nx, py, v in pool if abs(nx - px) <= thr]
            if len(grp) > len(best):
                best, bx = grp, px
        return sorted(best, key=lambda t: t[0]) if len(best) >= 2 else [], bx

    y_a, xc_a = best_col(float_pool)
    pool_b = [(px, py, v) for px, py, v in float_pool
              if xc_a is None or abs(px - xc_a) > thr]
    y_b, _  = best_col(pool_b) if pool_b else ([], None)

    return x_ticks, y_a, y_b


def _interp(pixel, ticks):
    if not ticks or len(ticks) < 2:
        return float(pixel)
    p0, v0 = ticks[0]
    p1, v1 = ticks[-1]
    if p1 == p0:
        return float(v0)
    return v0 + (pixel - p0) * (v1 - v0) / (p1 - p0)


# ══════════════════════════════════════════════════════════════
# FILTRO DE CURVAS DE DATOS
# ══════════════════════════════════════════════════════════════

def extract_data_curves(polylines, x_ticks, y_ticks_a, y_ticks_b=None):
    y_all = list(y_ticks_a) + list(y_ticks_b or [])
    xmin = min(t[0] for t in x_ticks) if x_ticks else None
    xmax = max(t[0] for t in x_ticks) if x_ticks else None
    ymin = min(t[0] for t in y_all)   if y_all   else None
    ymax = max(t[0] for t in y_all)   if y_all   else None

    curves = []
    for pl in polylines:
        pts = pl['points']
        if len(pts) < MIN_PUNTOS_CURVA:
            continue
        pxs = [p[0] for p in pts]
        pys = [p[1] for p in pts]
        if xmin is not None and (max(pxs) < xmin or min(pxs) > xmax):
            continue
        if ymin is not None and (max(pys) < ymin or min(pys) > ymax):
            continue
        curves.append({
            'tiempo':   [_interp(p[0], x_ticks) for p in pts],
            'y_pixels': [p[1] for p in pts],
            'n_puntos': len(pts),
        })
    return curves


# ══════════════════════════════════════════════════════════════
# ASIGNACIÓN DE CURVAS A EJES  (Y1=freq, Y2=potencia)
# ══════════════════════════════════════════════════════════════

def _asignar_curvas(curves, y_ticks_a):
    """
    Y1 (y_ticks_a) = frecuencia  →  valores ~45-65 Hz
    Y2 (y_ticks_b) = potencia    →  cualquier otro rango

    Convierte cada curva con y_ticks_a; la que caiga en rango Hz
    es la frecuencia.  La otra es la potencia (usa y_ticks_b).
    """
    FREQ_MIN, FREQ_MAX = 45.0, 65.0

    freq_curve = pow_curve = None

    for c in sorted(curves, key=lambda c: c['n_puntos'], reverse=True):
        vals_a = np.array([_interp(py, y_ticks_a) for py in c['y_pixels']])
        mean_a = float(np.mean(vals_a))
        if freq_curve is None and FREQ_MIN <= mean_a <= FREQ_MAX:
            freq_curve = c
        elif pow_curve is None:
            pow_curve = c
        if freq_curve and pow_curve:
            break

    # Si ninguna cae en rango Hz, asignar por orden de aparición
    if freq_curve is None and curves:
        freq_curve = curves[0]
    if pow_curve is None:
        rem = [c for c in curves if c is not freq_curve]
        pow_curve = rem[0] if rem else None

    return freq_curve, pow_curve


# ══════════════════════════════════════════════════════════════
# PROCESAMIENTO DE UN ARCHIVO EMF
# ══════════════════════════════════════════════════════════════

def procesar_emf_cobee(filepath, unit_name):
    """Retorna DataFrame: tiempo_s | frecuencia_hz | {unit_name}"""
    fname = os.path.basename(filepath)
    try:
        records = parse_emf(filepath)
        polys   = extract_polylines(records)
        texts   = extract_texts(records)
        x_ticks, y_ticks_a, y_ticks_b = detect_axes(texts)

        if DEBUG:
            print(f"    [DBG] x={len(x_ticks)}, y1={len(y_ticks_a)}, "
                  f"y2={len(y_ticks_b) if y_ticks_b else 0}, polys={len(polys)}")

        if not x_ticks or not y_ticks_a:
            print(f"  [AVISO] Ejes no detectados en '{fname}'")
            return None

        ticks_pow = y_ticks_b if y_ticks_b else y_ticks_a
        curves    = extract_data_curves(polys, x_ticks, y_ticks_a, ticks_pow)

        if not curves:
            print(f"  [AVISO] Sin curvas en '{fname}' (polilíneas={len(polys)})")
            return None

        freq_curve, pow_curve = _asignar_curvas(curves, y_ticks_a)

        ref    = max(curves, key=lambda c: c['n_puntos'])
        t_base = np.array(ref['tiempo'])
        data   = {'tiempo_s': t_base}

        if freq_curve is not None:
            yf = [_interp(py, y_ticks_a) for py in freq_curve['y_pixels']]
            data['frecuencia_hz'] = (yf if len(freq_curve['tiempo']) == len(t_base)
                                     else np.interp(t_base, freq_curve['tiempo'], yf))

        if pow_curve is not None:
            yp = [_interp(py, ticks_pow) for py in pow_curve['y_pixels']]
            data[unit_name] = (yp if len(pow_curve['tiempo']) == len(t_base)
                               else np.interp(t_base, pow_curve['tiempo'], yp))

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
        df = procesar_emf_cobee(os.path.join(carpeta, fname), unit_name)

        print(f"  -> {fname}")
        if df is not None and not df.empty:
            csv_path = os.path.join(carpeta, f"{unit_name}.csv")
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            cols = [c for c in df.columns if c != 'tiempo_s']
            print(f"     OK  {len(df)} puntos | columnas: {', '.join(cols)}")
            exitos += 1
        else:
            print(f"     [AVISO] Sin datos — no se generó CSV")

    print(f"\n  FIN — {exitos}/{len(emf_files)} CSV generados en:\n  {carpeta}")


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
    semestres = sorted(d for d in os.listdir(RAIZ)
                       if os.path.isdir(os.path.join(RAIZ, d)))
    if not semestres:
        print(f"ERROR: No se encontraron semestres en:\n  {RAIZ}")
        sys.exit(1)

    print("\nSelecciona el semestre:")
    for i, s in enumerate(semestres, 1):
        print(f"  {i}. {s}")
    semestre = semestres[elegir_idx(len(semestres), "")]

    # 2. Seleccionar evento
    for cand in ("Análisis_todos_los_eventos", "Analisis_todos_los_eventos"):
        eventos_raiz = os.path.join(RAIZ, semestre, cand)
        if os.path.isdir(eventos_raiz):
            break
    else:
        print(f"ERROR: No se encontró la carpeta de eventos en '{semestre}'.")
        sys.exit(1)

    eventos = sorted(d for d in os.listdir(eventos_raiz)
                     if os.path.isdir(os.path.join(eventos_raiz, d)))
    if not eventos:
        print(f"ERROR: No hay eventos en:\n  {eventos_raiz}")
        sys.exit(1)

    print(f"\nEventos del semestre '{semestre}':")
    for i, ev in enumerate(eventos, 1):
        cobee = os.path.join(eventos_raiz, ev, CARPETA_COBEE)
        n_emf = (len([f for f in os.listdir(cobee) if f.lower().endswith('.emf')])
                 if os.path.isdir(cobee) else 0)
        estado = f"{n_emf} EMF" if n_emf else "sin Resultados_COBEE"
        print(f"  {i:2d}. {ev:<20}  [{estado}]")

    evento = eventos[elegir_idx(len(eventos), "")]

    # 3. Verificar y procesar
    cobee_path = os.path.join(eventos_raiz, evento, CARPETA_COBEE)
    if not os.path.isdir(cobee_path):
        print(f"\nERROR: No se encontró la carpeta:\n  {cobee_path}")
        sys.exit(1)

    print(f"\n  Semestre : {semestre}")
    print(f"  Evento   : {evento}")
    print("=" * 62)

    procesar_carpeta(cobee_path)
    print("=" * 62)

    input("\nPresiona Enter para cerrar...")