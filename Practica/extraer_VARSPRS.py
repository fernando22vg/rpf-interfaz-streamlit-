#!/usr/bin/env python3
"""
extraer_VARSPRS.py
══════════════════
Extrae parámetros de regulación desde archivos VARSPRS.DAT del sistema
RTVX Power (Reivax) y genera un CSV compatible con Excel.

Uso:
    python extraer_VARSPRS.py VARSPRS.DAT
    python extraer_VARSPRS.py VARSPRS.DAT -o salida.csv
    python extraer_VARSPRS.py VARSPRS.DAT --todos
    python extraer_VARSPRS.py VARSPRS.DAT --modulos RTX_OEL RTX_MEL

Opciones:
    -o, --output      Nombre del archivo CSV de salida (default: <entrada>_parametros.csv)
    --todos           Exporta TODOS los 1694 parámetros (no solo regulación)
    --modulos         Lista de módulos específicos a extraer
    --separador       Separador CSV (default: ;)
    --sin-descripcion Omite la columna de descripción automática
    --json            Salida en formato JSON en lugar de CSV

Requiere: Python 3.10+  (sin dependencias externas)
Autor: Generado para proyecto Zongo BOT03 — COBEE/Andritz/Reivax
"""

import re
import csv
import json
import sys
import argparse
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional


# ═══════════════════════════════════════════════════════════════════════
# 1. MODELO DE DATOS
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class VarEntry:
    """Representa una variable del archivo VARSPRS.DAT."""
    var_num: int
    name: str
    module: str
    tp: int            # Tipo: 8=INT, 9=BOOL, 10=FLOAT
    raw_value: str
    value: float | int | str = 0
    description: str = ""
    category: str = ""
    dsl_equiv: str = ""

    def __post_init__(self):
        try:
            if self.tp == 9:  # BOOL
                self.value = int(float(self.raw_value))
            elif self.tp == 8:  # INT
                self.value = int(float(self.raw_value))
            else:  # FLOAT (tp=10)
                self.value = float(self.raw_value)
        except (ValueError, TypeError):
            self.value = self.raw_value

    @property
    def full_name(self) -> str:
        return f"{self.name}@{self.module}"

    @property
    def type_str(self) -> str:
        return {8: "INT", 9: "BOOL", 10: "FLOAT"}.get(self.tp, f"TP{self.tp}")

    def value_fmt(self) -> str:
        """Formatea el valor para CSV."""
        if isinstance(self.value, float):
            if self.value == int(self.value) and abs(self.value) < 1e10:
                return str(int(self.value))
            return f"{self.value:.6g}"
        return str(self.value)


# ═══════════════════════════════════════════════════════════════════════
# 2. PARSER DEL ARCHIVO DAT
# ═══════════════════════════════════════════════════════════════════════

def parse_dat(filepath: str | Path) -> list[VarEntry]:
    """
    Parsea un archivo VARSPRS.DAT y retorna lista de VarEntry.

    Formato esperado:
        [Var N]
        NS=NombreVariable@Modulo
        Tp=10
        V=+1.234567890000000E+002
    """
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(f"No se encontró: {filepath}")

    # Intentar varias codificaciones
    content = None
    for enc in ('utf-8', 'latin-1', 'cp1252', 'iso-8859-1'):
        try:
            content = filepath.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue

    if content is None:
        raise ValueError(f"No se pudo leer {filepath} con ninguna codificación soportada")

    entries: list[VarEntry] = []
    pattern = re.compile(
        r'\[Var\s+(\d+)\]\s*\n'
        r'NS=(.+?)@(.+?)\s*\n'
        r'Tp=(\d+)\s*\n'
        r'V=(.+)',
        re.MULTILINE
    )

    for m in pattern.finditer(content):
        entry = VarEntry(
            var_num=int(m.group(1)),
            name=m.group(2).strip(),
            module=m.group(3).strip(),
            tp=int(m.group(4)),
            raw_value=m.group(5).strip()
        )
        entries.append(entry)

    return entries


# ═══════════════════════════════════════════════════════════════════════
# 3. DICCIONARIO DE DESCRIPCIÓN Y EQUIVALENCIA DSL
# ═══════════════════════════════════════════════════════════════════════

# Módulos de regulación relevantes
MODULOS_REGULACION = {
    'RTX_AI', 'RTX_CTR', 'RTX_FAIL', 'RTX_FREQ_RESP', 'RTX_LOCAL',
    'RTX_LOG', 'RTX_MEL', 'RTX_OEL', 'RTX_PRESET', 'RTX_SCL',
    'RTX_UEL', 'RTX_VHZ',
    'RVX_AI', 'RVX_AI_PELTON', 'RVX_CTR', 'RVX_CTR_PELTON',
    'RVX_CTR_PELTON_T2', 'RVX_FAIL', 'RVX_FAIL_PELTON',
    'RVX_LOCAL', 'RVX_LOG', 'RVX_LOG_PELTON', 'RVX_Pe_x_Y',
    'RVX_REG',
    'SYS_LOCAL',
}

# (name, module) → (descripción, símbolo DSL equivalente, tabla DSL)
# Solo los parámetros con equivalencia directa o importancia para modelado
PARAM_INFO: dict[tuple[str, str], tuple[str, str, str]] = {
    # ── SYS_LOCAL ──
    ('Xs', 'SYS_LOCAL'):          ('Reactancia síncrona Xs (pu)', 'Xs (GENROU)', 'Generador'),
    ('Xq', 'SYS_LOCAL'):          ('Reactancia eje q Xq (pu)', 'Xq (GENROU) / Xcomp_PSS', 'Generador / Tabla 92'),
    ('Xd', 'SYS_LOCAL'):          ('Reactancia síncrona eje d Xd (pu)', 'Xd (GENROU)', 'Generador'),
    ('Xl', 'SYS_LOCAL'):          ('Reactancia de dispersión Xl (pu)', 'Xl (GENROU)', 'Generador'),
    ('Itmax', 'SYS_LOCAL'):       ('Corriente estatórica máxima (pu)', 'REF_SCL_TERM', 'Tabla 97'),
    ('NominalPF', 'SYS_LOCAL'):   ('Factor de potencia nominal', '—', 'Dato de placa'),
    ('Efdmax', 'SYS_LOCAL'):      ('Efd máximo para curva capabilidad (pu)', '—', '—'),
    ('Efdmin', 'SYS_LOCAL'):      ('Efd mínimo para curva capabilidad (pu)', '—', '—'),
    ('MaxAngInt', 'SYS_LOCAL'):    ('Ángulo interno máximo (°)', '—', '—'),
    ('PNom', 'SYS_LOCAL'):         ('Potencia nominal (pu)', '—', '—'),
    ('Saliency', 'SYS_LOCAL'):     ('Tipo máquina (1=polos salientes)', '—', '—'),
    ('OEL_FixFactor', 'SYS_LOCAL'):('Factor corrección base OEL', '—', '—'),
    ('MEL_FixFactor', 'SYS_LOCAL'):('Factor corrección base MEL', '—', '—'),
    ('En_Sys', 'SYS_LOCAL'):       ('Habilitación funciones sistema', '—', '—'),

    # ── RTX_AI ──
    ('Ifd_Base_A', 'RTX_AI'):     ('Base corriente campo (A)', '—', 'Base conversión pu'),
    ('Efd_Base_V', 'RTX_AI'):     ('Base tensión campo (V)', '—', 'Base conversión pu'),
    ('IbridgeBase_A', 'RTX_AI'):  ('Base corriente puente (A)', '—', 'Base conversión pu'),
    ('Ifd_Tau', 'RTX_AI'):        ('Cte. filtro medición Ifd (s)', 'Tm / TIF (OEL/MEL)', 'Tabla 96/98'),
    ('Efd_Tau', 'RTX_AI'):        ('Cte. filtro medición Efd (s)', '—', '—'),

    # ── RTX_CTR ──
    ('Max_Ref_Vt', 'RTX_CTR'):    ('Referencia máxima tensión terminal (pu)', 'Lim_Ref_Max', 'Tabla 89'),
    ('Min_Ref_Vt', 'RTX_CTR'):    ('Referencia mínima tensión terminal (pu)', 'Lim_Ref_Min', 'Tabla 89'),
    ('Rate_Ref_Vt', 'RTX_CTR'):   ('Tasa rampa referencia Vt (%/s)', '—', '—'),
    ('alpha_min', 'RTX_CTR'):     ('Ángulo disparo mínimo rectificador (°)', '—', '—'),
    ('alpha_max', 'RTX_CTR'):     ('Ángulo disparo máximo rectificador (°)', '—', '—'),
    ('PWM_max', 'RTX_CTR'):       ('Máximo señal PWM (%) rectificador', '—', '—'),
    ('VtNominal', 'RTX_CTR'):     ('Tensión terminal nominal (pu)', '—', '—'),
    ('IfdNominal', 'RTX_CTR'):    ('Corriente campo nominal (pu RTVX)', '—', '—'),
    ('Kp_Q', 'RTX_CTR'):          ('Ganancia proporcional lazo Q', '—', '—'),
    ('Ki_Q', 'RTX_CTR'):          ('Ganancia integral lazo Q', '—', '—'),

    # ── RTX_VHZ ──
    ('Ref_VHZ', 'RTX_VHZ'):       ('VHZ - límite V/Hz (pu/pu)', 'RefVHZ', 'Tabla 93'),
    ('Kp_VHZ', 'RTX_VHZ'):        ('VHZ - ganancia proporcional', 'Kp', 'Tabla 93'),
    ('Ten_VHZ', 'RTX_VHZ'):       ('VHZ - tiempo retardo activación (s)', 'TEN', 'Tabla 93'),
    ('VHZHoff', 'RTX_VHZ'):       ('VHZ - histéresis desactivación', 'VHZOFF', 'Tabla 93'),
    ('Toff', 'RTX_VHZ'):          ('VHZ - tiempo desactivación (s)', 'TOFF', 'Tabla 93'),

    # ── RTX_UEL ──
    ('Kp_UEL', 'RTX_UEL'):        ('UEL - ganancia proporcional', 'Kp', 'Tabla 94'),
    ('TauPe_UEL', 'RTX_UEL'):     ('UEL - cte. filtro potencia activa (s)', 'TFilt_P', 'Tabla 94'),
    ('TauQ_UEL', 'RTX_UEL'):      ('UEL - cte. filtro potencia reactiva (s)', 'TFilt_Q', 'Tabla 94'),
    ('QHoff', 'RTX_UEL'):         ('UEL - histéresis desactivación', 'QHOFF', 'Tabla 94'),
    ('Toff', 'RTX_UEL'):          ('UEL - tiempo desactivación (s)', 'TOFF', 'Tabla 94'),
    ('TenUEL', 'RTX_UEL'):        ('UEL - tiempo retardo activación (s)', 'TEN', 'Tabla 94'),

    # ── RTX_LOCAL (curva UEL) ──
    ('k1_UEL', 'RTX_LOCAL'):       ('UEL - exponente influencia Vt sobre P', '—', '—'),
    ('k2_UEL', 'RTX_LOCAL'):       ('UEL - exponente influencia Vt sobre Q', '—', '—'),
    ('UEL_P_0', 'RTX_LOCAL'):      ('UEL - P punto 0 curva (pu)', 'Array UEL P(0)', 'Tabla 95'),
    ('UEL_P_1', 'RTX_LOCAL'):      ('UEL - P punto 1 curva (pu)', 'Array UEL P(1)', 'Tabla 95'),
    ('UEL_P_2', 'RTX_LOCAL'):      ('UEL - P punto 2 curva (pu)', 'Array UEL P(2)', 'Tabla 95'),
    ('UEL_Q_0', 'RTX_LOCAL'):      ('UEL - Q punto 0 curva (pu)', 'Array UEL Q(0)', 'Tabla 95'),
    ('UEL_Q_1', 'RTX_LOCAL'):      ('UEL - Q punto 1 curva (pu)', 'Array UEL Q(1)', 'Tabla 95'),
    ('UEL_Q_2', 'RTX_LOCAL'):      ('UEL - Q punto 2 curva (pu)', 'Array UEL Q(2)', 'Tabla 95'),
    ('UEL_n', 'RTX_LOCAL'):        ('UEL - número de puntos en curva', '—', '—'),

    # ── RTX_OEL ──
    ('Itherm_OEL', 'RTX_OEL'):    ('OEL - Ifd límite operación contínua (pu RTVX)', 'REF_IFD_TERM', 'Tabla 96'),
    ('Iinst_OEL', 'RTX_OEL'):     ('OEL - Ifd límite instantáneo (pu RTVX)', 'REF_IFD_PICO', 'Tabla 96'),
    ('tref_Iinst_OEL', 'RTX_OEL'):('OEL - tiempo máx. permanencia Iinst (s)', 'TIEMPO_DEFINIDO', 'Tabla 96'),
    ('TD_OEL', 'RTX_OEL'):        ('OEL - dial de tiempo curva térmica (s)', 'DT', 'Tabla 96'),
    ('Curve_OEL', 'RTX_OEL'):     ('OEL - tipo curva (0=tiempo definido)', 'HAB_CURVA', 'Tabla 96'),
    ('Kp_OEL', 'RTX_OEL'):        ('OEL - ganancia proporcional', 'Kp', 'Tabla 96'),
    ('Ten_OEL', 'RTX_OEL'):       ('OEL - tiempo retardo activación (s)', 'TEN', 'Tabla 96'),
    ('ITHoff_OEL', 'RTX_OEL'):    ('OEL - histéresis desactivación', 'IFHOFF', 'Tabla 96'),
    ('Toff_OEL', 'RTX_OEL'):      ('OEL - tiempo desactivación (s)', 'TOFF', 'Tabla 96'),
    ('Ta_OEL', 'RTX_OEL'):        ('OEL - cte. tiempo antecipación (s)', '—', '—'),

    # ── RTX_SCL ──
    ('Itherm_SCL', 'RTX_SCL'):    ('SCL - Ist límite operación contínua (pu)', 'REF_SCL_TERM', 'Tabla 97'),
    ('TD_SCL', 'RTX_SCL'):        ('SCL - dial de tiempo curva térmica (s)', 'TEMPO_DEFINIDO', 'Tabla 97'),
    ('Curve_SCL', 'RTX_SCL'):     ('SCL - tipo curva (0=tiempo definido)', 'HAB_CURVA', 'Tabla 97'),
    ('Kp_SCL', 'RTX_SCL'):        ('SCL - ganancia proporcional', 'Kp', 'Tabla 97'),
    ('Vtmin_SCL', 'RTX_SCL'):     ('SCL - tensión terminal mín. sobreexcitación (pu)', 'VT_MIN', 'Tabla 97'),
    ('Ixmin_SCL', 'RTX_SCL'):     ('SCL - corriente reactiva mín. activación (pu)', 'IX_MIN', 'Tabla 97'),
    ('TenOEL', 'RTX_SCL'):        ('SCL - tiempo habilitación sobreexcitación (s)', 'TEN_O', 'Tabla 97'),
    ('TenUEL', 'RTX_SCL'):        ('SCL - tiempo habilitación subexcitación (s)', 'TEN_U', 'Tabla 97'),
    ('ITHoff_SCL', 'RTX_SCL'):    ('SCL - histéresis desactivación', 'IXHOFF', 'Tabla 97'),
    ('Toff_SCL', 'RTX_SCL'):      ('SCL - tiempo desactivación (s)', 'TOFF', 'Tabla 97'),
    ('TVt_SCL', 'RTX_SCL'):       ('SCL - cte. filtro tensión terminal (s)', 'T_VT', 'Tabla 97'),
    ('It_inst_SCL', 'RTX_SCL'):   ('SCL - Ist límite instantáneo (pu)', 'REF_SCL_PICO', 'Tabla 97'),

    # ── RTX_MEL ──
    ('Ref_MEL', 'RTX_MEL'):       ('MEL - límite mín. corriente campo (pu RTVX)', 'REF_MEL', 'Tabla 98'),
    ('Kp_MEL', 'RTX_MEL'):        ('MEL - ganancia proporcional', 'Kp', 'Tabla 98'),
    ('ITHoff_MEL', 'RTX_MEL'):    ('MEL - histéresis desactivación', 'IFHOFF', 'Tabla 98'),
    ('Toff_MEL', 'RTX_MEL'):      ('MEL - tiempo desactivación (s)', 'TOFF', 'Tabla 98'),
    ('Ten_MEL', 'RTX_MEL'):       ('MEL - tiempo retardo activación (s)', 'TEN', 'Tabla 98'),
    ('Ten_MEL_52', 'RTX_MEL'):    ('MEL - retardo activación post-sincronización (s)', '—', '—'),

    # ── RTX_PRESET ──
    ('Tldo_PRESET', 'RTX_PRESET'):('T\'do campo abierto eje d (s)', 'Td0p (GENROU)', 'Generador'),
    ('Te_PRESET', 'RTX_PRESET'):  ('Cte. tiempo excitación para preset (s)', '—', '—'),

    # ── RVX_CTR — Gobernador ──
    ('bp_INTER', 'RVX_CTR'):      ('Estatismo permanente interconectado (pu/pu)', 'bp', 'Tabla 81'),
    ('bp_ISOL', 'RVX_CTR'):       ('Estatismo permanente aislado (pu/pu)', '—', '—'),
    ('bt_online', 'RVX_CTR'):     ('Banda proporcional en carga (pu/pu)', 'bt', 'Tabla 81'),
    ('bt_offline', 'RVX_CTR'):    ('Banda proporcional a vacío (pu/pu)', '—', '—'),
    ('Td_online', 'RVX_CTR'):     ('Tiempo integral en carga (s)', 'Td', 'Tabla 81'),
    ('Td_offline', 'RVX_CTR'):    ('Tiempo integral a vacío (s)', '—', '—'),
    ('Kw', 'RVX_CTR'):            ('Ganancia emulador servomotor (1/s)', 'Kw', 'Tabla 81'),
    ('aDB_w_INTER', 'RVX_CTR'):   ('Banda muerta frecuencia interconectado (pu)', 'Bmuerta', 'Tabla 81'),
    ('aDB_w_ISOL', 'RVX_CTR'):    ('Banda muerta frecuencia aislado (pu)', '—', '—'),
    ('Tg', 'RVX_CTR'):            ('Cte. filtro referencia Pe/Y (s)', '—', '—'),
    ('Tf', 'RVX_CTR'):            ('Cte. filtro realimentación Pe/Y (s)', '—', '—'),
    ('WNominal', 'RVX_CTR'):      ('Velocidad nominal (pu)', '—', '—'),

    # ── RVX_CTR_PELTON — Posición Pelton ──
    ('Kp_Ynz1', 'RVX_CTR_PELTON'):('Aguja 1 - ganancia proporcional', 'KP1 (Kp_Ynz1)', 'Tabla 84'),
    ('Ki_Ynz1', 'RVX_CTR_PELTON'):('Aguja 1 - ganancia integral (1/s)', '—', '—'),
    ('Kp_Yde1', 'RVX_CTR_PELTON'):('Deflector 1 - ganancia proporcional', 'KP2 (Kp_Yde1)', 'Tabla 84'),
    ('Ki_Yde1', 'RVX_CTR_PELTON'):('Deflector 1 - ganancia integral (1/s)', '—', '—'),
    ('FDither', 'RVX_CTR_PELTON'):('Frecuencia señal dither (Hz)', '—', '—'),
    ('MaxInteg_Ynz1', 'RVX_CTR_PELTON'):('Aguja 1 - máximo integrador (pu)', '—', '—'),
    ('MinInteg_Ynz1', 'RVX_CTR_PELTON'):('Aguja 1 - mínimo integrador (pu)', '—', '—'),
    ('MaxInteg_Yde1', 'RVX_CTR_PELTON'):('Deflector 1 - máximo integrador (pu)', '—', '—'),
    ('MinInteg_Yde1', 'RVX_CTR_PELTON'):('Deflector 1 - mínimo integrador (pu)', '—', '—'),
    ('MaxCtr_Ynz1', 'RVX_CTR_PELTON'): ('Aguja 1 - señal control máxima (pu)', '—', '—'),
    ('MinCtr_Ynz1', 'RVX_CTR_PELTON'): ('Aguja 1 - señal control mínima (pu)', '—', '—'),
    ('MaxCtr_Yde1', 'RVX_CTR_PELTON'): ('Deflector 1 - señal control máxima (pu)', '—', '—'),
    ('MinCtr_Yde1', 'RVX_CTR_PELTON'): ('Deflector 1 - señal control mínima (pu)', '—', '—'),

    # ── RVX_Pe_x_Y — Curva Pe vs Y ──
    ('P_loss', 'RVX_Pe_x_Y'):     ('Pérdidas en vacío (pu)', '—', '—'),
    ('NumCurv', 'RVX_Pe_x_Y'):    ('Número de curvas Pe-Y activas', '—', '—'),
}

# Categorías por módulo
MODULE_CATEGORY = {
    'SYS_LOCAL':         'Datos generador / Sistema',
    'RTX_AI':            'Mediciones excitación',
    'RTX_CTR':           'Control tensión (AVR)',
    'RTX_FAIL':          'Fallas excitación',
    'RTX_FREQ_RESP':     'Respuesta en frecuencia AVR',
    'RTX_LOCAL':         'Configuración AVR / Curva UEL',
    'RTX_LOG':           'Lógica relés excitación',
    'RTX_MEL':           'Limitador MEL',
    'RTX_OEL':           'Limitador OEL',
    'RTX_PRESET':        'Preset excitación',
    'RTX_SCL':           'Limitador SCL',
    'RTX_UEL':           'Limitador UEL',
    'RTX_VHZ':           'Limitador V/Hz',
    'RVX_AI':            'Mediciones hidráulicas',
    'RVX_AI_PELTON':     'Mediciones Pelton (agujas/deflectores)',
    'RVX_CTR':           'Control velocidad (gobernador)',
    'RVX_CTR_PELTON':    'Control posición Pelton',
    'RVX_CTR_PELTON_T2': 'Control número deflectores',
    'RVX_FAIL':          'Fallas gobernador',
    'RVX_FAIL_PELTON':   'Fallas Pelton',
    'RVX_LOCAL':         'Configuración gobernador',
    'RVX_LOG':           'Lógica relés gobernador',
    'RVX_LOG_PELTON':    'Lógica Pelton',
    'RVX_Pe_x_Y':        'Curvas Pe vs Y',
    'RVX_REG':           'Registrador gobernador',
}


# ═══════════════════════════════════════════════════════════════════════
# 4. ENRIQUECIMIENTO DE DATOS
# ═══════════════════════════════════════════════════════════════════════

def enrich(entries: list[VarEntry]) -> list[VarEntry]:
    """Agrega descripción, categoría y equivalencia DSL."""
    for e in entries:
        key = (e.name, e.module)
        if key in PARAM_INFO:
            desc, dsl, tabla = PARAM_INFO[key]
            e.description = desc
            e.dsl_equiv = f"{dsl} ({tabla})" if dsl != '—' else '—'
        else:
            e.description = ""
            e.dsl_equiv = ""

        e.category = MODULE_CATEGORY.get(e.module, e.module)

    return entries


# ═══════════════════════════════════════════════════════════════════════
# 5. FILTRADO
# ═══════════════════════════════════════════════════════════════════════

def filter_regulacion(entries: list[VarEntry]) -> list[VarEntry]:
    """Retorna solo los parámetros de módulos de regulación."""
    return [e for e in entries if e.module in MODULOS_REGULACION]


def filter_modulos(entries: list[VarEntry], modulos: list[str]) -> list[VarEntry]:
    """Filtra por lista específica de módulos."""
    mod_set = set(m.upper() for m in modulos)
    return [e for e in entries if e.module.upper() in mod_set]


# ═══════════════════════════════════════════════════════════════════════
# 6. EXPORTACIÓN
# ═══════════════════════════════════════════════════════════════════════

def export_csv(entries: list[VarEntry], output: Path, sep: str = ';',
               include_desc: bool = True):
    """Exporta a CSV compatible con Excel (UTF-8 BOM)."""
    header = ['Var#', 'Nombre', 'Modulo', 'Categoria', 'Tipo', 'Valor']
    if include_desc:
        header += ['Descripcion', 'Equivalencia DSL']

    with open(output, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=sep)
        writer.writerow(header)
        for e in entries:
            row = [e.var_num, e.full_name, e.module, e.category,
                   e.type_str, e.value_fmt()]
            if include_desc:
                row += [e.description, e.dsl_equiv]
            writer.writerow(row)

    print(f"  CSV guardado: {output}  ({len(entries)} parámetros)")


def export_json(entries: list[VarEntry], output: Path):
    """Exporta a JSON."""
    data = []
    for e in entries:
        data.append({
            'var': e.var_num,
            'name': e.full_name,
            'module': e.module,
            'category': e.category,
            'type': e.type_str,
            'value': e.value,
            'description': e.description,
            'dsl_equiv': e.dsl_equiv,
        })

    with open(output, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"  JSON guardado: {output}  ({len(entries)} parámetros)")


# ═══════════════════════════════════════════════════════════════════════
# 7. RESUMEN EN CONSOLA
# ═══════════════════════════════════════════════════════════════════════

def print_summary(entries: list[VarEntry], all_entries: list[VarEntry]):
    """Imprime resumen por módulo."""
    print(f"\n{'═' * 65}")
    print(f"  VARSPRS.DAT — Resumen de extracción")
    print(f"{'═' * 65}")
    print(f"  Total variables en archivo:    {len(all_entries)}")
    print(f"  Variables extraídas:           {len(entries)}")

    # Count by module
    mod_counts: dict[str, int] = {}
    for e in entries:
        mod_counts[e.module] = mod_counts.get(e.module, 0) + 1

    print(f"  Módulos:                       {len(mod_counts)}")
    print(f"{'─' * 65}")

    for mod in sorted(mod_counts.keys()):
        cat = MODULE_CATEGORY.get(mod, mod)
        print(f"    {mod:30s} {mod_counts[mod]:4d}  ({cat})")

    # Count DSL equivalences
    with_dsl = sum(1 for e in entries if e.dsl_equiv and e.dsl_equiv != '—')
    print(f"{'─' * 65}")
    print(f"  Con equivalencia DSL:          {with_dsl}")
    print(f"  Sin equivalencia DSL:          {len(entries) - with_dsl}")
    print(f"{'═' * 65}\n")


# ═══════════════════════════════════════════════════════════════════════
# 8. MAIN
# ═══════════════════════════════════════════════════════════════════════

BASE_DIR = Path(r"C:\Users\jose.lozano\Downloads\05_PARAMETROS REGULADORES")


def main():
    parser = argparse.ArgumentParser(
        description='Extrae parámetros de regulación de archivos VARSPRS.DAT (RTVX Power)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python extraer_VARSPRS.py --unidad BOT03
  python extraer_VARSPRS.py --unidad BOT03 --modulos RTX_OEL RTX_MEL
  python extraer_VARSPRS.py VARSPRS.DAT
  python extraer_VARSPRS.py VARSPRS.DAT -o resultado.csv
  python extraer_VARSPRS.py VARSPRS.DAT --todos
  python extraer_VARSPRS.py VARSPRS.DAT --modulos RTX_OEL RTX_MEL RTX_SCL
  python extraer_VARSPRS.py VARSPRS.DAT --json
        """
    )
    parser.add_argument('archivo', nargs='?', help='Ruta al archivo VARSPRS.DAT (opcional si se usa --unidad)')
    parser.add_argument('--unidad', help='Nombre de la unidad (ej: BOT03) — construye rutas automáticamente')
    parser.add_argument('-o', '--output', help='Nombre del archivo de salida')
    parser.add_argument('--todos', action='store_true',
                        help='Exportar TODOS los parámetros (no solo regulación)')
    parser.add_argument('--modulos', nargs='+',
                        help='Lista de módulos específicos a extraer')
    parser.add_argument('--separador', default=';',
                        help='Separador CSV (default: ;)')
    parser.add_argument('--sin-descripcion', action='store_true',
                        help='Omitir columnas de descripción y equivalencia DSL')
    parser.add_argument('--json', action='store_true',
                        help='Salida en formato JSON')

    args = parser.parse_args()

    # Resolver ruta del archivo
    if args.unidad:
        filepath = BASE_DIR / args.unidad / "VARSPRS.DAT"
    elif args.archivo:
        filepath = Path(args.archivo)
    else:
        parser.error("Se requiere 'archivo' o '--unidad NOMBRE_UNIDAD'")

    print(f"\n  Leyendo: {filepath}")
    all_entries = parse_dat(filepath)
    print(f"  Variables parseadas: {len(all_entries)}")

    # Enrich
    all_entries = enrich(all_entries)

    # Filter
    if args.modulos:
        entries = filter_modulos(all_entries, args.modulos)
    elif args.todos:
        entries = all_entries
    else:
        entries = filter_regulacion(all_entries)

    # Sort by module then var number
    entries.sort(key=lambda e: (e.module, e.var_num))

    # Summary
    print_summary(entries, all_entries)

    # Output path
    if args.output:
        outpath = Path(args.output)
    elif args.unidad:
        suffix = '.json' if args.json else '.csv'
        outpath = filepath.parent / f"{args.unidad}_parametros{suffix}"
    else:
        suffix = '.json' if args.json else '.csv'
        outpath = filepath.with_name(filepath.stem + '_parametros' + suffix)

    # Export
    if args.json:
        export_json(entries, outpath)
    else:
        export_csv(entries, outpath, sep=args.separador,
                   include_desc=not args.sin_descripcion)

    print(f"  Listo.\n")


if __name__ == '__main__':
    main()
