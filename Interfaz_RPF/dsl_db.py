"""
dsl_db.py
---------
Capa de acceso a PostgreSQL para el módulo de optimización de parámetros DSL.
Base de datos: rpf_intelligence (servidor 192.168.0.92)

Tablas nuevas:
  dsl_experimentos  — metadatos de cada experimento
  dsl_exp_params    — parámetros DSL usados en el experimento
  dsl_exp_kpis      — KPIs obtenidos de la simulación
"""

from __future__ import annotations
import os
import pandas as pd
from datetime import datetime

# ── Configuración de conexión ──────────────────────────────────────────────────
_PG_DEFAULTS = {
    "host":     os.getenv("PG_HOST",     "192.168.0.92"),
    "port":     int(os.getenv("PG_PORT", "5432")),
    "dbname":   os.getenv("PG_DB",       "rpf_intelligence"),
    "user":     os.getenv("PG_USER",     "jose"),
    "password": os.getenv("PG_PASSWORD", ""),
}

# ── DDL ───────────────────────────────────────────────────────────────────────
_DDL = """
CREATE TABLE IF NOT EXISTS dsl_experimentos (
    id          SERIAL PRIMARY KEY,
    sym         VARCHAR(10)  NOT NULL,
    familia     VARCHAR(5),
    evento_ref  VARCHAR(100),
    nombre      VARCHAR(200),
    fecha       TIMESTAMPTZ  DEFAULT NOW(),
    notas       TEXT,
    estado      VARCHAR(20)  DEFAULT 'configurado',
    excel_path  TEXT
);

CREATE TABLE IF NOT EXISTS dsl_exp_params (
    id           SERIAL PRIMARY KEY,
    exp_id       INT REFERENCES dsl_experimentos(id) ON DELETE CASCADE,
    bloque       VARCHAR(30),
    simbolo      VARCHAR(60)  NOT NULL,
    descripcion  TEXT,
    valor_base   FLOAT,
    valor        FLOAT        NOT NULL,
    es_ajustable BOOLEAN      DEFAULT TRUE
);

CREATE TABLE IF NOT EXISTS dsl_exp_kpis (
    id          SERIAL PRIMARY KEY,
    exp_id      INT REFERENCES dsl_experimentos(id) ON DELETE CASCADE,
    f0          FLOAT,
    f_min       FLOAT,
    t_min       FLOAT,
    delta_f     FLOAT,
    f_delta_t   FLOAT,
    p0          FLOAT,
    p_max       FLOAT,
    p_delta_t   FLOAT,
    rocof       FLOAT,
    delta_p     FLOAT,
    delta_p_pct FLOAT,
    aporta_rpf  BOOLEAN
);

CREATE INDEX IF NOT EXISTS idx_dsl_exp_sym   ON dsl_experimentos(sym);
CREATE INDEX IF NOT EXISTS idx_dsl_exp_fam   ON dsl_experimentos(familia);
CREATE INDEX IF NOT EXISTS idx_dsl_exp_fecha ON dsl_experimentos(fecha DESC);
CREATE INDEX IF NOT EXISTS idx_dsl_params_exp ON dsl_exp_params(exp_id);
CREATE INDEX IF NOT EXISTS idx_dsl_kpis_exp   ON dsl_exp_kpis(exp_id);
"""


# ── Conexión ──────────────────────────────────────────────────────────────────

def _conectar(**overrides):
    """Devuelve una conexión psycopg2. Lanza ImportError si no está instalado."""
    try:
        import psycopg2
    except ImportError as e:
        raise ImportError("Instalar psycopg2-binary: pip install psycopg2-binary") from e
    cfg = {**_PG_DEFAULTS, **overrides}
    return psycopg2.connect(**cfg)


def probar_conexion(**overrides) -> tuple[bool, str]:
    """Devuelve (ok, mensaje). Útil para el widget de estado en Streamlit."""
    try:
        conn = _conectar(**overrides)
        conn.close()
        return True, "Conectado a rpf_intelligence"
    except Exception as e:
        return False, str(e)


# ── Inicialización de tablas ──────────────────────────────────────────────────

def crear_tablas(**overrides):
    """Crea las tablas DSL si no existen (idempotente)."""
    conn = _conectar(**overrides)
    try:
        with conn.cursor() as cur:
            cur.execute(_DDL)
            cur.execute(
                "ALTER TABLE dsl_experimentos ADD COLUMN IF NOT EXISTS carpeta_curvas TEXT"
            )
        conn.commit()
    finally:
        conn.close()


# ── Experimentos ──────────────────────────────────────────────────────────────

def registrar_experimento(sym: str, familia: str, evento_ref: str,
                          nombre: str, notas: str = "",
                          excel_path: str = "", **overrides) -> int:
    """INSERT en dsl_experimentos. Devuelve el id generado."""
    conn = _conectar(**overrides)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """INSERT INTO dsl_experimentos
                   (sym, familia, evento_ref, nombre, notas, excel_path)
                   VALUES (%s,%s,%s,%s,%s,%s) RETURNING id""",
                (sym, familia, evento_ref, nombre, notas, excel_path),
            )
            exp_id = cur.fetchone()[0]
        conn.commit()
        return exp_id
    finally:
        conn.close()


def actualizar_estado(exp_id: int, estado: str, **overrides):
    """Actualiza estado: 'configurado' | 'simulado' | 'analizado'."""
    conn = _conectar(**overrides)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE dsl_experimentos SET estado=%s WHERE id=%s",
                (estado, exp_id),
            )
        conn.commit()
    finally:
        conn.close()


def vincular_curva(exp_id: int, carpeta_curvas: str, **overrides):
    """UPDATE dsl_experimentos.carpeta_curvas. No cambia el estado (eso lo hace registrar_kpis)."""
    conn = _conectar(**overrides)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE dsl_experimentos SET carpeta_curvas=%s WHERE id=%s",
                (carpeta_curvas, exp_id),
            )
        conn.commit()
    finally:
        conn.close()


# ── Parámetros ────────────────────────────────────────────────────────────────

def registrar_params(exp_id: int, params: list[dict], **overrides):
    """
    INSERT en dsl_exp_params.
    params = [{"bloque": str, "simbolo": str, "descripcion": str,
               "valor_base": float, "valor": float, "es_ajustable": bool}, ...]
    """
    if not params:
        return
    conn = _conectar(**overrides)
    try:
        with conn.cursor() as cur:
            cur.executemany(
                """INSERT INTO dsl_exp_params
                   (exp_id, bloque, simbolo, descripcion, valor_base, valor, es_ajustable)
                   VALUES (%s,%s,%s,%s,%s,%s,%s)""",
                [
                    (exp_id,
                     p.get("bloque", ""),
                     p["simbolo"],
                     p.get("descripcion", ""),
                     p.get("valor_base"),
                     p["valor"],
                     p.get("es_ajustable", True))
                    for p in params
                ],
            )
        conn.commit()
    finally:
        conn.close()


# ── KPIs ──────────────────────────────────────────────────────────────────────

def registrar_kpis(exp_id: int, kpis: dict, **overrides):
    """
    INSERT/UPDATE en dsl_exp_kpis.
    kpis = {"f0": float, "f_min": float, "t_min": float, "delta_f": float,
             "f_delta_t": float, "p0": float, "p_max": float, "p_delta_t": float,
             "rocof": float, "delta_p": float, "delta_p_pct": float,
             "aporta_rpf": bool}
    """
    conn = _conectar(**overrides)
    try:
        with conn.cursor() as cur:
            cur.execute(
                """INSERT INTO dsl_exp_kpis
                   (exp_id, f0, f_min, t_min, delta_f, f_delta_t,
                    p0, p_max, p_delta_t, rocof, delta_p, delta_p_pct, aporta_rpf)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                   ON CONFLICT DO NOTHING""",
                (exp_id,
                 kpis.get("f0"), kpis.get("f_min"), kpis.get("t_min"),
                 kpis.get("delta_f"), kpis.get("f_delta_t"),
                 kpis.get("p0"), kpis.get("p_max"), kpis.get("p_delta_t"),
                 kpis.get("rocof"), kpis.get("delta_p"), kpis.get("delta_p_pct"),
                 kpis.get("aporta_rpf")),
            )
        conn.commit()
        actualizar_estado(exp_id, "simulado", **overrides)
    finally:
        conn.close()


# ── Consultas ─────────────────────────────────────────────────────────────────

def listar_experimentos(sym: str | None = None,
                        familia: str | None = None,
                        estado: str | None = None,
                        desde: datetime | None = None,
                        hasta: datetime | None = None,
                        **overrides) -> pd.DataFrame:
    """
    SELECT de experimentos con KPIs JOIN.
    Devuelve DataFrame listo para mostrar en Streamlit.
    """
    where, args = [], []
    if sym:
        where.append("e.sym = %s");    args.append(sym)
    if familia:
        where.append("e.familia = %s"); args.append(familia)
    if estado:
        where.append("e.estado = %s");  args.append(estado)
    if desde:
        where.append("e.fecha >= %s");  args.append(desde)
    if hasta:
        where.append("e.fecha <= %s");  args.append(hasta)

    sql = """
        SELECT e.id, e.sym, e.familia, e.evento_ref, e.nombre,
               e.fecha, e.estado, e.notas, e.carpeta_curvas,
               k.f_min, k.t_min, k.delta_f, k.rocof,
               k.delta_p, k.delta_p_pct, k.aporta_rpf
        FROM   dsl_experimentos e
        LEFT JOIN dsl_exp_kpis k ON k.exp_id = e.id
    """
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY e.fecha DESC"

    conn = _conectar(**overrides)
    try:
        return pd.read_sql(sql, conn, params=args)
    finally:
        conn.close()


def params_de_experimento(exp_id: int, **overrides) -> pd.DataFrame:
    """Devuelve los parámetros de un experimento."""
    sql = """
        SELECT bloque, simbolo, descripcion, valor_base, valor, es_ajustable
        FROM   dsl_exp_params
        WHERE  exp_id = %s
        ORDER  BY bloque, simbolo
    """
    conn = _conectar(**overrides)
    try:
        return pd.read_sql(sql, conn, params=[exp_id])
    finally:
        conn.close()


def exportar_para_ia(sym: str | None = None, **overrides) -> pd.DataFrame:
    """
    Devuelve un DataFrame wide con una fila por experimento:
      features = parámetros DSL, targets = KPIs
    Listo para entrenamiento de modelos ML.
    """
    exps = listar_experimentos(sym=sym, estado="simulado", **overrides)
    if exps.empty:
        return exps

    filas = []
    for _, row in exps.iterrows():
        params_df = params_de_experimento(int(row["id"]), **overrides)
        fila = {
            "exp_id":       row["id"],
            "sym":          row["sym"],
            "familia":      row["familia"],
            "evento_ref":   row["evento_ref"],
            "fecha":        row["fecha"],
            "f_min":        row["f_min"],
            "t_min":        row["t_min"],
            "delta_f":      row["delta_f"],
            "rocof":        row["rocof"],
            "delta_p":      row["delta_p"],
            "delta_p_pct":  row["delta_p_pct"],
            "aporta_rpf":   row["aporta_rpf"],
        }
        for _, p in params_df.iterrows():
            fila[p["simbolo"]] = p["valor"]
        filas.append(fila)

    return pd.DataFrame(filas)
