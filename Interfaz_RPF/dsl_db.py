"""
dsl_db.py
---------
Almacenamiento local de experimentos DSL en archivos CSV.
Reemplaza la capa PostgreSQL anterior; mantiene exactamente la misma API pública
para que bloque_dsl_params.py no necesite ningún cambio.

Archivos en C:\\Datos Cobee\\03_DATOS GEN\\dsl_local\\:
  experimentos.csv  — metadatos de cada experimento
  exp_params.csv    — parámetros DSL por experimento
  exp_kpis.csv      — KPIs por experimento
"""

from __future__ import annotations
import os
import pandas as pd
from datetime import datetime
from pathlib import Path

# ── Ubicación de almacenamiento local ─────────────────────────────────────────
_LOCAL_DIR = Path(r"C:\Datos Cobee\03_DATOS GEN\dsl_local")

_EXP_CSV    = _LOCAL_DIR / "experimentos.csv"
_PARAMS_CSV = _LOCAL_DIR / "exp_params.csv"
_KPIS_CSV   = _LOCAL_DIR / "exp_kpis.csv"

# ── Esquemas de columnas ──────────────────────────────────────────────────────
_EXP_COLS = [
    "id", "sym", "familia", "evento_ref", "nombre",
    "fecha", "notas", "estado", "excel_path", "carpeta_curvas",
]
_PARAMS_COLS = [
    "id", "exp_id", "bloque", "simbolo", "descripcion",
    "valor_base", "valor", "es_ajustable",
]
_KPIS_COLS = [
    "id", "exp_id",
    "f0", "f_min", "t_min", "delta_f", "f_delta_t",
    "p0", "p_max", "p_delta_t", "rocof",
    "delta_p", "delta_p_pct", "aporta_rpf",
]


# ── Helpers de lectura/escritura ───────────────────────────────────────────────

def _leer_csv(path: Path, columns: list[str]) -> pd.DataFrame:
    if path.is_file():
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
            if "id" in df.columns:
                df["id"] = pd.to_numeric(df["id"], errors="coerce").fillna(0).astype(int)
            if "exp_id" in df.columns:
                df["exp_id"] = pd.to_numeric(df["exp_id"], errors="coerce").fillna(0).astype(int)
            return df
        except Exception:
            return pd.DataFrame(columns=columns)
    return pd.DataFrame(columns=columns)


def _guardar_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def _next_id(df: pd.DataFrame) -> int:
    if not df.empty and "id" in df.columns:
        ids = pd.to_numeric(df["id"], errors="coerce").dropna()
        if not ids.empty:
            return int(ids.max()) + 1
    return 1


# ── API pública ───────────────────────────────────────────────────────────────

def probar_conexion(**_) -> tuple[bool, str]:
    """Verifica que el directorio local sea accesible."""
    try:
        _LOCAL_DIR.mkdir(parents=True, exist_ok=True)
        return True, f"Almacenamiento local activo: {_LOCAL_DIR}"
    except Exception as e:
        return False, str(e)


def crear_tablas(**_) -> None:
    """Crea los CSVs con cabeceras si no existen (idempotente)."""
    _LOCAL_DIR.mkdir(parents=True, exist_ok=True)
    for path, cols in [(_EXP_CSV, _EXP_COLS),
                       (_PARAMS_CSV, _PARAMS_COLS),
                       (_KPIS_CSV, _KPIS_COLS)]:
        if not path.is_file():
            pd.DataFrame(columns=cols).to_csv(path, index=False, encoding="utf-8-sig")


def registrar_experimento(sym: str, familia: str, evento_ref: str,
                          nombre: str, notas: str = "",
                          excel_path: str = "", **_) -> int:
    """Añade un nuevo experimento. Devuelve el id generado."""
    df = _leer_csv(_EXP_CSV, _EXP_COLS)
    new_id = _next_id(df)
    nuevo = pd.DataFrame([{
        "id":          new_id,
        "sym":         sym,
        "familia":     familia,
        "evento_ref":  evento_ref,
        "nombre":      nombre,
        "fecha":       datetime.now().isoformat(timespec="seconds"),
        "notas":       notas,
        "estado":      "configurado",
        "excel_path":  excel_path,
        "carpeta_curvas": "",
    }])
    _guardar_csv(pd.concat([df, nuevo], ignore_index=True), _EXP_CSV)
    return new_id


def actualizar_estado(exp_id: int, estado: str, **_) -> None:
    """Actualiza el campo estado de un experimento."""
    df = _leer_csv(_EXP_CSV, _EXP_COLS)
    mask = df["id"] == int(exp_id)
    if mask.any():
        df.loc[mask, "estado"] = estado
        _guardar_csv(df, _EXP_CSV)


def vincular_curva(exp_id: int, carpeta_curvas: str, **_) -> None:
    """Actualiza carpeta_curvas de un experimento."""
    df = _leer_csv(_EXP_CSV, _EXP_COLS)
    mask = df["id"] == int(exp_id)
    if mask.any():
        df.loc[mask, "carpeta_curvas"] = carpeta_curvas
        _guardar_csv(df, _EXP_CSV)


def registrar_params(exp_id: int, params: list[dict], **_) -> None:
    """
    Añade filas a exp_params.csv.
    params = [{"bloque": str, "simbolo": str, "descripcion": str,
               "valor_base": float, "valor": float, "es_ajustable": bool}, ...]
    """
    if not params:
        return
    df = _leer_csv(_PARAMS_CSV, _PARAMS_COLS)
    next_id = _next_id(df)
    nuevas = []
    for i, p in enumerate(params):
        nuevas.append({
            "id":          next_id + i,
            "exp_id":      int(exp_id),
            "bloque":      p.get("bloque", ""),
            "simbolo":     p["simbolo"],
            "descripcion": p.get("descripcion", ""),
            "valor_base":  p.get("valor_base"),
            "valor":       p["valor"],
            "es_ajustable": p.get("es_ajustable", True),
        })
    _guardar_csv(pd.concat([df, pd.DataFrame(nuevas)], ignore_index=True), _PARAMS_CSV)


def registrar_kpis(exp_id: int, kpis: dict, **_) -> None:
    """
    Añade (o reemplaza si ya existe) la fila de KPIs de un experimento.
    kpis = {"f0": float, "f_min": float, ...}
    """
    df = _leer_csv(_KPIS_CSV, _KPIS_COLS)
    df = df[df["exp_id"] != int(exp_id)]  # upsert: quitar fila anterior si existe
    next_id = _next_id(df)
    nueva = pd.DataFrame([{
        "id":          next_id,
        "exp_id":      int(exp_id),
        "f0":          kpis.get("f0"),
        "f_min":       kpis.get("f_min"),
        "t_min":       kpis.get("t_min"),
        "delta_f":     kpis.get("delta_f"),
        "f_delta_t":   kpis.get("f_delta_t"),
        "p0":          kpis.get("p0"),
        "p_max":       kpis.get("p_max"),
        "p_delta_t":   kpis.get("p_delta_t"),
        "rocof":       kpis.get("rocof"),
        "delta_p":     kpis.get("delta_p"),
        "delta_p_pct": kpis.get("delta_p_pct"),
        "aporta_rpf":  kpis.get("aporta_rpf"),
    }])
    _guardar_csv(pd.concat([df, nueva], ignore_index=True), _KPIS_CSV)
    actualizar_estado(exp_id, "simulado")


def listar_experimentos(sym: str | None = None,
                        familia: str | None = None,
                        estado: str | None = None,
                        desde: datetime | None = None,
                        hasta: datetime | None = None,
                        **_) -> pd.DataFrame:
    """
    Devuelve DataFrame de experimentos con KPIs JOIN.
    Columnas idénticas a la versión PostgreSQL.
    """
    df_e = _leer_csv(_EXP_CSV, _EXP_COLS)
    df_k = _leer_csv(_KPIS_CSV, _KPIS_COLS)

    if df_e.empty:
        return pd.DataFrame(columns=[
            "id", "sym", "familia", "evento_ref", "nombre",
            "fecha", "estado", "notas", "carpeta_curvas",
            "f_min", "t_min", "delta_f", "rocof",
            "delta_p", "delta_p_pct", "aporta_rpf",
        ])

    # Filtros
    if sym:
        df_e = df_e[df_e["sym"] == sym]
    if familia:
        df_e = df_e[df_e["familia"] == familia]
    if estado:
        df_e = df_e[df_e["estado"] == estado]
    if desde:
        df_e = df_e[pd.to_datetime(df_e["fecha"], errors="coerce") >= pd.Timestamp(desde)]
    if hasta:
        df_e = df_e[pd.to_datetime(df_e["fecha"], errors="coerce") <= pd.Timestamp(hasta)]

    # Join con KPIs
    kpi_cols = ["exp_id", "f_min", "t_min", "delta_f", "rocof",
                "delta_p", "delta_p_pct", "aporta_rpf"]
    df_k_sub = df_k[kpi_cols] if not df_k.empty else pd.DataFrame(columns=kpi_cols)

    result = df_e.merge(df_k_sub, left_on="id", right_on="exp_id", how="left")
    result = result.drop(columns=["exp_id"], errors="ignore")

    # Ordenar más reciente primero
    if "fecha" in result.columns:
        result["fecha"] = pd.to_datetime(result["fecha"], errors="coerce")
        result = result.sort_values("fecha", ascending=False)

    return result.reset_index(drop=True)


def params_de_experimento(exp_id: int, **_) -> pd.DataFrame:
    """Devuelve los parámetros de un experimento."""
    df = _leer_csv(_PARAMS_CSV, _PARAMS_COLS)
    if df.empty:
        return pd.DataFrame(columns=[
            "bloque", "simbolo", "descripcion", "valor_base", "valor", "es_ajustable"])
    result = df[df["exp_id"] == int(exp_id)][
        ["bloque", "simbolo", "descripcion", "valor_base", "valor", "es_ajustable"]
    ].copy()
    return result.sort_values(["bloque", "simbolo"]).reset_index(drop=True)


def exportar_para_ia(sym: str | None = None, **_) -> pd.DataFrame:
    """
    Devuelve DataFrame wide con una fila por experimento:
      features = parámetros DSL, targets = KPIs.
    Listo para entrenamiento de modelos ML.
    """
    exps = listar_experimentos(sym=sym, estado="simulado")
    if exps.empty:
        return exps

    filas = []
    for _, row in exps.iterrows():
        params_df = params_de_experimento(int(row["id"]))
        fila = {
            "exp_id":      row["id"],
            "sym":         row["sym"],
            "familia":     row["familia"],
            "evento_ref":  row["evento_ref"],
            "fecha":       row["fecha"],
            "f_min":       row["f_min"],
            "t_min":       row["t_min"],
            "delta_f":     row["delta_f"],
            "rocof":       row["rocof"],
            "delta_p":     row["delta_p"],
            "delta_p_pct": row["delta_p_pct"],
            "aporta_rpf":  row["aporta_rpf"],
        }
        for _, p in params_df.iterrows():
            fila[p["simbolo"]] = p["valor"]
        filas.append(fila)

    return pd.DataFrame(filas)
