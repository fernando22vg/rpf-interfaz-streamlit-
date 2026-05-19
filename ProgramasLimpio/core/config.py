# core/config.py
"""
PipelinePaths — configuración centralizada de rutas.
Carga desde ~/.cobee_pipeline.toml (se crea con save_toml()).
"""
from __future__ import annotations

import glob as _glob
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Optional

CONFIG_PATH = Path.home() / ".cobee_pipeline.toml"

CARPETAS_RESULTADOS = [
    "Resultados_COBEE",
    "Resultados_ENDE_Andina",
    "Resultados_GUABIRA",
    "Resultados_SCADA",
    "Resultados_ENDE_Corani",
    "Resultados_ENDE_Guaracachi",
    "Resultados_AGUAI",
    "Resultados_HB",
    "Resultados_ENDE_Valle_Hermoso",
]


@dataclass
class PipelinePaths:
    raiz_cndc: str = r"C:\Datos del CNDC\01_INFO CNDC_RPF"
    raiz_loc_names: str = (
        r"C:\Datos del CNDC\DATOS EXTRAIDOS DE DIGSILENT\Designacion de loc_name"
    )
    raiz_programas: str = ""
    pf_modelo: str = "PMP_NOV25_OCT29_31102025(1)"
    pf_base: str = r"C:\Program Files\DIgSILENT\PowerFactory 2025 SP2"
    pf_user: str = ""

    def __post_init__(self) -> None:
        if not self.raiz_programas:
            # Padre de core/
            self.raiz_programas = str(Path(__file__).parent.parent)

    # ── Rutas derivadas ───────────────────────────────────────────────────────
    @property
    def loc_gen(self) -> Path:
        return Path(self.raiz_loc_names) / "loc_names_gen.xlsx"

    @property
    def loc_cargas(self) -> Path:
        return Path(self.raiz_loc_names) / "loc_name_cargas.xlsx"

    @property
    def loc_xfo(self) -> Path:
        return Path(self.raiz_loc_names) / "loc_names_xfo.xlsx"

    @property
    def loc_lineas(self) -> Path:
        return Path(self.raiz_loc_names) / "loc_names_lineas.xlsx"

    @property
    def datos_sin(self) -> Path:
        return Path(self.raiz_programas) / "Programas_1_uso_modelo" / "DatosSINdigsilent.xlsx"

    def evento_path(self, semestre: str, evento_num: int) -> Path:
        return (
            Path(self.raiz_cndc)
            / semestre
            / "Análisis_todos_los_eventos"
            / f"Evento {evento_num}"
        )

    # ── Escaneo dinámico ──────────────────────────────────────────────────────
    def semestres(self) -> list[str]:
        raiz = Path(self.raiz_cndc)
        if not raiz.exists():
            return []
        return sorted(d.name for d in raiz.iterdir() if d.is_dir())

    def eventos_de_semestre(self, semestre: str) -> list[int]:
        ev_raiz = Path(self.raiz_cndc) / semestre / "Análisis_todos_los_eventos"
        if not ev_raiz.exists():
            return []
        nums: list[int] = []
        for d in ev_raiz.iterdir():
            if d.is_dir() and d.name.startswith("Evento "):
                try:
                    nums.append(int(d.name.split()[-1]))
                except ValueError:
                    pass
        return sorted(nums)

    def tabla_eventos_path(self, semestre: str) -> Optional[Path]:
        results = _glob.glob(
            str(Path(self.raiz_cndc) / semestre / "Tabla_Eventos_*.xlsx")
        )
        return Path(results[0]) if results else None

    def datos_simulacion_path(self, semestre: str, evento_num: int) -> Optional[Path]:
        ev = self.evento_path(semestre, evento_num)
        results = _glob.glob(str(ev / "datos_simulacion_*.xlsx"))
        return Path(results[0]) if results else None

    def condiciones_iniciales_path(self, semestre: str, evento_num: int) -> Optional[Path]:
        ev = self.evento_path(semestre, evento_num)
        results = _glob.glob(str(ev / "condiciones_iniciales_*.xlsx"))
        return Path(results[0]) if results else None

    # ── Estado del pipeline ───────────────────────────────────────────────────
    def fase_status(
        self,
        semestre: Optional[str] = None,
        evento_num: Optional[int] = None,
    ) -> dict[str, bool]:
        s0 = self.datos_sin.exists()
        s1 = all(
            [
                self.loc_gen.exists(),
                self.loc_xfo.exists(),
                self.loc_lineas.exists(),
                self.loc_cargas.exists(),
            ]
        )
        s2 = False
        s3 = False
        if semestre and evento_num:
            s2 = (
                self.datos_simulacion_path(semestre, evento_num) is not None
                and self.condiciones_iniciales_path(semestre, evento_num) is not None
            )
        return {"fase0": s0, "fase1": s1, "fase2": s2, "fase3": s3}

    # ── Persistencia TOML ─────────────────────────────────────────────────────
    @classmethod
    def from_toml(cls) -> "PipelinePaths":
        if not CONFIG_PATH.exists():
            return cls()
        try:
            import tomllib  # Python 3.11+
            with open(CONFIG_PATH, "rb") as f:
                data = tomllib.load(f)
            p = data.get("paths", {})
            obj = cls()
            for fld in ["raiz_cndc", "raiz_loc_names", "raiz_programas",
                        "pf_modelo", "pf_base", "pf_user"]:
                if fld in p:
                    setattr(obj, fld, p[fld])
            return obj
        except Exception:
            return cls()

    def save_toml(self) -> None:
        lines = ["[paths]\n"]
        for k, v in asdict(self).items():
            esc = str(v).replace("\\", "\\\\")
            lines.append(f'{k} = "{esc}"\n')
        CONFIG_PATH.write_text("".join(lines), encoding="utf-8")
