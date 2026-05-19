import pdfplumber
import pandas as pd
from pathlib import Path
from itertools import groupby

# ============================
# CONFIGURACIÓN
# ============================

PDF_PATH = r"C:\Datos Cobee\03_DATOS GEN\03_BOT\BOT03\02_MODELADO\TABLAS COMPOSITE MODEL.pdf"
OUTPUT_EXCEL = Path("salida_tablas_corregidas.xlsx")

# ============================
# UTILIDADES
# ============================

def limpiar_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.replace("", pd.NA, inplace=True)
    df.dropna(how="all", inplace=True)

    # Normalizar columnas
    df.columns = [str(c).strip() for c in df.columns]

    return df.reset_index(drop=True)


def reconstruir_tabla_desde_palabras(words, tolerancia_y=3):
    """
    Reconstruye una tabla agrupando palabras por coordenada vertical (y)
    """
    filas = []
    words = sorted(words, key=lambda w: (round(w["top"]), w["x0"]))

    for _, grupo in groupby(words, key=lambda w: round(w["top"] / tolerancia_y)):
        fila = [w["text"] for w in sorted(grupo, key=lambda w: w["x0"])]
        filas.append(fila)

    return filas


# ============================
# EXTRACCIÓN PRINCIPAL
# ============================

def extraer_tablas(pdf_path: Path) -> pd.DataFrame:
    dataframes = []

    with pdfplumber.open(pdf_path) as pdf:
        for num_pagina, pagina in enumerate(pdf.pages, start=1):

            # --- 1️⃣ INTENTO: tablas con líneas
            tablas = pagina.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 3,
                "snap_tolerance": 3,
                "join_tolerance": 3
            })

            for t_idx, tabla in enumerate(tablas, start=1):
                df = pd.DataFrame(tabla)
                if df.shape[1] < 2:
                    continue

                # Encabezados
                df.columns = df.iloc[0]
                df = df.iloc[1:]

                df["pagina"] = num_pagina
                df["origen"] = "lines"
                df["tabla"] = t_idx

                dataframes.append(limpiar_df(df))

            # --- 2️⃣ INTENTO: sin líneas (texto alineado)
            tablas_stream = pagina.extract_tables({
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "intersection_tolerance": 3,
            })

            for t_idx, tabla in enumerate(tablas_stream, start=1):
                df = pd.DataFrame(tabla)
                if df.shape[1] < 2:
                    continue

                df.columns = df.iloc[0]
                df = df.iloc[1:]

                df["pagina"] = num_pagina
                df["origen"] = "text"
                df["tabla"] = t_idx

                dataframes.append(limpiar_df(df))

            # --- 3️⃣ FALLBACK: reconstrucción manual
            words = pagina.extract_words(use_text_flow=True)
            if len(words) > 20:
                filas = reconstruir_tabla_desde_palabras(words)
                if len(filas) > 2:
                    df = pd.DataFrame(filas)
                    df["pagina"] = num_pagina
                    df["origen"] = "reconstruida"
                    df["tabla"] = 1
                    dataframes.append(limpiar_df(df))

    if not dataframes:
        raise RuntimeError("No se logró extraer información útil del PDF")

    return pd.concat(dataframes, ignore_index=True)


# ============================
# EXPORTACIÓN
# ============================

def exportar_excel(df: pd.DataFrame, output: Path):
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="DATA", index=False)

    print(f"✅ Archivo generado: {output.resolve()}")


# ============================
# MAIN
# ============================

if __name__ == "__main__":
    df_final = extraer_tablas(PDF_PATH)
    exportar_excel(df_final, OUTPUT_EXCEL)