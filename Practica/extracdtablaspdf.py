import pdfplumber
import pandas as pd
from pathlib import Path

# ============================
# CONFIGURACIÓN
# ============================

PDF_PATH =r"C:\Datos Cobee\03_DATOS GEN\03_BOT\BOT03\02_MODELADO\TABLAS COMPOSITE MODEL.pdf"
OUTPUT_EXCEL = Path("salida_tablas.xlsx")

# ============================
# EXTRACCIÓN DE TABLAS
# ============================

def extraer_tablas(pdf_path: Path):
    resultados = []

    with pdfplumber.open(pdf_path) as pdf:
        for num_pagina, pagina in enumerate(pdf.pages, start=1):

            tablas = pagina.extract_tables(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                }
            )

            if not tablas:
                continue

            for idx, tabla in enumerate(tablas, start=1):
                df = pd.DataFrame(tabla)

                # Eliminar filas completamente vacías
                df.replace("", pd.NA, inplace=True)
                df.dropna(how="all", inplace=True)

                # Usar la primera fila como encabezado si es razonable
                if df.shape[0] > 1:
                    df.columns = df.iloc[0]
                    df = df.drop(index=0).reset_index(drop=True)

                nombre = f"Pag_{num_pagina}_Tabla_{idx}"
                resultados.append((nombre, df))

    return resultados


# ============================
# EXPORTACIÓN A EXCEL
# ============================

def exportar_excel(tablas, output_file: Path):
    if not tablas:
        print("⚠️ No se encontraron tablas en el PDF.")
        return

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for nombre, df in tablas:
            # Excel solo permite 31 caracteres por hoja
            hoja = nombre[:31]
            df.to_excel(writer, sheet_name=hoja, index=False)

    print(f"✅ Archivo generado: {output_file.resolve()}")


# ============================
# MAIN
# ============================

if __name__ == "__main__":
    tablas = extraer_tablas(PDF_PATH)
    exportar_excel(tablas, OUTPUT_EXCEL)