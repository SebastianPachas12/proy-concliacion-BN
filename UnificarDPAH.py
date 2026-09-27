import pandas as pd
from tkinter import Tk, filedialog
import os
import math

MAX_EXCEL_ROWS = 1_048_576
DATA_LIMIT = MAX_EXCEL_ROWS - 1

HOJAS_A_PROCESAR = [
    'SAL 762', 'SAL 772', 'SAL 512',
    'ENT 721', 'DPAH GEN',
    'COMIS 562', 'COMIS 512'
]

# ===============================
# SELECCIONAR ARCHIVOS (robusto)
# ===============================
def seleccionar_archivos():
    root = Tk()
    root.withdraw()

    rutas = filedialog.askopenfilenames(
        title="Seleccionar archivos Excel",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )

    root.destroy()
    return rutas


# ===============================
# LEER Y UNIR HOJAS
# ===============================
def leer_y_unir(archivos):
    datos = {}

    for archivo in archivos:

        if not os.path.exists(archivo):
            raise FileNotFoundError(f"No existe:\n{archivo}")

        xls = pd.ExcelFile(archivo)

        for hoja in HOJAS_A_PROCESAR:
            if hoja not in xls.sheet_names:
                continue

            df = pd.read_excel(
                archivo,
                sheet_name=hoja,
                dtype=str,
                engine="openpyxl"
            )

            datos.setdefault(hoja, []).append(df)

    return {
        hoja: pd.concat(dfs, ignore_index=True)
        for hoja, dfs in datos.items()
    }


# ===============================
# ORDENAR DATAFRAME
# ===============================
def ordenar_df(df):
    for col in ['Nro Cta', 'FECHA', 'HORA']:
        if col not in df.columns:
            raise ValueError(f"Falta la columna {col}")

    df['Nro Cta'] = pd.to_numeric(df['Nro Cta'], errors='coerce').fillna(0)
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['HORA'] = df['HORA'].astype(str).str.zfill(6)

    return df.sort_values(['Nro Cta', 'FECHA', 'HORA'])


# ===============================
# ESCRIBIR CON SPLIT AUTOMÁTICO
# ===============================
def escribir_con_split(writer, df, nombre_base):
    total_filas = len(df)
    partes = math.ceil(total_filas / DATA_LIMIT)

    for i in range(partes):
        inicio = i * DATA_LIMIT
        fin = inicio + DATA_LIMIT
        df_chunk = df.iloc[inicio:fin]

        nombre_hoja = nombre_base if i == 0 else f"{nombre_base}-{i+1}"

        df_chunk.to_excel(writer, sheet_name=nombre_hoja, index=False)


# ===============================
# CREAR ARCHIVO FINAL
# ===============================
def crear_archivo_final(dataframes, directorio):

    os.makedirs(directorio, exist_ok=True)

    salida = os.path.join(directorio, "DPAH UNIFICADO.xlsx")

    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        for hoja, df in dataframes.items():
            if df.empty:
                continue

            df_ordenado = ordenar_df(df)
            escribir_con_split(writer, df_ordenado, hoja)

    return salida


# ===============================
# EJECUTAR (para el main GUI)
# ===============================
def ejecutar(archivos):

    if not archivos:
        raise ValueError("No se seleccionaron archivos Excel.")

    directorio = os.path.dirname(archivos[0])

    dataframes = leer_y_unir(archivos)
    salida = crear_archivo_final(dataframes, directorio)

    return f"Archivo generado correctamente:\n{salida}"


# ===============================
# TEST LOCAL
# ===============================
if __name__ == "__main__":
    archivos = seleccionar_archivos()
    print(ejecutar(archivos))
