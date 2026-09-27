import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os


def seleccionar_archivos():
    root = Tk()
    root.withdraw()

    archivo_cce = askopenfilename(title="Selecciona el archivo CCE", filetypes=[("Excel", "*.xlsx")])
    archivo_dpah = askopenfilename(title="Selecciona el archivo DPAH", filetypes=[("Excel", "*.xlsx")])

    root.destroy()

    if not archivo_cce or not archivo_dpah:
        raise ValueError("Selección cancelada.")

    return archivo_cce, archivo_dpah


def leer_hoja(path, nombre):
    try:
        return pd.read_excel(path, sheet_name=nombre, dtype=str)
    except:
        return None


def _armar_salida(noA, noB):
    return pd.concat(
        [
            noA.reset_index(drop=True),
            pd.DataFrame({"DIFF": [""] * len(noA)}),
            noB.reset_index(drop=True),
            pd.DataFrame({"DIFF2": [""] * len(noB)})
        ],
        axis=1
    )


def cruzar_sal320_762(df_320_list, df_762_list):

    dfA = pd.concat([d for d in df_320_list if d is not None], ignore_index=True)
    dfB = pd.concat([d for d in df_762_list if d is not None], ignore_index=True)

    if dfA.empty or dfB.empty:
        return pd.DataFrame()

    dfA = dfA.copy()
    dfB = dfB.copy()

    dfA["HOR"] = dfA["HOR"].astype(str)
    dfA["HORA +1"] = dfA["HORA +1"].astype(str)
    dfB["HORA"] = dfB["HORA"].astype(str)

    dfA["KEY"] = dfA.iloc[:, 3].astype(str)+"_"+dfA.iloc[:, 7].astype(str)+"_"+dfA.iloc[:, 11].astype(str)+"_"+dfA["HOR"]
    dfA["KEY2"] = dfA.iloc[:, 3].astype(str)+"_"+dfA.iloc[:, 7].astype(str)+"_"+dfA.iloc[:, 11].astype(str)+"_"+dfA["HORA +1"]

    dfB["KEY"] = dfB.iloc[:, 0].astype(str)+"_"+dfB.iloc[:, 6].astype(str)+"_"+dfB.iloc[:, 4].astype(str)+"_"+dfB.iloc[:, 5]

    mask = dfA["KEY"].isin(dfB["KEY"]) | dfA["KEY2"].isin(dfB["KEY"])

    noA = dfA[~mask].drop(columns=["KEY","KEY2"], errors="ignore")
    noB = dfB[~dfB["KEY"].isin(pd.concat([dfA["KEY"], dfA["KEY2"]]))].drop(columns=["KEY"], errors="ignore")

    return _armar_salida(noA, noB)


def cruzar_sal325_772(df_325_list, df_772_list):
    return cruzar_sal320_762(df_325_list, df_772_list)


def cruzar_ent320_721(df_320_list, df_721_list):

    dfA = pd.concat([d for d in df_320_list if d is not None], ignore_index=True)
    dfB = pd.concat([d for d in df_721_list if d is not None], ignore_index=True)

    if dfA.empty or dfB.empty:
        return pd.DataFrame()

    dfA = dfA.copy()
    dfB = dfB.copy()

    dfA["HOR"] = dfA["HOR"].astype(str)
    dfA["HORA +3"] = dfA["HORA +3"].astype(str)
    dfB["HORA"] = dfB["HORA"].astype(str)

    dfA["KEY"] = dfA.iloc[:, 3].astype(str)+"_"+dfA.iloc[:, 7].astype(str)+"_"+dfA["T.CTA BN DEST"].astype(str)+"_"+dfA["HOR"]
    dfA["KEY2"] = dfA.iloc[:, 3].astype(str)+"_"+dfA.iloc[:, 7].astype(str)+"_"+dfA["T.CTA BN DEST"].astype(str)+"_"+dfA["HORA +3"]

    dfB["KEY"] = dfB.iloc[:, 0].astype(str)+"_"+dfB.iloc[:, 6].astype(str)+"_"+dfB.iloc[:, 4].astype(str)+"_"+dfB.iloc[:, 5]

    mask = dfA["KEY"].isin(dfB["KEY"]) | dfA["KEY2"].isin(dfB["KEY"])

    noA = dfA[~mask].drop(columns=["KEY","KEY2"], errors="ignore")
    noB = dfB[~dfB["KEY"].isin(pd.concat([dfA["KEY"], dfA["KEY2"]]))].drop(columns=["KEY"], errors="ignore")

    return _armar_salida(noA, noB)


def aplicar_formatos(writer, sheet_name, df):
    ws = writer.sheets[sheet_name]
    text_format = writer.book.add_format({'num_format': '@'})

    for i, col in enumerate(df.columns):
        if col not in ['IMP', 'IMPORTE']:
            ws.set_column(i, i, None, text_format)


def escribir(writer, df, nombre):
    df.to_excel(writer, sheet_name=nombre, index=False)
    aplicar_formatos(writer, nombre, df)


def ejecutar(archivo_cce, archivo_dpah):

    output_dir = os.path.dirname(archivo_cce)
    output_file = os.path.join(output_dir, "DIFERENCIAS CCE - DPAH.xlsx")

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

        
        escribir(writer, cruzar_sal320_762(
            [leer_hoja(archivo_cce,"SAL 320"), leer_hoja(archivo_cce,"SAL 320-2")],
            [leer_hoja(archivo_dpah,"SAL 762"), leer_hoja(archivo_dpah,"SAL 762-2")]
        ), "SAL 762 v 320")

        escribir(writer, cruzar_sal325_772(
            [leer_hoja(archivo_cce,"SAL 325"), leer_hoja(archivo_cce,"SAL 325-2")],
            [leer_hoja(archivo_dpah,"SAL 772"), leer_hoja(archivo_dpah,"SAL 772-2")]
        ), "TAR 772 v 325")

        escribir(writer, cruzar_ent320_721(
            [leer_hoja(archivo_cce,"ENT 320"), leer_hoja(archivo_cce,"ENT 320-2")],
            [leer_hoja(archivo_dpah,"ENT 721"), leer_hoja(archivo_dpah,"ENT 721-2")]
        ), "ENT 721 v 320")

       
        for hoja in ["RECHAZADOS", "T.C"]:
            df_extra = leer_hoja(archivo_cce, hoja)
            if df_extra is not None:
                escribir(writer, df_extra, hoja)

    return f"Diferencias generadas correctamente:\n{output_file}"


if __name__ == "__main__":
    cce, dpah = seleccionar_archivos()
    print(ejecutar(cce, dpah))
