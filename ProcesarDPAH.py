import pandas as pd
from datetime import datetime, timedelta
import os
from collections import defaultdict
import math
from tkinter import Tk, filedialog


MAX_EXCEL_ROWS = 1048576
DATA_ROWS_LIMIT = MAX_EXCEL_ROWS - 1


def seleccionar_archivos():
    root = Tk()
    root.withdraw()

    rutas = filedialog.askopenfilenames(
        title="Seleccionar archivos TXT",
        filetypes=[("TXT", "*.txt")]
    )

    root.destroy()
    return rutas


def parse_txt_to_dataframe(file_path):

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"No existe:\n{file_path}")

    rows = []

    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:

            line = line.replace("\t", " ").strip()

            if not ((line.startswith("0-") or line.startswith("4-")) and "TINL" in line):
                continue

            parts = line.split()
            if len(parts) < 10:
                continue

            nro_cta, age, caj, tnx, fecha_str, hora = parts[:6]

            importe = None
            importe_index = None

            for i in range(6, len(parts)):
                try:
                    val = float(parts[i].replace(",", ""))
                    if "." in parts[i]:
                        importe = val
                        importe_index = i
                        break
                except:
                    continue

            if importe_index is None:
                continue

            operacion = " ".join(parts[importe_index + 1:-2])
            nro_tarjeta = parts[-2]
            tsist = parts[-1]

            try:
                fecha = datetime.strptime(fecha_str, "%Y%m%d").date()
            except:
                continue

            nro_cta = "0" + nro_cta.replace("-", "")

            rows.append([
                nro_cta, age, caj, tnx, fecha, hora,
                importe, operacion, nro_tarjeta, tsist
            ])

    return pd.DataFrame(rows, columns=[
        "Nro Cta", "AGE", "CAJ", "TNX", "FECHA",
        "HORA", "IMPORTE", "OPERACION",
        "Nro Tarjeta", "TSIST"
    ])


def split_by_date_and_hour(df):
    df["HORA"] = df["HORA"].astype(str).str.zfill(6)
    result = defaultdict(dict)

    for fecha, sub in df.groupby("FECHA"):
        result[fecha]["ANT"] = sub[sub["HORA"] <= "155999"]
        result[fecha]["DES"] = sub[sub["HORA"] >= "160000"]

    return result


def write_df_with_split(writer, df, base_sheet_name):

    total_rows = len(df)

    if total_rows == 0:
        df.to_excel(writer, sheet_name=base_sheet_name, index=False)
        return

    total_sheets = math.ceil(total_rows / DATA_ROWS_LIMIT)

    for i in range(total_sheets):
        start = i * DATA_ROWS_LIMIT
        end = start + DATA_ROWS_LIMIT

        chunk = df.iloc[start:end]
        sheet_name = base_sheet_name if i == 0 else f"{base_sheet_name}-{i+1}"

        chunk.to_excel(writer, sheet_name=sheet_name, index=False)


def export_excel(df, file_path):

    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:

        write_df_with_split(writer, df, "DPAH GEN")

        filtros = {
            "SAL 762": "762",
            "SAL 772": "772",
            "COMIS 512": "512",
            "ENT 721": "721",
            "COMIS 562": "562"
        }

        for sheet, tnx in filtros.items():
            write_df_with_split(writer, df[df["TNX"] == tnx], sheet)

    return file_path



def combine_consecutive_days(data_dict, output_dir):

    fechas = sorted(data_dict.keys())
    generados = []

    for i in range(len(fechas) - 1):

        f1, f2 = fechas[i], fechas[i+1]

        if f2 != f1 + timedelta(days=1):
            continue

        if "DES" not in data_dict[f1] or "ANT" not in data_dict[f2]:
            continue

        df = pd.concat(
            [data_dict[f1]["DES"], data_dict[f2]["ANT"]],
            ignore_index=True
        )

        df.sort_values(["Nro Cta", "FECHA", "HORA"], inplace=True)

        output_file = os.path.join(output_dir, f"DPAH {f1.day}-{f2.day}.xlsx")

        export_excel(df, output_file)
        generados.append(output_file)

    return generados


# =========================
# EJECUTAR (para GUI)
# =========================
def ejecutar(files):

    if not files:
        raise ValueError("No se seleccionaron archivos TXT.")

    output_dir = os.path.dirname(files[0])
    os.makedirs(output_dir, exist_ok=True)

    df_total = pd.concat(
        (parse_txt_to_dataframe(f) for f in files),
        ignore_index=True
    )

    data_dict = split_by_date_and_hour(df_total)
    generados = combine_consecutive_days(data_dict, output_dir)

    return f"Proceso completado.\nArchivos generados: {len(generados)}"


# =========================
# TEST LOCAL
# =========================
if __name__ == "__main__":
    files = seleccionar_archivos()
    print(ejecutar(files))
