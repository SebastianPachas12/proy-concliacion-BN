import xlsxwriter
import datetime
import os
import tkinter as tk
from tkinter import filedialog

# Función para procesar una línea que comienza con 'MV'
def procesar_linea(linea):
    campos = linea.strip().split()
    if len(campos) < 5 or not campos[0].startswith('MV'):
        return None  # No procesar si no cumple
    
    mv_part = campos[0]
    formato = None
    
    if len(mv_part) == 42:
        # Formato original: MV + 20 Origen + 20 Destino
        formato = 'original'
        origen = mv_part[2:22]
        destino = mv_part[22:42]
        importe_str = campos[1]
        campo3 = campos[2]
        id_cce = campos[3]
        sit_bco = campos[4]
    elif len(mv_part) == 22 and len(campos) >= 6:
        # Nuevo formato: MV + 20 Origen, luego Destino separado
        formato = 'nuevo'
        origen = mv_part[2:22]
        destino_raw = campos[1]
        destino = (' ' * (20 - len(destino_raw))) + destino_raw  # Completar a 20 con espacios al inicio
        importe_str = campos[2]
        campo3 = campos[3]
        id_cce = campos[4]
        sit_bco = campos[5]
    else:
        return None
    
    # Extraer de Origen (igual para ambos formatos)
    bco_orig = origen[0:3]
    cci_ori = origen
    core_ori = origen[6] if len(origen) > 6 else ''
    tcta_bn_ori = origen[7:18] if len(origen) > 17 else origen[7:]
    
    # Extraer de Destino
    cci_dest = destino
    if formato == 'original':
        core_dest = destino[6] if len(destino) > 6 else ''
        tcta_bn_dest = destino[7:18] if len(destino) > 17 else destino[7:]
    else:  # nuevo
        core_dest = destino[6] if len(destino) > 6 else ''
        tcta_bn_dest = destino[7:18] if len(destino) > 17 else destino[7:]
    
    # Procesar Importe: '700000' -> 7000.00
    imp = float(importe_str[:-2] + '.' + importe_str[-2:]) if len(importe_str) >= 2 else 0.0
    
    # Procesar campo3
    comis = None
    signo = None
    trf = None
    fecha = None
    hora = None
    canal = None
    trace = None
    
    if '+' in campo3 or '-' in campo3:
        # Caso con signo: '80+3202025102316011591030358' o '100-...'
        if '+' in campo3:
            partes = campo3.split('+')
            signo = '+'
        else:
            partes = campo3.split('-')
            signo = '-'
        if len(partes) == 2:
            comis_str = partes[0]
            resto = partes[1]
            # COMIS: convertir a decimal
            comis = float(comis_str[:-2] + '.' + comis_str[-2:]) if len(comis_str) >= 2 else 0.0
        else:
            resto = campo3  # Si no se parte bien, asumir sin signo
    else:
        resto = campo3
    
    # Procesar resto: TRF, FECHA, HORA, CANAL, TRACE
    if len(resto) >= 3 + 8 + 6 + 2:  # Mínimo esperado
        trf = resto[0:3]
        fecha_str = resto[3:11]
        hora = resto[11:17]
        canal = resto[17:19]
        trace = resto[19:]
        # Convertir FECHA a datetime con horas, minutos y segundos siempre en 0
        try:
            fecha = datetime.datetime.strptime(fecha_str + '000000', '%Y%m%d%H%M%S')
        except ValueError:
            fecha = None
    else:
        trf = resto if len(resto) >= 3 else ''
        fecha = None
        hora = ''
        canal = ''
        trace = ''
    
    # ID CCE
    id_cce_val = id_cce
    
    # SIT y BCO DEST
    sit = sit_bco[:7] if len(sit_bco) >= 7 else sit_bco
    bco_dest = sit_bco[-4:] if len(sit_bco) >= 4 else ''
    
    # Retornar lista con los valores en orden de columnas
    return [
        bco_orig,      # BCO ORIG
        cci_ori,       # CCI ORI
        core_ori,      # CORE ORI
        tcta_bn_ori,   # T.CTA BN ORI
        cci_dest,      # CCI DEST
        core_dest,      # CORE DEST
        tcta_bn_dest,   # T.CTA BN DEST
        imp,           # IMP
        comis,         # COMIS
        signo,         # SIGNO
        trf,           # TRF
        fecha,         # FECHA (ahora datetime con 00:00:00)
        hora,          # HORA
        canal,          # CANAL
        trace,         # TRACE
        id_cce_val,    # ID CCE
        sit,           # SIT
        bco_dest       # BCO DEST
    ]

# Función para incrementar la hora en segundos
def incrementar_hora(hora_str, segundos_a_agregar):
    if not hora_str or len(hora_str) != 6:
        return hora_str  # Retornar tal cual si no es válido
    try:
        h = int(hora_str[0:2])
        m = int(hora_str[2:4])
        s = int(hora_str[4:6])
        total_seg = h * 3600 + m * 60 + s + segundos_a_agregar
        h_new = total_seg // 3600
        m_new = (total_seg % 3600) // 60
        s_new = total_seg % 60
        return f"{h_new:02d}{m_new:02d}{s_new:02d}"
    except ValueError:
        return hora_str

# Función para seleccionar archivo
def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    ruta_txt = filedialog.askopenfilename(title="Selecciona el archivo .txt", filetypes=[("Archivos de texto", "*.txt")])
    root.destroy()
    return ruta_txt

# Función para escribir filas en hojas, manejando el límite de filas (1,048,575 filas de datos por hoja)
def write_to_sheets(workbook, sheet_name, rows, headers, write_func, incrementos=None):
    max_rows_per_sheet = 1048575  # Máximo filas de datos por hoja (header en fila 0)
    sheet_num = 1
    current_sheet = workbook.add_worksheet(f"{sheet_name}" if sheet_num == 1 else f"{sheet_name}-{sheet_num}")
    
    # Definir formatos (asumiendo que están definidos globalmente)
    text_format = workbook.add_format({'num_format': '@'})
    accounting_format = workbook.add_format({'num_format': '#,##0.00'})
    custom_date_format = workbook.add_format({'num_format': 'yyyy/mm/dd hh:mm:ss'})
    general_format = workbook.add_format({'num_format': 'General'})
    
    # Escribir headers
    for col_num, header in enumerate(headers):
        current_sheet.write(0, col_num, header)
    
    row_num = 1
    for fila in rows:
        if row_num > max_rows_per_sheet:
            sheet_num += 1
            current_sheet = workbook.add_worksheet(f"{sheet_name}-{sheet_num}")
            # Escribir headers de nuevo
            for col_num, header in enumerate(headers):
                current_sheet.write(0, col_num, header)
            row_num = 1
        write_func(current_sheet, row_num, fila, incrementos, text_format, accounting_format, custom_date_format, general_format)
        row_num += 1

# Función para escribir fila en GENERAL
def write_general_row(ws, row_num, fila, incrementos, text_format, accounting_format, custom_date_format, general_format):
    ws.write(row_num, 0, fila[0], text_format)   # BCO ORIG
    ws.write(row_num, 1, fila[1], text_format)   # CCI ORI
    ws.write(row_num, 2, fila[2], text_format)   # CORE ORI
    ws.write(row_num, 3, fila[3], text_format)   # T.CTA BN ORI
    ws.write(row_num, 4, fila[4], text_format)   # CCI DEST
    ws.write(row_num, 5, fila[5], text_format)   # CORE DEST
    ws.write(row_num, 6, fila[6], text_format)   # T.CTA BN DEST
    ws.write(row_num, 7, fila[7], accounting_format)  # IMP
    if fila[8] is not None:
        ws.write(row_num, 8, fila[8], accounting_format)  # COMIS
    else:
        ws.write(row_num, 8, fila[8], text_format)
    ws.write(row_num, 9, fila[9], general_format)  # SIGNO
    ws.write(row_num, 10, fila[10], text_format)  # TRF
    ws.write(row_num, 11, fila[11], custom_date_format)  # FECHA
    ws.write(row_num, 12, fila[12], text_format)  # HORA
    ws.write(row_num, 13, fila[13], text_format)  # CANAL
    ws.write(row_num, 14, fila[14], text_format)  # TRACE
    ws.write(row_num, 15, fila[15], text_format)  # ID CCE
    ws.write(row_num, 16, fila[16], text_format)  # SIT
    ws.write(row_num, 17, fila[17], text_format)  # BCO DEST

# Función para escribir fila en otras hojas (con incrementos de hora)
def write_row_otras(ws, row_num, fila, incrementos, text_format, accounting_format, custom_date_format, general_format):
    ws.write(row_num, 0, fila[0], text_format)   # BCO ORIG
    ws.write(row_num, 1, fila[1], text_format)   # CCI ORI
    ws.write(row_num, 2, fila[2], text_format)   # CORE ORI
    ws.write(row_num, 3, fila[3], text_format)   # T.CTA BN ORI
    ws.write(row_num, 4, fila[4], text_format)   # CCI DEST
    ws.write(row_num, 5, fila[5], text_format)   # CORE DEST
    ws.write(row_num, 6, fila[6], text_format)   # T.CTA BN DEST
    ws.write(row_num, 7, fila[7], accounting_format)  # IMP
    if fila[8] is not None:
        ws.write(row_num, 8, fila[8], accounting_format)  # COMIS
    else:
        ws.write(row_num, 8, fila[8], text_format)
    ws.write(row_num, 9, fila[9], general_format)  # SIGNO
    ws.write(row_num, 10, fila[10], text_format)  # TRF
    ws.write(row_num, 11, fila[11], custom_date_format)  # FECHA
    ws.write(row_num, 12, fila[12], text_format)  # HORA
    col = 13
    for inc in incrementos:
        nueva_hora = incrementar_hora(fila[12], inc)
        ws.write(row_num, col, nueva_hora, text_format)
        col += 1
    ws.write(row_num, col, fila[13], text_format)  # CANAL
    col += 1
    ws.write(row_num, col, fila[14], text_format)  # TRACE
    col += 1
    ws.write(row_num, col, fila[15], text_format)  # ID CCE
    col += 1
    ws.write(row_num, col, fila[16], text_format)  # SIT
    col += 1
    ws.write(row_num, col, fila[17], text_format)  # BCO DEST


def procesar_archivo_txt(ruta_txt):
    
    if not os.path.exists(ruta_txt):
        raise FileNotFoundError(f"El archivo no existe:\n{ruta_txt}")
    
    directorio = os.path.dirname(ruta_txt)
    os.makedirs(directorio, exist_ok=True)
    
    nombre_base = os.path.splitext(os.path.basename(ruta_txt))[0]
    nombre_excel = f"{nombre_base}.xlsx"
    ruta_excel = os.path.join(directorio, nombre_excel)
    
    workbook = xlsxwriter.Workbook(ruta_excel)
    
    headers_base = ['BCO ORIG', 'CCI ORI', 'CORE ORI', 'T.CTA BN ORI', 'CCI DEST', 'CORE DEST', 'T.CTA BN DEST', 'IMP', 'COMIS', 'SIGNO', 'TRF', 'FEC', 'HOR', 'CANAL', 'TRACE', 'ID CCE', 'SIT', 'BCO DEST']
    
    headers_sal = headers_base[:13] + ['HORA +1'] + headers_base[13:]
    
    headers_ent = headers_base[:13] + ['HORA +3'] + headers_base[13:]
    
    general_rows = []
    sal320_rows = []
    sal325_rows = []
    ent320_rows = []
    rech_rows = []
    tc_rows = []  
    
    with open(ruta_txt, 'r', encoding='utf-8') as f:
        for linea in f:
            fila = procesar_linea(linea)
            if fila:
                general_rows.append(fila)
                
                bco_orig = fila[0]
                trf = fila[10]
                sit = fila[16]
                bco_dest = fila[17]
                comis = fila[8]
                
                if bco_orig == '018' and trf == '320' and sit == '490000A':
                    sal320_rows.append(fila)
                
                if bco_orig == '018' and trf == '325' and sit == '490000A':
                    sal325_rows.append(fila)
                
                if bco_dest == '0018' and trf == '320' and sit == '490000A':
                    ent320_rows.append(fila)
                
                if bco_orig == '018' and trf == '320' and sit == '490000R' and comis is not None:
                    rech_rows.append(fila)
                
                if bco_orig == '018' and trf == '325' and comis is not None and comis != 0.0:
                    tc_rows.append(fila)
    

    sal320_rows.sort(key=lambda x: (x[3], x[11] or datetime.datetime.min))
    sal325_rows.sort(key=lambda x: (x[3], x[11] or datetime.datetime.min))
    rech_rows.sort(key=lambda x: (x[3], x[11] or datetime.datetime.min))
    tc_rows.sort(key=lambda x: (x[3], x[11] or datetime.datetime.min))  
    ent320_rows.sort(key=lambda x: (x[6], x[11] or datetime.datetime.min))
    
    write_to_sheets(workbook, 'CCE GEN', general_rows, headers_base, write_general_row)
    write_to_sheets(workbook, 'SAL 320', sal320_rows, headers_sal, write_row_otras, [1])
    write_to_sheets(workbook, 'SAL 325', sal325_rows, headers_sal, write_row_otras, [1])
    write_to_sheets(workbook, 'ENT 320', ent320_rows, headers_ent, write_row_otras, [3])  # Cambiado a solo [3], omitiendo +4
    write_to_sheets(workbook, 'RECHAZADOS', rech_rows, headers_sal, write_row_otras, [1])
    write_to_sheets(workbook, 'T.C', tc_rows, headers_sal, write_row_otras, [1])  # Nueva hoja T.C con mismo formato que SAL 325
    
    # Cerrar workbook
    workbook.close()
    print(f"Archivo Excel creado: {ruta_excel}")

def ejecutar(ruta_txt):

        if not ruta_txt:
            raise ValueError("No se selecciono ningun archivo.")
        procesar_archivo_txt(ruta_txt)
        return f"CCE procesado correctamente: {os.path.basename(ruta_txt)}"

# Ejecutar
if __name__ == "__main__":
    ruta = seleccionar_archivo()
    print(ejecutar(ruta))
