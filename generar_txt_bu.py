#!/usr/bin/env python3
"""
Generador de TXT Banco Unión (formato 110 chars/línea)
Soporta XLS binario (xlrd) y XLS/HTML (pandas read_html).
Uso:
    python3 generar_txt_bu.py <archivo.xls> [salida.txt]
    Si no se indica salida, el nombre se genera automáticamente.
Ejemplos:
    python3 generar_txt_bu.py "ExtractosBancarios (6).xls"
    python3 generar_txt_bu.py "ExtractosBancarios (39).xls" MTE_011719_16032026.txt
"""

import sys
import os
import unicodedata
import pandas as pd
import xlrd


def a_ascii(texto: str) -> str:
    """Elimina TODO carácter no-ASCII (acentos, °, ´, ñ, etc.)."""
    normalizado = unicodedata.normalize('NFD', str(texto))
    return normalizado.encode('ascii', errors='ignore').decode('ascii')


CUENTA_FIJA = 'BUNCC10000006036425  '  # 21 chars exactos
CABECERA    = '000               BUNCC10000006036425          CCSaldo Inicial'
PREFIJO     = 'MTE_011719'


def nombre_automatico(fecha_str: str) -> str:
    """Genera nombre MTE_011719_DDMMYYYY_HHMM.txt"""
    from datetime import datetime
    dd = fecha_str[6:8]
    mm = fecha_str[4:6]
    yyyy = fecha_str[0:4]
    hora = datetime.now().strftime('%H%M')
    return f"{PREFIJO}_{dd}{mm}{yyyy}_{hora}.txt"


def leer_excel(ruta_excel: str) -> pd.DataFrame:
    """Intenta leer el XLS como binario (xlrd) o como HTML (pandas)."""

    # — Intento 1: XLS binario con xlrd —
    try:
        wb = xlrd.open_workbook(ruta_excel)
        ws = wb.sheet_by_index(0)
        print(f"   Formato            : XLS binario ({ws.nrows} filas, {ws.ncols} cols)")

        inicio = 0
        for i in range(min(5, ws.nrows)):
            val = str(ws.cell_value(i, 1)).strip()
            if val and val != 'Fecha' and val != '':
                if '-' in val or val.replace('-', '').isdigit():
                    inicio = i
                    break

        filas = [ws.row_values(r) for r in range(inicio, ws.nrows)]
        df = pd.DataFrame(filas, columns=[
            'Consulta', 'Fecha', 'Sec', 'NroComprob', 'CodOper',
            'NroDocumento', 'Glosa', 'CuentaTransf',
            'Debitos', 'Creditos', 'Saldos', 'ImporteConciliar', 'Estado'
        ])
        return df

    except Exception as e:
        print(f"   xlrd falló ({e}), intentando con read_html...")

    # — Intento 2: XLS como HTML (pandas) —
    dfs = pd.read_html(ruta_excel)
    df = max(dfs, key=len)
    ncols = len(df.columns)
    print(f"   Formato            : HTML ({ncols} cols)")

    if ncols == 13:
        df.columns = [
            'Consulta', 'Fecha', 'Sec', 'NroComprob', 'CodOper',
            'NroDocumento', 'Glosa', 'CuentaTransf',
            'Debitos', 'Creditos', 'Saldos', 'ImporteConciliar', 'Estado'
        ]
    elif ncols == 16:
        df.columns = [
            'Consulta', 'Fecha', 'Sec', 'NroComprob', 'CodOper',
            'NroDocumento', 'Glosa', 'CuentaTransf',
            'DebitosOrigen', 'CreditosOrigen', 'SaldosOrigen',
            'Debitos', 'Creditos', 'Saldos',
            'ImporteConciliar', 'Estado'
        ]
    else:
        print(f"❌ Formato no reconocido: {ncols} columnas.")
        sys.exit(1)

    return df


def excel_a_txt(ruta_excel: str, ruta_salida: str = None):
    print(f"📂 Leyendo: {ruta_excel}")

    df = leer_excel(ruta_excel)

    # Limpiar filas de cabecera repetidas y vacías
    df = df[df['Fecha'].astype(str).str.strip() != ''].reset_index(drop=True)
    df = df[df['Fecha'].astype(str).str.strip() != 'Fecha'].reset_index(drop=True)

    print(f"   Total filas leídas : {len(df)}")

    # Filtrar DEP + TEC + CRV con crédito > 0
    dep = df[df['CodOper'].astype(str).str.strip().isin(['DEP', 'TEC', 'CRV'])].copy().reset_index(drop=True)
    dep['Creditos_f'] = (
        dep['Creditos']
        .astype(str)
        .str.replace(',', '', regex=False)
        .apply(lambda x: float(x) if x.strip() not in ('', 'nan') else 0.0)
    )
    dep = dep[dep['Creditos_f'] > 0].reset_index(drop=True)

    print(f"   Registros (DEP+TEC+CRV>0): {len(dep)}")

    ignorados = 0
    lineas = [CABECERA]
    fecha_str_global = None

    for idx, row in dep.iterrows():
        try:
            nro_int = int(float(str(row['NroDocumento'])))
        except (ValueError, TypeError):
            print(f"   ⚠️  Fila {idx}: NroDocumento no numérico → '{row['NroDocumento']}' — ignorado")
            ignorados += 1
            continue

        fecha_raw = str(row['Fecha']).strip()
        if '.' in fecha_raw and fecha_raw.replace('.', '').isdigit():
            fecha_dt = xlrd.xldate_as_datetime(float(fecha_raw), 0)
            fecha_str = fecha_dt.strftime('%Y%m%d')
        else:
            fecha_str = fecha_raw.replace('-', '')

        # Guardar primera fecha válida para el nombre del archivo
        if fecha_str_global is None and len(fecha_str) == 8:
            fecha_str_global = fecha_str

        monto_cent = int(round(row['Creditos_f'] * 100))
        monto_str  = str(monto_cent).ljust(18)
        glosa_fmt  = a_ascii(str(row['Glosa']))[:28].ljust(28)
        nro_fmt    = str(nro_int).ljust(11)
        resto      = '21 2    ' + ' ' * 15

        linea = monto_str + CUENTA_FIJA + fecha_str + 'C' + glosa_fmt + nro_fmt + resto

        if len(linea) != 110:
            print(f"   ❌ Fila {idx}: longitud {len(linea)} ≠ 110 — ignorado")
            ignorados += 1
            continue

        lineas.append(linea)

    # Determinar nombre de salida
    if ruta_salida is None:
        if fecha_str_global:
            ruta_salida = nombre_automatico(fecha_str_global)
        else:
            ruta_salida = nombre_automatico('00000000')

    with open(ruta_salida, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(lineas))

    generados = len(lineas) - 1
    print(f"\n✅ {generados} registros generados | {ignorados} ignorados")
    print(f"📄 Archivo: {os.path.abspath(ruta_salida)}")

    # Verificar longitudes
    errores = []
    with open(ruta_salida, 'r') as f:
        for i, line in enumerate(f, 1):
            line = line.rstrip('\n')
            if i > 1 and len(line) != 110:
                errores.append(f"   línea {i}: {len(line)} chars")
    if errores:
        print(f"⚠️  {len(errores)} líneas con longitud incorrecta:")
        for e in errores[:10]:
            print(e)
    else:
        print("✅ Todas las líneas miden exactamente 110 chars")

    return generados


# ──────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    entrada = sys.argv[1]
    salida  = sys.argv[2] if len(sys.argv) >= 3 else None

    if not os.path.exists(entrada):
        print(f"❌ No se encontró el archivo: {entrada}")
        sys.exit(1)

    excel_a_txt(entrada, salida)