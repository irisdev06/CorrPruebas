import pandas as pd
import holidays
from datetime import date
from io import BytesIO

# Cargar el CSV y convierte las columnas de fecha
def cargar_datos(archivo) -> pd.DataFrame:
    datos = pd.read_csv(archivo, sep=";")
    datos['FECHA RADICACION'] = pd.to_datetime(datos['FECHA RADICACION'], errors='coerce', dayfirst=False)
    datos['FECHA RECIBIDO CORRESPONDENCIA'] = pd.to_datetime(datos['FECHA RECIBIDO CORRESPONDENCIA'], errors='coerce', dayfirst=False)
    return datos

# Rellenar fecha de recibido con fecha actual
def rellenar_fecha_recibido(datos: pd.DataFrame) -> pd.DataFrame:
    fecha_actual = pd.Timestamp(date.today())
    datos['FECHA RECIBIDO CORRESPONDENCIA'] = datos['FECHA RECIBIDO CORRESPONDENCIA'].fillna(fecha_actual)
    return datos

# Diferencia de días
def calcular_indicador(datos: pd.DataFrame) -> pd.DataFrame:
    years = list(set(datos['FECHA RADICACION'].dt.year.dropna().astype(int).tolist() +
                     datos['FECHA RECIBIDO CORRESPONDENCIA'].dt.year.dropna().astype(int).tolist()))
    col_holidays = holidays.Colombia(years=years)

    def dias_habiles(row):
        start = row['FECHA RADICACION'].normalize()
        end = row['FECHA RECIBIDO CORRESPONDENCIA'].normalize()
        if pd.isna(start) or pd.isna(end) or start > end:
            return None  
        bdays = pd.bdate_range(start, end).difference(pd.to_datetime(list(col_holidays.keys())))
        dias = len(bdays) - 1
        return dias if dias >= 0 else 0

    datos["INDICADOR"] = datos.apply(dias_habiles, axis=1)
    return datos

# Evaluar Término
def evaluar_termino(dias: int) -> str:
    return "EN TERMINO" if 0 <= dias < 2 else "FUERA DE TERMINO"

# Crear columna indicador
def agregar_termino(datos: pd.DataFrame) -> pd.DataFrame:
    datos['TERMINO'] = datos['INDICADOR'].apply(evaluar_termino)
    return datos

# Clasificación por Proveedor
def generarcol_proveedor(datos: pd.DataFrame) -> pd.DataFrame:
    proveedor_map = {
        '3 GRUPO JUNTAS DE CALIFICACIÓN': 'BELISARIO',
        '3 GRUPO CENTRO DE EXCELENCIA': 'BELISARIO',
        '4 GRUPO JUNTAS DE CALIFICACIÓN': 'UTMDL',
        '4 GRUPO CENTRO DE EXCELENCIA': 'UTMDL',
        '5 GRUPO CENTRO DE EXCELENCIA': 'BELISARIO397',
        '5 GRUPO JUNTAS DE CALIFICACIÓN': 'BELISARIO397',
        '6 GRUPO CENTRO DE EXCELENCIA': 'GESTAR INNOVACION',
        '6 GRUPO JUNTAS DE CALIFICACIÓN': 'GESTAR INNOVACION',
        'GERENCIA MEDICA EXCELENCIA': 'GER.MED.EXCELENCIA',
        'GERENCIA MEDICA JUNTAS DE CALIFICACIÓN': 'GER.MED.JUNTAS DE CALIFICACIÓN'
    }

    datos['Proveedor'] = datos["DEPENDENCIA QUE ENVIA"].map(proveedor_map).fillna("DESCONOCIDO")
    return datos

COLUMNAS_EXTRA = ['OPORTUNIDAD FINAL', 'OBSERVACIÓN', 'DEFINICION']

def agregar_columnas_vacias(df: pd.DataFrame) -> pd.DataFrame:
    for col in COLUMNAS_EXTRA:
        df[col] = ''
    return df

def obtener_dfs_filtrados(datos: pd.DataFrame):
    df_consolidado = datos[datos['MEDIO DE ENVIO'] != 'Courier'].copy()
    df_courier = datos[datos['MEDIO DE ENVIO'] == 'Courier'].copy()
    return df_consolidado, df_courier

def obtener_dfs_por_proveedor(df_courier: pd.DataFrame):
    if 'Proveedor' not in df_courier.columns:
        return []
    proveedores = df_courier['Proveedor'].dropna().unique()
    return [
        (str(proveedor)[:31], df_courier[df_courier['Proveedor'] == proveedor].copy())
        for proveedor in proveedores
    ]

# Tablas
def generar_tabla_resumen(datos: pd.DataFrame) -> dict:
    proveedores_objetivo = ['UTMDL', 'GESTAR INNOVACION', 'BELISARIO397', 'BELISARIO']
    datos = datos[datos['Proveedor'].isin(proveedores_objetivo)].copy()

    tablas_por_proveedor = {}

    for proveedor in proveedores_objetivo:
        df_prov = datos[datos['Proveedor'] == proveedor].copy()

        resumen = df_prov.groupby(['MES']).agg(
            UNIVERSO=('TERMINO', 'count'),
            FUERA_DE_TERMINO=('TERMINO', lambda x: (x == 'FUERA DE TERMINO').sum()),
            EXCLUSIONES=('OBSERVACIÓN', lambda x: (x.str.contains('EXCLUIR', na=False)).sum() if 'OBSERVACIÓN' in x else 0),
            TERMINOS=('TERMINO', lambda x: (x == 'EN TERMINO').sum())
        ).reset_index()

        resumen['PORCENTAJE INDICADO'] = (
            (resumen['TERMINOS'] / resumen['UNIVERSO']) * 100
        ).round(2).astype(str) + '%'

        tablas_por_proveedor[proveedor] = resumen

    return tablas_por_proveedor

# Generar Excel
def generar_excel(datos: pd.DataFrame) -> bytes:
    output = BytesIO()
    df_consolidado, df_courier = obtener_dfs_filtrados(datos)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # HOJA: BASE
        df_base = pd.concat([df_consolidado, df_courier], ignore_index=True)  # Unir los dos DataFrames
        if not df_base.empty:
            agregar_columnas_vacias(df_base).to_excel(writer, sheet_name='BASE', index=False)

            # HOJAS: Un proveedor por hoja
            for nombre_hoja, df_proveedor in obtener_dfs_por_proveedor(df_courier):
                agregar_columnas_vacias(df_proveedor).to_excel(writer, sheet_name=nombre_hoja, index=False)

            # HOJA: IND COURIER (Resumen mensual por proveedor en una sola hoja)
            resumenes = generar_tabla_resumen(df_courier)
            sheet_name = 'IND COURIER'
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet

            formato_titulo = workbook.add_format({'bold': True, 'bg_color': "#F5F8C9"})

            startrow = 0
            for proveedor, df_resumen in resumenes.items():
                worksheet.write(startrow, 0, f"Proveedor: {proveedor}", formato_titulo)
                startrow += 1

                # Escribir encabezados
                for col_num, col_name in enumerate(df_resumen.columns):
                    worksheet.write(startrow, col_num, col_name, formato_titulo)

                # Escribir filas
                for row_num, row in enumerate(df_resumen.itertuples(index=False), start=startrow + 1):
                    for col_num, value in enumerate(row):
                        worksheet.write(row_num, col_num, value)

                worksheet.autofilter(startrow, 0, startrow + len(df_resumen), len(df_resumen.columns) - 1)
                startrow += len(df_resumen) + 3

    return output.getvalue()
