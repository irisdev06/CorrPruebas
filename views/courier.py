import pandas as pd
import matplotlib.pyplot as plt
import tempfile
import holidays
from datetime import date
from io import BytesIO
from datetime import datetime, timedelta


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

#  Hoja BASE
def generar_hoja_base(datos: pd.DataFrame, writer) -> pd.DataFrame:
    df_consolidado, df_courier = obtener_dfs_filtrados(datos)

    # Unir los datos de Consolidado y Courier
    df_base = pd.concat([df_consolidado, df_courier], ignore_index=True)
    
    # Si df_base no está vacío, crear la hoja BASE en el archivo Excel
    if not df_base.empty:
        agregar_columnas_vacias(df_base).to_excel(writer, sheet_name='BASE', index=False)

    return df_base

# Función para generar la hoja MEDIO DE ENVIO
def generar_medio_envio(df_base: pd.DataFrame, workbook) -> None:
    sheet_name = 'MEDIO DE ENVIO'
    worksheet = workbook.add_worksheet(sheet_name)

    formato_titulo = workbook.add_format({'bold': True, 'bg_color': "#2480E9"})
    formato_celdas = workbook.add_format({'text_wrap': True, 'valign': 'top'})

    # Escribir encabezado
    worksheet.write('A1', 'Proveedor', formato_titulo)

    # Obtener proveedores únicos de la columna 'Proveedor'
    proveedores = df_base['Proveedor'].dropna().unique()

    # Obtener medios de envío únicos de la columna 'MEDIO DE ENVIO', asegurándonos de capturar todos
    medios_envio = df_base['MEDIO DE ENVIO'].dropna().unique()

    # Escribir los encabezados de los medios de envío
    for i, medio_envio in enumerate(medios_envio, start=1):
        worksheet.write(0, i, medio_envio, formato_titulo)

    # Columna adicional para el Total
    worksheet.write(0, len(medios_envio) + 1, 'Total', formato_titulo)

    startrow = 1  # Empezamos en la fila 2 para los datos

    # Iterar sobre los proveedores
    for proveedor in proveedores:
        df_proveedor = df_base[df_base['Proveedor'] == proveedor]
        if df_proveedor.empty:
            continue

        # Escribir proveedor en la columna A
        worksheet.write(startrow, 0, proveedor, formato_celdas)

        total_proveedor = 0  # Variable para el total del proveedor

        # Calcular los totales por medio de envío
        for i, medio_envio in enumerate(medios_envio, start=1):
            # Filtrar por medio de envío
            df_medio_envio = df_proveedor[df_proveedor['MEDIO DE ENVIO'] == medio_envio]

            # Escribir el total para cada medio de envío
            total_medio = len(df_medio_envio)
            worksheet.write(startrow, i, total_medio, formato_celdas)

            # Sumar al total del proveedor
            total_proveedor += total_medio

        # Escribir el total del proveedor en la columna de "Total"
        worksheet.write(startrow, len(medios_envio) + 1, total_proveedor, formato_celdas)

        startrow += 1  # Incrementar fila

    # Fila de totales
    total_row = startrow
    worksheet.write(total_row, 0, 'Total', formato_titulo)

    # Sumar los totales por cada medio de envío (en la última fila de la tabla)
    for i in range(1, len(medios_envio) + 1):
        worksheet.write_formula(total_row, i, f'SUM({chr(65 + i)}2:{chr(65 + i)}{total_row})', formato_celdas)

    # Sumar los totales verticales (la columna Total)
    worksheet.write_formula(total_row, len(medios_envio) + 1, f'SUM({chr(65 + len(medios_envio) + 1)}2:{chr(65 + len(medios_envio) + 1)}{total_row})', formato_celdas)

    worksheet.autofilter(0, 0, startrow, len(medios_envio) + 1)

# Función para generar la hoja Alerta
# Función para generar la hoja Alerta
def generar_alerta(df_courier: pd.DataFrame, workbook) -> None:
    sheet_name = 'Alerta'
    worksheet = workbook.add_worksheet(sheet_name)

    formato_titulo = workbook.add_format({'bold': True, 'bg_color': "#FF9F00"})
    formato_celdas = workbook.add_format({'text_wrap': True, 'valign': 'top'})

    # Obtener la fecha de ayer
    hoy = datetime.today()
    ayer = hoy - timedelta(days=1)

    # Escribir encabezados
    worksheet.write('A1', 'Fecha de Radicación', formato_titulo)
    worksheet.write('B1', ayer.strftime('%Y-%m-%d'), formato_titulo)  # Fecha de ayer
    worksheet.write('C1', 'Total general', formato_titulo)

    # Filtrar los registros que corresponden al día de ayer
    df_courier['FECHA RADICACION'] = pd.to_datetime(df_courier['FECHA RADICACION'], errors='coerce')
    df_ayer = df_courier[df_courier['FECHA RADICACION'].dt.date == ayer.date()]

    # Obtener los proveedores únicos y contar los registros
    proveedores_ayer = df_ayer['Proveedor'].dropna().unique()

    startrow = 1  # Comenzamos a escribir en la segunda fila

    total_general = 0  # Inicializamos el contador total

    # Escribir los proveedores y la cantidad de registros por proveedor
    for proveedor in proveedores_ayer:
        df_proveedor_ayer = df_ayer[df_ayer['Proveedor'] == proveedor]
        total_registros = len(df_proveedor_ayer)

        worksheet.write(startrow, 0, proveedor, formato_celdas)
        worksheet.write(startrow, 1, total_registros, formato_celdas)

        total_general += total_registros  # Sumar al total general

        startrow += 1  # Incrementar fila

    # Escribir el total general en la fila siguiente
    worksheet.write(startrow, 0, 'Total', formato_titulo)
    worksheet.write(startrow, 1, total_general, formato_titulo)

    worksheet.autofilter(0, 0, startrow, 2)  # Filtrar las columnas A, B y C

# Función para crear un gráfico de pastel basado en los datos de "MEDIO DE ENVIO"
def generar_grafico_pastel(df_base: pd.DataFrame, workbook) -> None:
    # Contamos los registros por medio de envío
    conteo_medios_envio = df_base['MEDIO DE ENVIO'].value_counts()

    # Crear un gráfico de pastel
    fig, ax = plt.subplots()
    ax.pie(conteo_medios_envio, labels=conteo_medios_envio.index, autopct='%1.1f%%', startangle=90)
    ax.axis('equal')  # Para que el gráfico de pastel sea circular

    # Crear un archivo temporal para guardar el gráfico
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmpfile:
        img_path = tmpfile.name
        plt.savefig(img_path)
        plt.close()

    # Insertar la imagen en la hoja de Excel
    sheet_name = 'MEDIO DE ENVIO'
    worksheet = workbook.get_worksheet_by_name(sheet_name)
    
    worksheet.insert_image('F2', img_path)

# Función para crear un gráfico de barras apiladas basado en "FECHA RADICACION" y "Proveedor"
def generar_grafico_barras_apiladas(df_courier: pd.DataFrame, workbook) -> None:
    # Agrupar los datos por FECHA RADICACION y Proveedor
    df_agrupado = df_courier.groupby(['FECHA RADICACION', 'Proveedor']).size().unstack(fill_value=0)

    # Crear un gráfico de barras apiladas
    ax = df_agrupado.plot(kind='bar', stacked=True, figsize=(10, 6))

    # Añadir título y etiquetas
    ax.set_title('Cantidad de Registros por Proveedor y Fecha de Radicación', fontsize=14)
    ax.set_xlabel('Fecha de Radicación', fontsize=12)
    ax.set_ylabel('Cantidad de Registros', fontsize=12)

    # Rotar las etiquetas del eje X para mejor visibilidad
    plt.xticks(rotation=45)

    # Crear un archivo temporal para guardar el gráfico
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmpfile:
        img_path = tmpfile.name
        plt.tight_layout()  # Asegurarse de que el gráfico se ajuste bien
        plt.savefig(img_path)
        plt.close()

    # Insertar la imagen en la hoja de Excel
    sheet_name = 'Alerta'  # Puedes cambiar esto al nombre de la hoja en la que deseas insertar el gráfico
    worksheet = workbook.get_worksheet_by_name(sheet_name)
    
    worksheet.insert_image('F2', img_path)
# Generar Excel
def generar_excel(datos: pd.DataFrame) -> bytes:
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Llamar a la función para generar la hoja BASE y obtener df_base
        df_base = generar_hoja_base(datos, writer)

        # HOJA: Courier (mantenerla tal como estaba antes)
        df_consolidado, df_courier = obtener_dfs_filtrados(datos)
        if not df_courier.empty:
            agregar_columnas_vacias(df_courier).to_excel(writer, sheet_name='Courier', index=False)

            # HOJAS: Un proveedor por hoja
            for nombre_hoja, df_proveedor in obtener_dfs_por_proveedor(df_courier):
                agregar_columnas_vacias(df_proveedor).to_excel(writer, sheet_name=nombre_hoja, index=False)

        # Llamar a la función para generar la hoja MEDIO DE ENVIO
        generar_medio_envio(df_base, workbook)

        # Llamar a la función para generar la hoja ALERTA
        generar_alerta(df_courier, workbook)

        # Llamar a la función para generar el gráfico de pastel en la hoja MEDIO DE ENVIO
        generar_grafico_pastel(df_base, workbook)
        
        # Llamar a la función para generar el gráfico de barras apiladas en la hoja Alerta
        generar_grafico_barras_apiladas(df_courier, workbook)
    return output.getvalue()
