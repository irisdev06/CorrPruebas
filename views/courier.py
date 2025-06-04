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

# Genera un archivo Excel en bytes con hojas por medio de envío.
def generar_excel(datos: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for medio in datos['MEDIO DE ENVIO'].dropna().unique():
            df_medio = datos[datos['MEDIO DE ENVIO'] == medio]
            df_medio.to_excel(writer, sheet_name=str(medio)[:31], index=False)
    return output.getvalue()