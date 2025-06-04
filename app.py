import streamlit as st
from views.courier import (
    cargar_datos,
    rellenar_fecha_recibido,
    calcular_indicador,
    agregar_termino,
    generar_excel
)


st.title("📅 Procesamiento de fechas en correspondencia")
archivo = st.file_uploader("Sube el archivo CSV con las fechas", type=["csv"])
if archivo is not None:
    datos = cargar_datos(archivo)
    datos = rellenar_fecha_recibido(datos)
    datos = calcular_indicador(datos)
    datos = agregar_termino(datos)
    
    # Mostrar tabla con datos procesados (opcional)
    # st.dataframe(datos)
    
    excel_bytes = generar_excel(datos)
    
    # Botón para descargar el archivo
    st.download_button(
        label="⬇️ Descargar Excel por Medio de Envío",
        data=excel_bytes,
        file_name="correspondencia_por_medio_envio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
