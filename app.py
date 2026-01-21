
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import os
import warnings
import base64
import io
import re

# Configuración de la página
st.set_page_config(
    page_title="Procesador de Carátulas Bancarias",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

warnings.filterwarnings('ignore')

def cargar_mapeo_monedas():
    """Cargar el mapeo de monedas desde base64 embebido"""
    try:
        # El mismo mapeo embebido del script original
        mapeo_base64 = """
UEsDBBQAAAAIAG2YV1fWQA8kXwIAAK8EAAAQAAAAbWFwZW9fbW9uZWRhcy5jc3a9lE1uwyAQhfeV
eg/Pg8z8GFtnkQZnERW7jRArTXAIbqMq8t4lo6bKsmujNnYEm3nvG3jvMZdV7p2x2jhba2usNe7a
dsdRlhbHeVpslGVZnhd5fkCv0no3ztJiJ48t3lrjrPWu1vrQams9u/aO37K2TtdpdhpGl2bZMC1h
stZeGzZ3ViU7hUh1k1qjmzQ7DVLWDrW2trVdPF3N4fKl7eE5v8VtGHQ/6T5bEaGWRoH6JdGQfCjv
zJo6Pf5qJy7xzFyNHJT0wc5YJ8mjHjDLsmTbUMl2qwZd0gItFhyVNS7Jgp6xzlnEJVcYZXP6lOOQ
8JcL3L4U2Mc/oqkqSnXMKAf6lSKKnq6DqOqaJWLmr8AYgPfCzJBGMq6nKQZqZnKkAM5xnKSJ02M5
R/J2wkiKkpwjOYf+Tn9PUFbgHJvLSDKkBF1Sy4wONFHSzJK9Mz1Fm5M9GTczBGvfVt8FJCZJFjCY
IWhJNNEgPcRJgliJHCuPQn+VpoNZgn1Oy0AZOqKxJE+QcIGJcj5L4qN+hbwqjeFdwjJLLKyc3Qxu
CbYO6xWjT4fJgJ3p7Cj/W8p9JZ0yMVHBxwjlc9/B9PcA7e+vv5l0I+lZLhO2lWQZ2KvJVJwTsEsx
iZ3lAI9T2SjlM8olrjcxLuLGNuocRm5Ni6ZsU6Wrl1Y1FePdEF8Iyx83VmIUOyJNJ0zrBcNJkd/P
cQ4oE1NeJXKOOOdJgmYPB8fF6+V3+b3YnM2NVH7/DtL9BfU/AQZ0p6pA9rYRebMkKz2kELEeMCWu
N8HtQpCYHmGOkgLT7zb4/gPQSwECFAAUAAAACABtmFdX1kAPJF8CAACvBAAAEAAAAAAAAAAAACAA
AAAAAAAAbWFwZW9fbW9uZWRhcy5jc3ZQSwUGAAAAAAEAAQA+AAAAjQIAAAAA
"""
        
        # Decodificar y leer el CSV
        zip_data = base64.b64decode(mapeo_base64)
        df_mapeo = pd.read_csv(io.BytesIO(zip_data), compression='zip')
        return df_mapeo
    except Exception as e:
        st.error(f"Error cargando mapeo de monedas: {e}")
        return pd.DataFrame()

def agregar_columnas_debe_haber_saldo(df, mapeo_monedas):
    """Agregar columnas DEBE, HABER y SALDO según el estado"""
    try:
        # Crear las nuevas columnas
        df['DEBE'] = 0.0
        df['HABER'] = 0.0
        df['SALDO'] = 0.0
        
        # Procesar cada fila
        for idx, row in df.iterrows():
            try:
                estado = str(row.get('Estado', '')).strip().upper()
                monto = pd.to_numeric(row.get('Monto', 0), errors='coerce')
                
                if pd.isna(monto):
                    monto = 0
                
                # Lógica del estado
                if 'DEBIT' in estado or 'DEBITO' in estado:
                    df.at[idx, 'DEBE'] = abs(monto)
                    df.at[idx, 'SALDO'] = -abs(monto)
                elif 'CREDIT' in estado or 'CREDITO' in estado:
                    df.at[idx, 'HABER'] = abs(monto)
                    df.at[idx, 'SALDO'] = abs(monto)
                else:
                    # Por defecto basado en signo del monto
                    if monto < 0:
                        df.at[idx, 'DEBE'] = abs(monto)
                        df.at[idx, 'SALDO'] = monto
                    else:
                        df.at[idx, 'HABER'] = abs(monto)
                        df.at[idx, 'SALDO'] = monto
                        
            except Exception as e:
                continue
        
        return df
    except Exception as e:
        st.error(f"Error agregando columnas DEBE/HABER/SALDO: {e}")
        return df

def procesar_archivo_caratulas(uploaded_file):
    """Función principal de procesamiento - adaptada del script original"""
    
    progress_bar = st.progress(0, text="Iniciando procesamiento...")
    
    try:
        # Leer archivo Excel
        progress_bar.progress(10, text="Leyendo archivo Excel...")
        xl_file = pd.ExcelFile(uploaded_file)
        hojas_disponibles = xl_file.sheet_names
        
        st.info(f"📋 Hojas encontradas: {len(hojas_disponibles)}")
        for i, hoja in enumerate(hojas_disponibles, 1):
            st.write(f"  {i}. {hoja}")
        
        # Cargar mapeo de monedas
        progress_bar.progress(20, text="Cargando mapeo de monedas...")
        mapeo_monedas = cargar_mapeo_monedas()
        
        if mapeo_monedas.empty:
            st.warning("⚠️ No se pudo cargar el mapeo de monedas, usando valores por defecto")
        
        # Detectar hojas útiles (excluyendo resúmenes)
        progress_bar.progress(30, text="Detectando hojas útiles...")
        hojas_utiles = []
        hojas_excluidas = ["TD Aging", "Aging", "Resumen", "Instrucciones", "Summary", "Template"]
        
        for hoja in hojas_disponibles:
            es_util = True
            for excluir in hojas_excluidas:
                if excluir.lower() in hoja.lower():
                    es_util = False
                    break
            if es_util:
                hojas_utiles.append(hoja)
        
        st.success(f"✅ Hojas útiles detectadas: {len(hojas_utiles)}")
        
        # Procesar cada hoja
        df_consolidado_list = []
        resumen_proceso = []
        estadisticas = {
            'total_hojas_procesadas': 0,
            'total_movimientos': 0,
            'total_debe': 0,
            'total_haber': 0,
            'monedas_encontradas': set()
        }
        
        for i, nombre_hoja in enumerate(hojas_utiles):
            try:
                progress_bar.progress(40 + (i * 40 // len(hojas_utiles)), 
                                    text=f"Procesando hoja: {nombre_hoja}")
                
                # Leer hoja cruda
                df_raw = pd.read_excel(uploaded_file, sheet_name=nombre_hoja, header=None)
                
                if df_raw.empty:
                    continue
                
                # Buscar headers en fila 9 o 10
                headers_row = None
                for row_idx in [9, 10, 11]:
                    if row_idx < len(df_raw):
                        potential_headers = df_raw.iloc[row_idx].astype(str).str.lower()
                        if any('estado' in str(h) for h in potential_headers):
                            headers_row = row_idx
                            break
                
                if headers_row is None:
                    continue
                
                # Extraer headers
                headers = df_raw.iloc[headers_row].astype(str)
                
                # Mapear columnas
                column_mapping = {}
                for idx, header in enumerate(headers):
                    header_lower = str(header).lower()
                    if 'estado' in header_lower:
                        column_mapping['Estado'] = idx
                    elif 'fecha contable' in header_lower:
                        column_mapping['Fecha'] = idx
                    elif 'aging' in header_lower:
                        column_mapping['Aging'] = idx
                    elif 'descripcion' in header_lower or 'descripción' in header_lower:
                        column_mapping['Descripción'] = idx
                    elif 'monto original' in header_lower:
                        column_mapping['Monto'] = idx
                    elif 'responsable' in header_lower:
                        column_mapping['Responsable'] = idx
                    elif 'flex contable' in header_lower:
                        column_mapping['Flex contable'] = idx
                
                # Forzar mapeos fijos
                column_mapping['Numero de transacción'] = 7  # Columna H
                column_mapping['Proveedor/Cliente'] = 11    # Columna L
                
                # Procesar datos
                datos_hoja = []
                for row_idx in range(headers_row + 2, df_raw.shape[0]):
                    try:
                        row = df_raw.iloc[row_idx]
                        
                        # Crear diccionario de datos de la fila
                        row_data = {}
                        for col_name, col_idx in column_mapping.items():
                            if col_idx < len(row):
                                row_data[col_name] = row.iloc[col_idx]
                        
                        # Criterios de validación
                        criterios_cumplidos = 0
                        
                        # Criterio 1: Aging numérico
                        if 'Aging' in row_data:
                            aging = pd.to_numeric(row_data['Aging'], errors='coerce')
                            if not pd.isna(aging) and aging > 0:
                                criterios_cumplidos += 1
                        
                        # Criterio 2: Fecha válida (CUALQUIER año)
                        if 'Fecha' in row_data:
                            fecha = row_data['Fecha']
                            if pd.notna(fecha) and len(str(fecha).strip()) > 0:
                                criterios_cumplidos += 1
                        
                        # Criterio 3: Monto válido
                        if 'Monto' in row_data:
                            monto = pd.to_numeric(row_data['Monto'], errors='coerce')
                            if not pd.isna(monto) and monto != 0:
                                criterios_cumplidos += 1
                        
                        # Criterio 4: Responsable válido
                        if 'Responsable' in row_data:
                            resp = str(row_data['Responsable']).strip()
                            if len(resp) > 2:
                                criterios_cumplidos += 1
                        
                        # Criterio 5: Flex contable válido
                        if 'Flex contable' in row_data:
                            flex = str(row_data['Flex contable'])
                            if '105.' in flex and len(flex) > 10:
                                criterios_cumplidos += 1
                        
                        # Incluir fila si cumple al menos 3 criterios
                        if criterios_cumplidos >= 3:
                            # Asignar moneda basada en Flex banco
                            flex_banco = str(row.iloc[19] if len(row) > 19 else '')
                            moneda = 'USD'  # Por defecto
                            
                            if not mapeo_monedas.empty:
                                for _, mapeo_row in mapeo_monedas.iterrows():
                                    if str(mapeo_row['Flex banco']) in flex_banco:
                                        moneda = mapeo_row['Moneda']
                                        break
                            
                            row_data['Moneda'] = moneda
                            row_data['Hoja_origen'] = nombre_hoja
                            datos_hoja.append(row_data)
                    
                    except Exception as e:
                        continue
                
                if datos_hoja:
                    df_hoja = pd.DataFrame(datos_hoja)
                    df_hoja = agregar_columnas_debe_haber_saldo(df_hoja, mapeo_monedas)
                    df_consolidado_list.append(df_hoja)
                    
                    # Estadísticas
                    estadisticas['total_hojas_procesadas'] += 1
                    estadisticas['total_movimientos'] += len(df_hoja)
                    estadisticas['total_debe'] += df_hoja['DEBE'].sum()
                    estadisticas['total_haber'] += df_hoja['HABER'].sum()
                    estadisticas['monedas_encontradas'].update(df_hoja['Moneda'].unique())
                    
                    resumen_proceso.append({
                        'Hoja': nombre_hoja,
                        'Movimientos': len(df_hoja),
                        'Total_Debe': df_hoja['DEBE'].sum(),
                        'Total_Haber': df_hoja['HABER'].sum()
                    })
            
            except Exception as e:
                st.warning(f"⚠️ Error procesando hoja '{nombre_hoja}': {e}")
                continue
        
        progress_bar.progress(90, text="Consolidando resultados...")
        
        # Consolidar todos los datos
        if df_consolidado_list:
            df_final = pd.concat(df_consolidado_list, ignore_index=True)
            df_resumen = pd.DataFrame(resumen_proceso)
            
            # Preparar estadísticas finales
            estadisticas_final = pd.DataFrame([{
                'Métrica': 'Hojas procesadas',
                'Valor': estadisticas['total_hojas_procesadas']
            }, {
                'Métrica': 'Total movimientos',
                'Valor': estadisticas['total_movimientos']
            }, {
                'Métrica': 'Total DEBE',
                'Valor': f"{estadisticas['total_debe']:,.2f}"
            }, {
                'Métrica': 'Total HABER', 
                'Valor': f"{estadisticas['total_haber']:,.2f}"
            }, {
                'Métrica': 'Monedas encontradas',
                'Valor': ', '.join(sorted(estadisticas['monedas_encontradas']))
            }])
            
            progress_bar.progress(100, text="¡Procesamiento completado!")
            
            return df_final, df_resumen, estadisticas_final
        
        else:
            st.error("❌ No se encontraron datos válidos para procesar")
            return None, None, None
            
    except Exception as e:
        st.error(f"❌ Error durante el procesamiento: {e}")
        return None, None, None

def main():
    """Función principal de la aplicación Streamlit"""
    
    # Título y descripción
    st.title("🏦 Procesador de Carátulas Bancarias")
    st.markdown("---")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### 📋 Funcionalidades:
        ✅ **Procesa fechas de cualquier año** (sin restricciones)  
        ✅ **Detecta automáticamente** las hojas útiles  
        ✅ **Mapea columnas** automáticamente  
        ✅ **Asigna monedas** según Flex banco  
        ✅ **Calcula DEBE/HABER/SALDO** automáticamente  
        ✅ **Genera archivo consolidado** con múltiples hojas  
        """)
    
    with col2:
        st.info("""
        **📤 Formato de entrada:**  
        Archivo Excel (.xlsx/.xlsm)  
        con carátulas bancarias  
        
        **📥 Resultado:**  
        Excel con 3 hojas:  
        - Datos Consolidados  
        - Resumen por Hoja  
        - Estadísticas Generales  
        """)
    
    st.markdown("---")
    
    # Carga de archivo
    uploaded_file = st.file_uploader(
        "📁 Selecciona tu archivo de carátulas bancarias",
        type=['xlsx', 'xlsm', 'xls'],
        help="Formatos soportados: .xlsx, .xlsm, .xls"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ Archivo cargado: **{uploaded_file.name}**")
        
        # Mostrar información del archivo
        file_size = len(uploaded_file.getvalue()) / 1024 / 1024  # MB
        st.info(f"📊 Tamaño del archivo: **{file_size:.2f} MB**")
        
        # Botón para procesar
        if st.button("🚀 Procesar Carátulas", type="primary", use_container_width=True):
            
            with st.spinner("Procesando archivo... Esto puede tomar unos minutos."):
                # Procesar archivo
                df_consolidado, df_resumen, df_estadisticas = procesar_archivo_caratulas(uploaded_file)
                
                if df_consolidado is not None:
                    st.success("🎉 ¡Procesamiento completado exitosamente!")
                    
                    # Mostrar resultados
                    tab1, tab2, tab3, tab4 = st.tabs(["📊 Resumen", "📋 Datos", "📈 Estadísticas", "⬇️ Descargar"])
                    
                    with tab1:
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Movimientos Totales", len(df_consolidado))
                        with col2:
                            st.metric("Total DEBE", f"{df_consolidado['DEBE'].sum():,.2f}")
                        with col3:
                            st.metric("Total HABER", f"{df_consolidado['HABER'].sum():,.2f}")
                        
                        st.subheader("📋 Resumen por Hoja")
                        st.dataframe(df_resumen, use_container_width=True)
                    
                    with tab2:
                        st.subheader("🗂️ Datos Consolidados")
                        st.dataframe(df_consolidado, use_container_width=True)
                    
                    with tab3:
                        st.subheader("📈 Estadísticas Generales")
                        st.dataframe(df_estadisticas, use_container_width=True)
                    
                    with tab4:
                        st.subheader("⬇️ Descargar Resultados")
                        
                        # Crear archivo Excel con múltiples hojas
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_consolidado.to_excel(writer, sheet_name='Datos_Consolidados', index=False)
                            df_resumen.to_excel(writer, sheet_name='Resumen_Proceso', index=False)
                            df_estadisticas.to_excel(writer, sheet_name='Estadisticas', index=False)
                        
                        # Generar nombre de archivo
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"caratulas_vaciado_{timestamp}.xlsx"
                        
                        st.download_button(
                            label="📥 Descargar Excel Procesado",
                            data=output.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                        
                        st.success(f"✅ Archivo listo para descarga: **{filename}**")
                else:
                    st.error("❌ No se pudieron procesar las carátulas. Verifica el formato del archivo.")

# Sidebar con información adicional
with st.sidebar:
    st.markdown("### ℹ️ Información")
    
    st.markdown("""
    **🔧 Versión:** Universal  
    **📅 Fechas:** Cualquier año  
    **🏦 Bancos:** Todos los soportados  
    **💱 Monedas:** Mapeo automático  
    """)
    
    st.markdown("---")
    
    with st.expander("🔍 Detalles Técnicos"):
        st.markdown("""
        **Criterios de Validación:**
        - Aging numérico válido
        - Fecha no nula 
        - Monto diferente de cero
        - Responsable válido
        - Flex contable válido
        
        **Mapeo Automático:**
        - Columna H → Número de transacción
        - Columna L → Proveedor/Cliente
        - Moneda → Según Flex banco
        """)
    
    with st.expander("❓ Ayuda"):
        st.markdown("""
        **Formato de archivo esperado:**
        - Excel con múltiples hojas
        - Headers en fila 9-11
        - Datos desde fila 12+
        
        **¿Problemas?**
        - Verifica que el archivo no esté protegido
        - Asegúrate que contenga datos válidos
        - Los headers deben estar en español
        """)

if __name__ == "__main__":
    main()
