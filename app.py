import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import base64
import warnings
import io
import os
import re

# Configurar warnings
warnings.filterwarnings('ignore')

# Excel embebido con mapeo de monedas (base64)
EXCEL_MONEDA_BASE64 = """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"""

def cargar_mapeo_monedas():
    """Carga el mapeo de monedas desde el Excel embebido."""
    try:
        excel_data = base64.b64decode(EXCEL_MONEDA_BASE64)
        df_monedas = pd.read_excel(io.BytesIO(excel_data), engine='openpyxl')
        
        # Crear diccionario de mapeo: {Flex Efectivo: Moneda}
        mapeo = {}
        for _, row in df_monedas.iterrows():
            flex_efectivo = str(row.get('Flex Efectivo', '')).strip()
            moneda = str(row.get('Moneda', '')).strip()
            if flex_efectivo and moneda:
                mapeo[flex_efectivo] = moneda
        
        return mapeo
    except Exception as e:
        st.error(f"Error al cargar mapeo de monedas: {str(e)}")
        return {}

def agregar_columnas_debe_haber_saldo(df):
    """Agrega las columnas DEBE, HABER y SALDO al DataFrame."""
    try:
        df = df.copy()
        
        # Asegurarse de que la columna Monto esté como numérica
        df['Monto'] = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)
        
        # Inicializar columnas
        df['DEBE'] = 0.0
        df['HABER'] = 0.0
        df['SALDO'] = df['Monto']
        
        # Lógica para DEBE y HABER basada en el signo del monto
        mask_positivo = df['Monto'] > 0
        mask_negativo = df['Monto'] < 0
        
        df.loc[mask_positivo, 'DEBE'] = df.loc[mask_positivo, 'Monto']
        df.loc[mask_negativo, 'HABER'] = abs(df.loc[mask_negativo, 'Monto'])
        
        # Reordenar columnas según el orden esperado
        columnas_deseadas = [
            'Estado', 'Aging', 'Fecha', 'Categoría', 'Numero de transacción',
            'Proveedor/Cliente', 'Monto', 'Concepto', 'Responsable',
            'Flex contable', 'Flex banco', 'Moneda', 'BANCO', 'DEBE', 'HABER', 'SALDO'
        ]
        
        # Solo incluir columnas que existan en el DataFrame
        columnas_finales = [col for col in columnas_deseadas if col in df.columns]
        df = df[columnas_finales]
        
        return df
    except Exception as e:
        st.error(f"Error al agregar columnas DEBE/HABER/SALDO: {str(e)}")
        return df

def procesar_archivo_excel(file, progreso, log_container):
    """Función principal de procesamiento del archivo Excel."""
    try:
        # Cargar mapeo de monedas
        mapeo_monedas = cargar_mapeo_monedas()
        log_container.text("✅ Mapeo de monedas cargado correctamente")
        progreso.progress(0.1)
        
        # Leer todas las hojas del archivo Excel
        hojas_excel = pd.read_excel(file, sheet_name=None, engine='openpyxl')
        log_container.text(f"📄 Encontradas {len(hojas_excel)} hojas en el archivo")
        progreso.progress(0.2)
        
        df_consolidado = pd.DataFrame()
        resumen_procesamiento = []
        total_registros = 0
        bancos_procesados = 0
        
        for nombre_hoja, df_hoja in hojas_excel.items():
            try:
                log_container.text(f"🔄 Procesando hoja: {nombre_hoja}")
                
                # Filtrar filas válidas (que tengan datos en columnas clave)
                df_valido = df_hoja.dropna(how='all')
                
                if df_valido.empty:
                    resumen_procesamiento.append({
                        'banco': nombre_hoja,
                        'registros': 0,
                        'estado': 'Vacía'
                    })
                    continue
                
                # Mapeo de columnas
                mapeo_columnas = {}
                
                # Buscar columnas por nombre/patrón
                for idx, col_name in enumerate(df_valido.columns):
                    col_str = str(col_name).lower().strip()
                    if 'estado' in col_str:
                        mapeo_columnas['Estado'] = idx
                    elif 'aging' in col_str:
                        mapeo_columnas['Aging'] = idx
                    elif 'fecha' in col_str:
                        mapeo_columnas['Fecha'] = idx
                    elif 'categoría' in col_str or 'categoria' in col_str:
                        mapeo_columnas['Categoría'] = idx
                    elif 'monto' in col_str:
                        mapeo_columnas['Monto'] = idx
                    elif 'concepto' in col_str:
                        mapeo_columnas['Concepto'] = idx
                    elif 'responsable' in col_str:
                        mapeo_columnas['Responsable'] = idx
                    elif 'flex contable' in col_str:
                        mapeo_columnas['Flex contable'] = idx
                    elif 'flex banco' in col_str:
                        mapeo_columnas['Flex banco'] = idx
                
                # Mapeos fijos como en el script original
                mapeo_columnas['Numero de transacción'] = 7  # Columna H
                mapeo_columnas['Proveedor/Cliente'] = 11     # Columna L
                
                # Crear DataFrame con mapeo
                datos_procesados = []
                for idx, fila in df_valido.iterrows():
                    row_data = fila.tolist()
                    
                    # Verificar si la fila tiene datos válidos
                    if 'Aging' in mapeo_columnas:
                        aging = pd.to_numeric(row_data[mapeo_columnas['Aging']], errors='coerce')
                        if pd.isna(aging) or aging < 0:
                            continue
                    
                    # Crear registro
                    registro = {}
                    for col_destino, col_idx in mapeo_columnas.items():
                        if col_idx < len(row_data):
                            registro[col_destino] = row_data[col_idx]
                        else:
                            registro[col_destino] = None
                    
                    # Agregar nombre del banco
                    registro['BANCO'] = nombre_hoja
                    
                    # Mapear moneda usando Flex banco
                    flex_banco = str(registro.get('Flex banco', '')).strip()
                    registro['Moneda'] = mapeo_monedas.get(flex_banco, 'USD')  # Default USD
                    
                    datos_procesados.append(registro)
                
                if datos_procesados:
                    df_hoja_procesada = pd.DataFrame(datos_procesados)
                    df_consolidado = pd.concat([df_consolidado, df_hoja_procesada], ignore_index=True)
                    
                    total_registros += len(datos_procesados)
                    bancos_procesados += 1
                    
                    resumen_procesamiento.append({
                        'banco': nombre_hoja,
                        'registros': len(datos_procesados),
                        'estado': 'Exitoso'
                    })
                else:
                    resumen_procesamiento.append({
                        'banco': nombre_hoja,
                        'registros': 0,
                        'estado': 'Sin datos válidos'
                    })
                
                progreso.progress(0.2 + (0.6 * (bancos_procesados + 1) / len(hojas_excel)))
                
            except Exception as e:
                resumen_procesamiento.append({
                    'banco': nombre_hoja,
                    'registros': 0,
                    'estado': f'Error: {str(e)[:50]}...'
                })
                log_container.text(f"❌ Error procesando {nombre_hoja}: {str(e)[:100]}...")
        
        # Agregar columnas DEBE, HABER, SALDO
        if not df_consolidado.empty:
            df_consolidado = agregar_columnas_debe_haber_saldo(df_consolidado)
            log_container.text("✅ Columnas DEBE/HABER/SALDO agregadas")
        
        progreso.progress(0.9)
        
        # Crear Excel de salida
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"caratulas_vaciado_{timestamp}.xlsx"
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Hoja de datos consolidados
            if not df_consolidado.empty:
                df_consolidado.to_excel(writer, sheet_name='Datos_Consolidados', index=False)
            
            # Hoja de resumen de procesamiento
            df_resumen = pd.DataFrame(resumen_procesamiento)
            df_resumen.to_excel(writer, sheet_name='Resumen_Proceso', index=False)
            
            # Hoja de estadísticas
            estadisticas = {
                'Total de registros': [total_registros],
                'Bancos procesados': [bancos_procesados],
                'Fecha de procesamiento': [datetime.now().strftime("%d/%m/%Y %H:%M:%S")]
            }
            df_stats = pd.DataFrame(estadisticas)
            df_stats.to_excel(writer, sheet_name='Estadisticas', index=False)
        
        progreso.progress(1.0)
        log_container.text(f"✅ Procesamiento completado. Total registros: {total_registros}")
        
        output.seek(0)
        return output, nombre_archivo, total_registros, bancos_procesados
        
    except Exception as e:
        log_container.text(f"❌ Error general: {str(e)}")
        raise e

def main():
    st.set_page_config(
        page_title="Vaciado de Carátulas Bancarias",
        page_icon="🏦",
        layout="wide"
    )
    
    st.title("🏦 Vaciado de Carátulas Bancarias")
    st.markdown("---")
    
    st.markdown("""
    ### Instrucciones:
    1. Seleccione el archivo Excel con las carátulas bancarias
    2. El archivo debe contener múltiples hojas (una por banco)
    3. Cada hoja debe tener las columnas estándar de carátulas
    4. Haga clic en "Procesar Archivo" para iniciar
    5. Descargue el archivo consolidado cuando esté listo
    """)
    
    # Sidebar para controles
    with st.sidebar:
        st.header("📁 Selección de Archivo")
        uploaded_file = st.file_uploader(
            "Cargar archivo Excel",
            type=['xlsx', 'xlsm', 'xls'],
            help="Seleccione el archivo Excel con las carátulas bancarias"
        )
        
        if uploaded_file:
            st.success(f"Archivo cargado: {uploaded_file.name}")
            
            # Información del archivo
            st.info(f"Tamaño: {len(uploaded_file.getvalue()} bytes")
    
    # Área principal
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if uploaded_file is not None:
            if st.button("🚀 Procesar Archivo", type="primary", use_container_width=True):
                with st.spinner("Procesando archivo..."):
                    # Contenedor para progreso y logs
                    progreso = st.progress(0)
                    log_container = st.empty()
                    
                    try:
                        # Procesar archivo
                        archivo_salida, nombre_archivo, total_registros, bancos_procesados = procesar_archivo_excel(
                            uploaded_file, progreso, log_container
                        )
                        
                        # Mostrar resultados
                        st.success("✅ Procesamiento completado exitosamente!")
                        
                        col_res1, col_res2, col_res3 = st.columns(3)
                        with col_res1:
                            st.metric("Registros Procesados", total_registros)
                        with col_res2:
                            st.metric("Bancos Procesados", bancos_procesados)
                        with col_res3:
                            st.metric("Archivo Generado", nombre_archivo.split('_')[2])
                        
                        # Botón de descarga
                        st.download_button(
                            label="📥 Descargar Archivo Consolidado",
                            data=archivo_salida.getvalue(),
                            file_name=nombre_archivo,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                    except Exception as e:
                        st.error(f"❌ Error durante el procesamiento: {str(e)}")
                        st.error("Por favor, verifique el formato del archivo y vuelva a intentar.")
        else:
            st.info("👆 Por favor, cargue un archivo Excel para comenzar el procesamiento.")
    
    with col2:
        st.header("📊 Información del Proceso")
        with st.expander("Detalles del Mapeo"):
            st.markdown("""
            **Columnas mapeadas automáticamente:**
            - Estado, Aging, Fecha, Categoría
            - Monto, Concepto, Responsable
            - Flex contable, Flex banco
            
            **Columnas con mapeo fijo:**
            - Número de transacción: Columna H (8)
            - Proveedor/Cliente: Columna L (12)
            
            **Columnas agregadas:**
            - BANCO (nombre de la hoja)
            - Moneda (mapeada desde Flex banco)
            - DEBE, HABER, SALDO (calculadas)
            """)
        
        with st.expander("Formato del Archivo"):
            st.markdown("""
            **Estructura esperada:**
            - Archivo Excel (.xlsx, .xlsm, .xls)
            - Múltiples hojas (una por banco)
            - Cada hoja con columnas estándar
            - Datos a partir de la fila 1
            
            **Salida generada:**
            - Datos_Consolidados: Todos los registros
            - Resumen_Proceso: Estado por banco
            - Estadísticas: Métricas generales
            """)

if __name__ == "__main__":
    main()
