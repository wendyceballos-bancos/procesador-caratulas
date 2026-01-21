
import streamlit as st
import pandas as pd
import numpy as np
import base64
import io
import os
from datetime import datetime
import warnings
import re

warnings.filterwarnings('ignore')

# Configuración de la página
st.set_page_config(
    page_title="🏦 Vaciado de Carátulas Bancarias",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

def cargar_mapeo_monedas():
    """Carga el mapeo de monedas"""
    mapeo_default = {
        'USD': 'USD', 'EUR': 'EUR', 'GBP': 'GBP', 'CAD': 'CAD', 'AUD': 'AUD',
        'JPY': 'JPY', 'CHF': 'CHF', 'SEK': 'SEK', 'NOK': 'NOK', 'DKK': 'DKK',
        'PLN': 'PLN', 'CZK': 'CZK', 'HUF': 'HUF', 'RON': 'RON', 'BGN': 'BGN',
        'HRK': 'HRK', 'RSD': 'RSD', 'TRY': 'TRY', 'RUB': 'RUB', 'UAH': 'UAH',
        'BYN': 'BYN', 'CNY': 'CNY', 'KRW': 'KRW', 'INR': 'INR', 'SGD': 'SGD',
        'HKD': 'HKD', 'TWD': 'TWD', 'THB': 'THB', 'MYR': 'MYR', 'IDR': 'IDR',
        'PHP': 'PHP', 'VND': 'VND', 'BRL': 'BRL', 'ARS': 'ARS', 'CLP': 'CLP',
        'COP': 'COP', 'PEN': 'PEN', 'MXN': 'MXN', 'ZAR': 'ZAR', 'EGP': 'EGP'
    }
    return mapeo_default

def limpiar_nombre_banco(nombre_hoja):
    """Extrae el nombre del banco desde el nombre de la hoja"""
    nombre_limpio = re.sub(r'\d+$', '', nombre_hoja)
    nombre_limpio = re.sub(r'\b(datos|data|info|información)\b', '', nombre_limpio, flags=re.IGNORECASE)
    nombre_limpio = nombre_limpio.strip().strip('_-. ')
    return nombre_limpio if nombre_limpio else nombre_hoja

def obtener_moneda_de_flex(flex_banco, mapeo_monedas):
    """Obtiene la moneda basada en el flex banco"""
    if pd.isna(flex_banco) or flex_banco == '':
        return ''
    
    flex_str = str(flex_banco).upper()
    for codigo, moneda in mapeo_monedas.items():
        if codigo in flex_str:
            return moneda
    return ''

def procesar_archivo_excel(archivo, progress_callback=None, log_callback=None):
    """Procesa el archivo Excel y devuelve los datos consolidados"""
    
    def log(mensaje):
        if log_callback:
            log_callback(mensaje)
        print(mensaje)
    
    def actualizar_progreso(valor):
        if progress_callback:
            progress_callback(valor)
    
    try:
        log("🔍 Iniciando procesamiento del archivo...")
        
        # Cargar el mapeo de monedas
        mapeo_monedas = cargar_mapeo_monedas()
        log(f"✅ Mapeo de monedas cargado: {len(mapeo_monedas)} monedas disponibles")
        
        # Leer todas las hojas del archivo
        try:
            todas_las_hojas = pd.read_excel(archivo, sheet_name=None, header=None, engine='openpyxl')
            log(f"📄 Archivo cargado: {len(todas_las_hojas)} hojas encontradas")
        except Exception as e:
            log(f"❌ Error al leer el archivo: {e}")
            return None, None, None
        
        # Filtrar hojas que no son de datos
        hojas_excluidas = ['resumen', 'summary', 'índice', 'index', 'instrucciones', 'instructions', 'totales', 'total']
        hojas_datos = []
        
        for nombre_hoja in todas_las_hojas.keys():
            if not any(excl.lower() in nombre_hoja.lower() for excl in hojas_excluidas):
                hojas_datos.append(nombre_hoja)
        
        log(f"📊 Hojas de datos detectadas: {hojas_datos}")
        actualizar_progreso(10)
        
        # Procesar cada hoja
        datos_consolidados = []
        resumen_proceso = []
        total_hojas = len(hojas_datos)
        
        for i, nombre_hoja in enumerate(hojas_datos):
            try:
                log(f"\n🔄 Procesando hoja '{nombre_hoja}'...")
                
                df_hoja = todas_las_hojas[nombre_hoja]
                
                if len(df_hoja) < 13:
                    log(f"⚠️  Hoja '{nombre_hoja}' tiene pocas filas ({len(df_hoja)}), omitiendo...")
                    resumen_proceso.append({
                        'Banco': limpiar_nombre_banco(nombre_hoja),
                        'Hoja': nombre_hoja,
                        'Registros': 0,
                        'Estado': 'Omitida - Pocas filas'
                    })
                    continue
                
                # Extraer encabezados desde la fila 11 (índice 10)
                encabezados = df_hoja.iloc[10].fillna('').astype(str)
                
                for j, enc in enumerate(encabezados):
                    if enc == '' or enc == 'nan':
                        encabezados.iloc[j] = f'Col_{j}'
                
                # Mapear columnas por nombre
                mapeo_columnas = {}
                for idx, nombre in enumerate(encabezados):
                    nombre_lower = str(nombre).lower().strip()
                    
                    if 'estado' in nombre_lower:
                        mapeo_columnas['Estado'] = idx
                    elif 'aging' in nombre_lower:
                        mapeo_columnas['Aging'] = idx
                    elif 'fecha' in nombre_lower and 'contable' in nombre_lower:
                        mapeo_columnas['Fecha'] = idx
                    elif 'fecha' in nombre_lower and 'transac' in nombre_lower:
                        mapeo_columnas['Fecha transacción'] = idx
                    elif 'categoría' in nombre_lower or 'categoria' in nombre_lower:
                        mapeo_columnas['Categoría'] = idx
                    elif 'monto' in nombre_lower and ('funcional' in nombre_lower or 'total' in nombre_lower):
                        mapeo_columnas['Monto'] = idx
                    elif 'concepto' in nombre_lower:
                        mapeo_columnas['Concepto'] = idx
                    elif 'responsable' in nombre_lower:
                        mapeo_columnas['Responsable'] = idx
                    elif 'flex' in nombre_lower and 'contable' in nombre_lower:
                        mapeo_columnas['Flex contable'] = idx
                    elif 'flex' in nombre_lower and ('banco' in nombre_lower or 'efectivo' in nombre_lower):
                        mapeo_columnas['Flex banco'] = idx
                    elif 'tipo' in nombre_lower and 'extracto' in nombre_lower:
                        mapeo_columnas['Tipo extracto'] = idx
                
                # Mapeos fijos por índice
                if len(encabezados) > 7:
                    mapeo_columnas['Numero de transacción'] = 7
                if len(encabezados) > 11:
                    mapeo_columnas['Proveedor/Cliente'] = 11
                
                log(f"📋 Columnas mapeadas: {list(mapeo_columnas.keys())}")
                
                # Procesar filas desde el índice 12
                filas_validas = []
                
                for fila_idx in range(12, len(df_hoja)):
                    fila = df_hoja.iloc[fila_idx]
                    criterios_cumplidos = 0
                    
                    # Criterio 1: Aging numérico > 0
                    if 'Aging' in mapeo_columnas:
                        try:
                            aging_val = pd.to_numeric(fila.iloc[mapeo_columnas['Aging']], errors='coerce')
                            if not pd.isna(aging_val) and aging_val > 0:
                                criterios_cumplidos += 1
                        except:
                            pass
                    
                    # Criterio 2: Fecha no vacía
                    fecha_encontrada = False
                    for col_fecha in ['Fecha', 'Fecha transacción']:
                        if col_fecha in mapeo_columnas:
                            fecha_val = str(fila.iloc[mapeo_columnas[col_fecha]])
                            if fecha_val and fecha_val != 'nan' and fecha_val.strip():
                                fecha_encontrada = True
                                break
                    if fecha_encontrada:
                        criterios_cumplidos += 1
                    
                    # Criterio 3: Monto numérico distinto de 0
                    if 'Monto' in mapeo_columnas:
                        try:
                            monto_val = pd.to_numeric(fila.iloc[mapeo_columnas['Monto']], errors='coerce')
                            if not pd.isna(monto_val) and monto_val != 0:
                                criterios_cumplidos += 1
                        except:
                            pass
                    
                    # Criterio 4: Responsable con longitud > 2
                    if 'Responsable' in mapeo_columnas:
                        responsable_val = str(fila.iloc[mapeo_columnas['Responsable']])
                        if responsable_val and responsable_val != 'nan' and len(responsable_val.strip()) > 2:
                            criterios_cumplidos += 1
                    
                    # Criterio 5: Flex contable que contenga '105.' y longitud > 10
                    if 'Flex contable' in mapeo_columnas:
                        flex_cont_val = str(fila.iloc[mapeo_columnas['Flex contable']])
                        if flex_cont_val and flex_cont_val != 'nan' and '105.' in flex_cont_val and len(flex_cont_val.strip()) > 10:
                            criterios_cumplidos += 1
                    
                    if criterios_cumplidos >= 3:
                        filas_validas.append(fila_idx)
                
                if not filas_validas:
                    log(f"⚠️  No se encontraron filas válidas en '{nombre_hoja}'")
                    resumen_proceso.append({
                        'Banco': limpiar_nombre_banco(nombre_hoja),
                        'Hoja': nombre_hoja,
                        'Registros': 0,
                        'Estado': 'Sin datos válidos'
                    })
                    continue
                
                # Crear DataFrame con las filas válidas
                df_procesado = pd.DataFrame()
                
                columnas_objetivo = [
                    'Estado', 'Aging', 'Fecha', 'Fecha transacción', 'Categoría', 
                    'Numero de transacción', 'Proveedor/Cliente', 'Monto', 'Concepto', 
                    'Responsable', 'Flex contable', 'Flex banco', 'Tipo extracto',
                    'Moneda', 'BANCO'
                ]
                
                for col in columnas_objetivo[:-2]:
                    if col in mapeo_columnas:
                        valores = [df_hoja.iloc[idx, mapeo_columnas[col]] if idx < len(df_hoja) and mapeo_columnas[col] < len(df_hoja.columns) else '' for idx in filas_validas]
                        df_procesado[col] = valores
                    else:
                        df_procesado[col] = [''] * len(filas_validas)
                
                # Agregar columna de moneda
                if 'Flex banco' in mapeo_columnas:
                    df_procesado['Moneda'] = df_procesado['Flex banco'].apply(
                        lambda x: obtener_moneda_de_flex(x, mapeo_monedas)
                    )
                else:
                    df_procesado['Moneda'] = ''
                
                # Agregar columna de banco
                df_procesado['BANCO'] = limpiar_nombre_banco(nombre_hoja)
                
                # Limpiar datos
                for col in df_procesado.columns:
                    if df_procesado[col].dtype == 'object':
                        df_procesado[col] = df_procesado[col].astype(str).str.strip()
                        df_procesado[col] = df_procesado[col].replace('nan', '')
                
                datos_consolidados.append(df_procesado)
                
                log(f"✅ Hoja '{nombre_hoja}' procesada: {len(filas_validas)} registros válidos")
                resumen_proceso.append({
                    'Banco': limpiar_nombre_banco(nombre_hoja),
                    'Hoja': nombre_hoja,
                    'Registros': len(filas_validas),
                    'Estado': 'Procesada correctamente'
                })
                
                progreso = 10 + (70 * (i + 1) / total_hojas)
                actualizar_progreso(int(progreso))
                
            except Exception as e:
                log(f"❌ Error procesando hoja '{nombre_hoja}': {e}")
                resumen_proceso.append({
                    'Banco': limpiar_nombre_banco(nombre_hoja),
                    'Hoja': nombre_hoja,
                    'Registros': 0,
                    'Estado': f'Error: {str(e)[:50]}'
                })
        
        if not datos_consolidados:
            log("❌ No se pudo procesar ninguna hoja")
            return None, None, None
        
        # Consolidar todos los datos
        log("\n🔗 Consolidando datos...")
        df_final = pd.concat(datos_consolidados, ignore_index=True)
        
        # Limpiar datos finales
        for col in df_final.columns:
            if df_final[col].dtype == 'object':
                df_final[col] = df_final[col].astype(str).str.strip()
                df_final[col] = df_final[col].replace('nan', '')
                if col not in ['Proveedor/Cliente', 'Moneda']:
                    df_final[col] = df_final[col].replace('', np.nan)
        
        # Agregar columnas DEBE/HABER/SALDO
        log("💰 Calculando DEBE, HABER y SALDO...")
        actualizar_progreso(85)
        
        col_estado = None
        col_monto = None
        col_tipo_extracto = None
        
        for col in df_final.columns:
            if 'estado' in col.lower():
                col_estado = col
            if 'monto' in col.lower():
                col_monto = col
            if 'tipo' in col.lower() and 'extracto' in col.lower():
                col_tipo_extracto = col
        
        log(f"🔍 Columnas identificadas - Estado: {col_estado}, Monto: {col_monto}, Tipo extracto: {col_tipo_extracto}")
        
        if col_estado and col_monto:
            df_final['DEBE'] = 0.0
            df_final['HABER'] = 0.0
            
            # Textos que indican DEBE (exactos según tu archivo)
            textos_debe = [
                "III. Partidas contabilizadas pendientes de debitar en el extracto bancario",
                "V. Partidas acreditadas en el extracto bancario, pendientes de contabilizar"
            ]
            
            # Textos que indican HABER (exactos según tu archivo)
            textos_haber = [
                "II. Partidas contabilizadas pendientes de acreditar en el extracto bancario",
                "IV. Partidas debitadas en el extracto bancario, pendientes de contabilizar"
            ]
            
            debe_count = 0
            haber_count = 0
            
            for idx, row in df_final.iterrows():
                estado_text = str(row[col_estado]).strip()
                tipo_extracto = str(row.get(col_tipo_extracto, '')).strip().upper() if col_tipo_extracto else ''
                
                try:
                    monto_val = pd.to_numeric(row[col_monto], errors='coerce')
                    if pd.isna(monto_val):
                        monto_val = 0
                except:
                    monto_val = 0
                
                # Priorizar clasificación por tipo de extracto si está disponible
                if tipo_extracto == 'DEBIT':
                    df_final.at[idx, 'DEBE'] = abs(monto_val)
                    df_final.at[idx, 'HABER'] = 0.0
                    debe_count += 1
                elif tipo_extracto == 'CREDIT':
                    df_final.at[idx, 'HABER'] = abs(monto_val)
                    df_final.at[idx, 'DEBE'] = 0.0
                    haber_count += 1
                # Si no hay tipo extracto, usar textos del estado
                elif any(texto in estado_text for texto in textos_debe):
                    df_final.at[idx, 'DEBE'] = abs(monto_val)
                    df_final.at[idx, 'HABER'] = 0.0
                    debe_count += 1
                elif any(texto in estado_text for texto in textos_haber):
                    df_final.at[idx, 'HABER'] = abs(monto_val)
                    df_final.at[idx, 'DEBE'] = 0.0
                    haber_count += 1
                else:
                    # Por defecto, usar el signo del monto
                    if monto_val > 0:
                        df_final.at[idx, 'DEBE'] = monto_val
                        df_final.at[idx, 'HABER'] = 0.0
                        debe_count += 1
                    elif monto_val < 0:
                        df_final.at[idx, 'HABER'] = abs(monto_val)
                        df_final.at[idx, 'DEBE'] = 0.0
                        haber_count += 1
            
            # Calcular SALDO
            df_final['SALDO'] = df_final['DEBE'] - df_final['HABER']
            
            log(f"💰 Cálculo completado: {debe_count} registros en DEBE, {haber_count} en HABER")
            
        else:
            log("⚠️  No se pudieron identificar columnas de Estado o Monto para calcular DEBE/HABER")
            df_final['DEBE'] = 0.0
            df_final['HABER'] = 0.0
            df_final['SALDO'] = 0.0
        
        # Reordenar columnas
        columnas_finales = [
            'Estado', 'Aging', 'Fecha', 'Fecha transacción', 'Categoría', 
            'Numero de transacción', 'Proveedor/Cliente', 'Monto', 'Concepto', 
            'Responsable', 'Flex contable', 'Flex banco', 'Tipo extracto',
            'Moneda', 'BANCO', 'DEBE', 'HABER', 'SALDO'
        ]
        
        for col in df_final.columns:
            if col not in columnas_finales:
                columnas_finales.append(col)
        
        df_final = df_final[[col for col in columnas_finales if col in df_final.columns]]
        
        # Crear DataFrames de resumen
        df_resumen = pd.DataFrame(resumen_proceso)
        
        df_estadisticas = pd.DataFrame({
            'Métrica': [
                'Total de registros procesados',
                'Número de bancos procesados', 
                'Total DEBE',
                'Total HABER',
                'SALDO FINAL',
                'Fecha de procesamiento',
                'Archivo procesado'
            ],
            'Valor': [
                len(df_final),
                len(df_final['BANCO'].unique()) if 'BANCO' in df_final.columns else 0,
                f"{df_final['DEBE'].sum():,.2f}" if 'DEBE' in df_final.columns else 0,
                f"{df_final['HABER'].sum():,.2f}" if 'HABER' in df_final.columns else 0,
                f"{df_final['SALDO'].sum():,.2f}" if 'SALDO' in df_final.columns else 0,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Archivo subido'
            ]
        })
        
        actualizar_progreso(100)
        log(f"\n🎉 ¡Procesamiento completado exitosamente!")
        log(f"📊 Total de registros: {len(df_final)}")
        log(f"🏦 Bancos procesados: {len(df_final['BANCO'].unique()) if 'BANCO' in df_final.columns else 0}")
        log(f"💰 Total DEBE: {df_final['DEBE'].sum():,.2f}")
        log(f"💰 Total HABER: {df_final['HABER'].sum():,.2f}")
        log(f"💰 SALDO: {df_final['SALDO'].sum():,.2f}")
        
        return df_final, df_resumen, df_estadisticas
        
    except Exception as e:
        log(f"❌ Error crítico durante el procesamiento: {e}")
        return None, None, None

def main():
    """Función principal de la aplicación Streamlit"""
    
    # Título y descripción
    st.title("🏦 Vaciado de Carátulas Bancarias")
    st.markdown("---")
    st.markdown("""
    ### 📋 Instrucciones de uso:
    1. **Sube tu archivo Excel** con las carátulas bancarias
    2. **Haz clic en 'Procesar archivo'** para iniciar el análisis
    3. **Descarga los resultados** en formato Excel
    
    ⚠️  **Importante**: El archivo debe tener encabezados en la fila 11 y datos a partir de la fila 13.
    """)
    
    # Sidebar para configuraciones
    st.sidebar.header("⚙️ Configuración")
    st.sidebar.info("""
    **Formato esperado:**
    - Encabezados en fila 11
    - Datos desde fila 13
    - Columnas: Estado, Aging, Fecha, Monto, etc.
    """)
    
    # Upload de archivo
    archivo_subido = st.file_uploader(
        "📁 Selecciona el archivo Excel con las carátulas",
        type=['xlsx', 'xlsm', 'xls'],
        help="Formatos soportados: .xlsx, .xlsm, .xls"
    )
    
    if archivo_subido is not None:
        st.success(f"✅ Archivo cargado: {archivo_subido.name}")
        
        # Mostrar información del archivo
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Nombre", archivo_subido.name)
        with col2:
            st.metric("📦 Tamaño", f"{archivo_subido.size / 1024:.1f} KB")
        with col3:
            st.metric("📋 Tipo", archivo_subido.type)
        
        # Botón para procesar
        if st.button("🚀 Procesar archivo", type="primary"):
            
            # Contenedores para progreso y logs
            progress_container = st.container()
            log_container = st.container()
            
            with progress_container:
                st.subheader("📊 Progreso del procesamiento")
                progress_bar = st.progress(0)
                status_text = st.empty()
            
            with log_container:
                st.subheader("📝 Registro de actividad")
                log_area = st.empty()
            
            # Variables para logs
            logs = []
            
            def actualizar_progreso(valor):
                progress_bar.progress(valor)
                status_text.text(f"Progreso: {valor}%")
            
            def agregar_log(mensaje):
                logs.append(mensaje)
                log_area.text_area(
                    "Logs:", 
                    value="\n".join(logs), 
                    height=200,
                    key=f"logs_{len(logs)}"
                )
            
            # Procesar archivo
            try:
                df_final, df_resumen, df_estadisticas = procesar_archivo_excel(
                    archivo_subido,
                    progress_callback=actualizar_progreso,
                    log_callback=agregar_log
                )
                
                if df_final is not None:
                    st.success("🎉 ¡Archivo procesado exitosamente!")
                    
                    # Mostrar estadísticas
                    st.subheader("📈 Estadísticas del procesamiento")
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("📊 Registros", len(df_final))
                    with col2:
                        bancos_unicos = len(df_final['BANCO'].unique()) if 'BANCO' in df_final.columns else 0
                        st.metric("🏦 Bancos", bancos_unicos)
                    with col3:
                        suma_debe = df_final['DEBE'].sum() if 'DEBE' in df_final.columns else 0
                        st.metric("💰 Total DEBE", f"{suma_debe:,.2f}")
                    with col4:
                        suma_haber = df_final['HABER'].sum() if 'HABER' in df_final.columns else 0
                        st.metric("💰 Total HABER", f"{suma_haber:,.2f}")
                    
                    # Mostrar SALDO final
                    saldo_final = suma_debe - suma_haber
                    if saldo_final > 0:
                        st.success(f"💰 **SALDO FINAL: ${saldo_final:,.2f}** (A favor)")
                    elif saldo_final < 0:
                        st.error(f"💰 **SALDO FINAL: ${abs(saldo_final):,.2f}** (En contra)")
                    else:
                        st.info(f"💰 **SALDO FINAL: $0.00** (Balanceado)")
                    
                    # Tabs para mostrar resultados
                    tab1, tab2, tab3 = st.tabs(["📋 Datos Consolidados", "📊 Resumen del Proceso", "📈 Estadísticas"])
                    
                    with tab1:
                        st.subheader("📋 Datos Consolidados")
                        st.dataframe(df_final, use_container_width=True)
                        st.info(f"Total de registros: {len(df_final)}")
                        
                        # Mostrar muestra de registros con DEBE/HABER
                        debe_registros = df_final[df_final['DEBE'] > 0]
                        haber_registros = df_final[df_final['HABER'] > 0]
                        
                        if len(debe_registros) > 0:
                            st.subheader(f"💰 Registros DEBE ({len(debe_registros)} registros)")
                            st.dataframe(debe_registros[['Estado', 'Monto', 'DEBE', 'BANCO']].head(10), use_container_width=True)
                        
                        if len(haber_registros) > 0:
                            st.subheader(f"💰 Registros HABER ({len(haber_registros)} registros)")
                            st.dataframe(haber_registros[['Estado', 'Monto', 'HABER', 'BANCO']].head(10), use_container_width=True)
                    
                    with tab2:
                        st.subheader("📊 Resumen del Proceso")
                        if df_resumen is not None:
                            st.dataframe(df_resumen, use_container_width=True)
                        else:
                            st.warning("No hay datos de resumen disponibles")
                    
                    with tab3:
                        st.subheader("📈 Estadísticas")
                        if df_estadisticas is not None:
                            st.dataframe(df_estadisticas, use_container_width=True)
                        else:
                            st.warning("No hay estadísticas disponibles")
                    
                    # Generar archivo Excel para descarga
                    st.subheader("💾 Descargar resultados")
                    
                    # Crear archivo Excel en memoria
                    output = io.BytesIO()
                    
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, sheet_name='Datos_Consolidados', index=False)
                        if df_resumen is not None:
                            df_resumen.to_excel(writer, sheet_name='Resumen_Proceso', index=False)
                        if df_estadisticas is not None:
                            df_estadisticas.to_excel(writer, sheet_name='Estadisticas', index=False)
                    
                    # Preparar archivo para descarga
                    excel_data = output.getvalue()
                    fecha_actual = datetime.now().strftime('%Y%m%d_%H%M%S')
                    nombre_archivo = f'caratulas_vaciado_{fecha_actual}.xlsx'
                    
                    st.download_button(
                        label="📥 Descargar archivo Excel",
                        data=excel_data,
                        file_name=nombre_archivo,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        type="primary"
                    )
                    
                    st.success("✅ ¡Listo! Haz clic en el botón de arriba para descargar el archivo procesado.")
                    
                else:
                    st.error("❌ No se pudo procesar el archivo. Revisa los logs para más información.")
                    
            except Exception as e:
                st.error(f"❌ Error durante el procesamiento: {e}")
                agregar_log(f"Error crítico: {e}")
    
    # Footer
    st.markdown("---")
    st.markdown("🔧 **Desarrollado para automatizar el procesamiento de carátulas bancarias**")
    st.markdown("💡 **Tip**: Los montos se clasifican automáticamente usando los textos del campo Estado y la columna Tipo de extracto")

if __name__ == "__main__":
    main()
