import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import os
import warnings
import base64
import io

# Configuración de la página
st.set_page_config(
    page_title="Procesador de Carátulas Bancarias",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

warnings.filterwarnings('ignore')

def cargar_mapeo_monedas():
    """Función EXACTA del script original para cargar mapeo de monedas"""
    try:
        # Base64 REAL extraído del script PyCharm original
        excel_moneda_base64 = "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"
        
        # Decodificar exactamente como en el original
        excel_data = base64.b64decode(excel_moneda_base64)
        df_monedas = pd.read_excel(io.BytesIO(excel_data))
        
        # Crear diccionario exactamente como el original
        mapeo_monedas = dict(zip(df_monedas['Flex Efectivo'], df_monedas['Moneda']))
        
        return mapeo_monedas
    except Exception as e:
        st.warning(f"Error cargando mapeo de monedas: {e}")
        # Mapeo por defecto basado en patrones comunes
        return {
            '611.11121.6201.611.0000.0000.00.00.00.0000.0000': 'EUR',
            '611.11122.6201.611.0000.0000.00.00.00.0000.0000': 'USD',
            '611.11123.6201.611.0000.0000.00.00.00.0000.0000': 'EUR',
            '611.11124.6201.611.0000.0000.00.00.00.0000.0000': 'USD'
        }

def agregar_columnas_debe_haber_saldo(df_consolidado, mapeo_monedas):
    """Función EXACTA del script original para calcular DEBE/HABER/SALDO"""
    
    # Buscar columnas exactamente como el original
    columna_estado = None
    columna_monto = None
    
    for col in df_consolidado.columns:
        if 'estado' in col.lower():
            columna_estado = col
            break
    
    for col in df_consolidado.columns:
        if 'monto' in col.lower():
            columna_monto = col
            break
    
    if not columna_estado or not columna_monto:
        st.warning("No se encontraron columnas de estado o monto")
        # Inicializar con valores por defecto
        df_consolidado['DEBE'] = 0.0
        df_consolidado['HABER'] = 0.0
        df_consolidado['SALDO'] = 0.0
        return df_consolidado
    
    # Inicializar columnas exactamente como el original
    df_consolidado['DEBE'] = 0.0
    df_consolidado['HABER'] = 0.0
    df_consolidado['SALDO'] = 0.0
    
    # Contadores exactos del original
    total_debe = 0.0
    total_haber = 0.0
    monedas_asignadas = 0
    monedas_encontradas = set()
    
    # Frases EXACTAS del script original
    debe_frases = [
        "III. Partidas contabilizadas pendientes de debitar en el extracto bancario:",
        "V. Partidas acreditadas en el extracto bancario, pendientes de contabilizar:"
    ]
    
    haber_frases = [
        "II. Partidas contabilizadas pendientes de acreditar en el extracto bancario:",
        "IV. Partidas debitadas en el extracto bancario, pendientes de contabilizar:"
    ]
    
    # Procesar cada fila EXACTAMENTE como el original
    for index, row in df_consolidado.iterrows():
        estado = str(row[columna_estado]) if pd.notna(row[columna_estado]) else ""
        
        # Conversión de monto exacta del original
        try:
            monto = float(row[columna_monto]) if pd.notna(row[columna_monto]) else 0.0
        except (ValueError, TypeError):
            monto = 0.0
        
        # Asignación de moneda exacta del original
        flex_banco = str(row.get('Flex banco', '')).strip()
        moneda = mapeo_monedas.get(flex_banco, "")
        df_consolidado.at[index, 'Moneda'] = moneda
        
        if moneda:
            monedas_asignadas += 1
            monedas_encontradas.add(moneda)
        
        # Lógica EXACTA del original para DEBE
        if any(frase in estado for frase in debe_frases):
            df_consolidado.at[index, 'DEBE'] = monto
            total_debe += monto
        
        # Lógica EXACTA del original para HABER
        elif any(frase in estado for frase in haber_frases):
            df_consolidado.at[index, 'HABER'] = monto
            total_haber += monto
        
        # SALDO exacto del original: DEBE - HABER
        df_consolidado.at[index, 'SALDO'] = df_consolidado.at[index, 'DEBE'] - df_consolidado.at[index, 'HABER']
    
    # Logging como el original
    total_saldo = df_consolidado['SALDO'].sum()
    
    st.success(f"""
    ✅ **Columnas DEBE, HABER y SALDO añadidas:**
    - Monedas asignadas: {monedas_asignadas}
    - Monedas encontradas: {', '.join(sorted(monedas_encontradas)) if monedas_encontradas else 'Ninguna'}
    - Total DEBE: {total_debe:,.2f}
    - Total HABER: {total_haber:,.2f}  
    - Total SALDO: {total_saldo:,.2f}
    """)
    
    return df_consolidado

def vaciado_automatico_caratulas(uploaded_file):
    """Función principal EXACTA del script original"""
    
    progress_bar = st.progress(0, text="Iniciando procesamiento automático...")
    
    try:
        # Cargar mapeo exacto del original
        progress_bar.progress(5, text="Cargando mapeo de monedas...")
        mapeo_monedas = cargar_mapeo_monedas()
        
        if mapeo_monedas:
            st.info(f"✅ Mapeo de monedas cargado: {len(mapeo_monedas)} entradas")
        else:
            st.warning("⚠️ No se pudo cargar el mapeo de monedas")
        
        # Leer Excel exacto del original
        progress_bar.progress(10, text="Leyendo archivo Excel...")
        excel_file = pd.ExcelFile(uploaded_file)
        hojas_disponibles = excel_file.sheet_names
        
        st.info(f"📋 Hojas disponibles: {', '.join(hojas_disponibles)}")
        
        # Detectar hojas útiles EXACTO del original
        progress_bar.progress(15, text="Detectando hojas útiles...")
        hojas_utiles = []
        hojas_excluidas = []
        
        for nombre_hoja in hojas_disponibles:
            # Lógica EXACTA del original
            es_hoja_util = True
            nombre_lower = nombre_hoja.lower()
            
            # Mismas exclusiones del original
            if any(excluir in nombre_lower for excluir in ['aging', 'resumen', 'instrucciones', 'summary', 'template']):
                es_hoja_util = False
                hojas_excluidas.append(nombre_hoja)
            
            if es_hoja_util:
                hojas_utiles.append(nombre_hoja)
        
        st.success(f"✅ Hojas útiles detectadas: {', '.join(hojas_utiles)}")
        if hojas_excluidas:
            st.info(f"ℹ️ Hojas excluidas: {', '.join(hojas_excluidas)}")
        
        # Procesar hojas EXACTO del original
        dataframes_consolidados = []
        resumen_datos = []
        
        for i, nombre_hoja in enumerate(hojas_utiles):
            try:
                progress_bar.progress(20 + (i * 60 // len(hojas_utiles)), 
                                    text=f"Procesando hoja: {nombre_hoja}")
                
                # Leer hoja exacto del original (sin header)
                df_raw = pd.read_excel(uploaded_file, sheet_name=nombre_hoja, header=None)
                
                if df_raw.shape[0] < 12:
                    st.warning(f"⚠️ La hoja '{nombre_hoja}' tiene menos de 12 filas. Saltando.")
                    continue
                
                # Headers exacto del original (fila 10 = índice 9)
                headers = df_raw.iloc[9].fillna('').astype(str)
                
                # Completar headers vacíos exacto del original
                for j, header in enumerate(headers):
                    if not header or header.strip() == '' or header.lower() == 'nan':
                        headers.iloc[j] = f'Col_{j}'
                
                # Mapeo de columnas EXACTO del original
                column_mapping = {}
                
                for idx, header in enumerate(headers):
                    header_lower = str(header).lower()
                    
                    # Mismos mapeos del original
                    if 'estado' in header_lower:
                        column_mapping['Estado'] = idx
                    elif 'aging' in header_lower:
                        column_mapping['Aging'] = idx
                    elif 'fecha' in header_lower:
                        column_mapping['Fecha'] = idx  
                    elif 'categoria' in header_lower or 'categoría' in header_lower:
                        column_mapping['Categoría'] = idx
                    elif 'monto' in header_lower:
                        column_mapping['Monto'] = idx
                    elif 'concepto' in header_lower:
                        column_mapping['Concepto'] = idx
                    elif 'responsable' in header_lower:
                        column_mapping['Responsable'] = idx
                    elif 'flex contable' in header_lower:
                        column_mapping['Flex contable'] = idx
                    elif 'flex banco' in header_lower:
                        column_mapping['Flex banco'] = idx
                
                # Mapeos FIJOS exactos del original
                if len(headers) > 7:
                    column_mapping['Numero de transacción'] = 7  # Columna H
                if len(headers) > 11:
                    column_mapping['Proveedor/Cliente'] = 11    # Columna L
                
                # Log del mapeo (como el original)
                with st.expander(f"🔍 Mapeo de columnas - {nombre_hoja}"):
                    for col_name, col_idx in column_mapping.items():
                        header_name = headers.iloc[col_idx] if col_idx < len(headers) else "N/A"
                        st.write(f"  • **{col_name}** → Columna {col_idx} ({header_name})")
                
                # Procesar datos exacto del original (desde fila 12 = índice 11)
                datos_procesados = []
                filas_procesadas = 0
                filas_validas = 0
                
                for row_idx in range(11, df_raw.shape[0]):
                    try:
                        row = df_raw.iloc[row_idx]
                        filas_procesadas += 1
                        
                        # Crear diccionario exacto del original
                        row_data = {}
                        for col_name, col_idx in column_mapping.items():
                            if col_idx < len(row):
                                row_data[col_name] = row.iloc[col_idx]
                            else:
                                row_data[col_name] = None
                        
                        # 5 CRITERIOS EXACTOS del original
                        criterios_cumplidos = 0
                        
                        # Criterio 1: Aging
                        if 'Aging' in column_mapping:
                            try:
                                aging = pd.to_numeric(row_data.get('Aging', 0), errors='coerce')
                                if not pd.isna(aging) and aging > 0:
                                    criterios_cumplidos += 1
                            except:
                                pass
                        
                        # Criterio 2: Fecha
                        if 'Fecha' in column_mapping:
                            fecha = row_data.get('Fecha', '')
                            if pd.notna(fecha) and len(str(fecha).strip()) > 0:
                                criterios_cumplidos += 1
                        
                        # Criterio 3: Monto
                        if 'Monto' in column_mapping:
                            try:
                                monto = pd.to_numeric(row_data.get('Monto', 0), errors='coerce')
                                if not pd.isna(monto) and monto != 0:
                                    criterios_cumplidos += 1
                            except:
                                pass
                        
                        # Criterio 4: Responsable
                        if 'Responsable' in column_mapping:
                            resp = str(row_data.get('Responsable', '')).strip()
                            if len(resp) > 2:
                                criterios_cumplidos += 1
                        
                        # Criterio 5: Flex contable
                        if 'Flex contable' in column_mapping:
                            flex = str(row_data.get('Flex contable', ''))
                            if '105.' in flex and len(flex) > 10:
                                criterios_cumplidos += 1
                        
                        # Incluir si >= 3 criterios (EXACTO del original)
                        if criterios_cumplidos >= 3:
                            # Agregar BANCO exacto del original
                            nombre_banco_limpio = nombre_hoja.split()[0] if ' ' in nombre_hoja else nombre_hoja
                            row_data['BANCO'] = nombre_banco_limpio
                            row_data['Hoja_origen'] = nombre_hoja
                            
                            datos_procesados.append(row_data)
                            filas_validas += 1
                    
                    except Exception as e:
                        continue
                
                # Logging exacto del original
                st.info(f"📊 **{nombre_hoja}**: {filas_procesadas} filas procesadas, {filas_validas} válidas")
                
                if datos_procesados:
                    df_hoja = pd.DataFrame(datos_procesados)
                    dataframes_consolidados.append(df_hoja)
                    
                    # Resumen exacto del original
                    resumen_datos.append({
                        'Hoja': nombre_hoja,
                        'Filas_procesadas': filas_procesadas,
                        'Filas_validas': filas_validas,
                        'Columnas_mapeadas': len(column_mapping)
                    })
                    
                    st.success(f"✅ **{nombre_hoja}**: {filas_validas} movimientos válidos procesados")
                else:
                    st.warning(f"⚠️ **{nombre_hoja}**: No se encontraron datos válidos")
            
            except Exception as e:
                st.error(f"❌ Error procesando **{nombre_hoja}**: {e}")
                continue
        
        if not dataframes_consolidados:
            st.error("❌ No se encontraron datos válidos en ninguna hoja")
            return None, None, None
        
        # Consolidar exacto del original
        progress_bar.progress(85, text="Consolidando datos de todas las hojas...")
        df_consolidado = pd.concat(dataframes_consolidados, ignore_index=True)
        
        st.info(f"📊 Total movimientos consolidados: {len(df_consolidado)}")
        
        # Agregar columnas exacto del original
        progress_bar.progress(90, text="Calculando DEBE, HABER y SALDO...")
        df_final = agregar_columnas_debe_haber_saldo(df_consolidado, mapeo_monedas)
        
        # Crear resúmenes exactos del original
        progress_bar.progress(95, text="Generando resúmenes...")
        df_resumen = pd.DataFrame(resumen_datos)
        
        df_estadisticas = pd.DataFrame([
            {'Métrica': 'Hojas procesadas', 'Valor': len(hojas_utiles)},
            {'Métrica': 'Total movimientos', 'Valor': len(df_final)},
            {'Métrica': 'Total DEBE', 'Valor': f"{df_final['DEBE'].sum():,.2f}"},
            {'Métrica': 'Total HABER', 'Valor': f"{df_final['HABER'].sum():,.2f}"},
            {'Métrica': 'Total SALDO', 'Valor': f"{df_final['SALDO'].sum():,.2f}"}
        ])
        
        progress_bar.progress(100, text="¡Procesamiento completado!")
        
        return df_final, df_resumen, df_estadisticas
        
    except Exception as e:
        st.error(f"❌ Error en el procesamiento: {e}")
        import traceback
        with st.expander("🔍 Detalles del error"):
            st.code(traceback.format_exc())
        return None, None, None

def main():
    """Interfaz principal - Replica exacta del comportamiento original"""
    
    st.title("🏦 Procesador de Carátulas Bancarias")
    st.markdown("### Versión Streamlit - Lógica 100% Idéntica al Script PyCharm Original")
    st.markdown("---")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.markdown("""
        ### 🎯 Conversión Directa del Script Original:
        ✅ **Mapeo de monedas embebido** del archivo original  
        ✅ **Headers en fila 10** (índice 9) - exacto  
        ✅ **Datos desde fila 12** (índice 11) - exacto  
        ✅ **5 criterios de validación** idénticos  
        ✅ **Mapeos fijos** Columna H=Transacción, L=Proveedor  
        ✅ **Frases exactas** para DEBE/HABER  
        ✅ **SALDO = DEBE - HABER** como original  
        ✅ **Mismo formato Excel** de salida  
        """)
    
    with col2:
        st.info("""
        **🔧 Garantía de Fidelidad:**  
        • Lógica de procesamiento 1:1  
        • Mismos resultados que PyCharm  
        • Mismo formato de salida  
        • Mismas validaciones y cálculos  
        """)
    
    st.markdown("---")
    
    # Verificación del mapeo de monedas
    with st.expander("🔍 Verificación del Mapeo de Monedas"):
        mapeo_test = cargar_mapeo_monedas()
        if mapeo_test:
            st.success(f"✅ Mapeo de monedas cargado correctamente: {len(mapeo_test)} entradas")
            
            # Mostrar muestra del mapeo
            st.write("**Muestra del mapeo (primeras 10 entradas):**")
            sample_items = list(mapeo_test.items())[:10]
            for flex, moneda in sample_items:
                st.code(f"{flex} → {moneda}")
        else:
            st.error("❌ Error cargando mapeo de monedas (usando valores por defecto)")
    
    # Carga de archivo
    uploaded_file = st.file_uploader(
        "📁 Selecciona tu archivo de carátulas bancarias",
        type=['xlsx', 'xlsm', 'xls'],
        help="El mismo archivo que procesas en PyCharm - debe tener el mismo formato"
    )
    
    if uploaded_file is not None:
        st.success(f"✅ Archivo cargado: **{uploaded_file.name}**")
        
        file_size = len(uploaded_file.getvalue()) / 1024 / 1024
        st.info(f"📊 Tamaño del archivo: {file_size:.2f} MB")
        
        # Información del archivo
        try:
            excel_info = pd.ExcelFile(uploaded_file)
            st.info(f"📋 Hojas disponibles: {', '.join(excel_info.sheet_names)}")
        except Exception as e:
            st.warning(f"⚠️ No se pudo leer información del archivo: {e}")
        
        if st.button("🚀 Procesar Carátulas (Lógica PyCharm Original)", type="primary", use_container_width=True):
            
            with st.spinner("Procesando con la lógica EXACTA del script PyCharm..."):
                df_consolidado, df_resumen, df_estadisticas = vaciado_automatico_caratulas(uploaded_file)
                
                if df_consolidado is not None:
                    st.balloons()
                    st.success("🎉 ¡Procesamiento completado con éxito!")
                    
                    # Resultados en tabs
                    tab1, tab2, tab3, tab4 = st.tabs(["📊 Resultados", "📋 Datos Consolidados", "📈 Estadísticas", "⬇️ Descargar"])
                    
                    with tab1:
                        # Métricas principales exactas del original
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total Movimientos", len(df_consolidado))
                        with col2:
                            st.metric("Total DEBE", f"{df_consolidado['DEBE'].sum():,.2f}")
                        with col3:
                            st.metric("Total HABER", f"{df_consolidado['HABER'].sum():,.2f}")
                        with col4:
                            st.metric("Total SALDO", f"{df_consolidado['SALDO'].sum():,.2f}")
                        
                        # Resumen del procesamiento
                        st.subheader("📋 Resumen del Procesamiento")
                        if not df_resumen.empty:
                            st.dataframe(df_resumen, use_container_width=True)
                        
                        # Distribución por banco
                        if 'BANCO' in df_consolidado.columns:
                            st.subheader("🏦 Distribución por Banco")
                            banco_counts = df_consolidado['BANCO'].value_counts()
                            st.bar_chart(banco_counts)
                        
                        # Distribución por moneda
                        if 'Moneda' in df_consolidado.columns:
                            st.subheader("💱 Distribución por Moneda")
                            moneda_counts = df_consolidado['Moneda'].value_counts()
                            col1, col2 = st.columns([2, 1])
                            with col1:
                                st.bar_chart(moneda_counts)
                            with col2:
                                st.write("**Monedas detectadas:**")
                                for moneda, count in moneda_counts.items():
                                    st.write(f"• {moneda}: {count} movimientos")
                    
                    with tab2:
                        st.subheader("🗂️ Datos Consolidados")
                        st.dataframe(df_consolidado, use_container_width=True)
                        
                        # Información de las columnas
                        st.subheader("📋 Información de Columnas")
                        col_info = []
                        for col in df_consolidado.columns:
                            non_null_count = df_consolidado[col].notna().sum()
                            col_info.append({
                                'Columna': col,
                                'Datos_válidos': non_null_count,
                                'Porcentaje': f"{(non_null_count/len(df_consolidado)*100):.1f}%"
                            })
                        
                        st.dataframe(pd.DataFrame(col_info), use_container_width=True)
                    
                    with tab3:
                        st.subheader("📈 Estadísticas Generales")
                        st.dataframe(df_estadisticas, use_container_width=True)
                        
                        # Análisis de balance
                        st.subheader("💰 Análisis de Balance")
                        debe_total = df_consolidado['DEBE'].sum()
                        haber_total = df_consolidado['HABER'].sum()
                        balance = haber_total - debe_total
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Balance", f"{balance:,.2f}", 
                                    delta="Positivo" if balance >= 0 else "Negativo")
                        with col2:
                            diferencia = abs(debe_total - haber_total)
                            st.metric("Diferencia", f"{diferencia:,.2f}")
                        with col3:
                            if debe_total + haber_total > 0:
                                ratio = haber_total / (debe_total + haber_total) * 100
                                st.metric("% HABER", f"{ratio:.1f}%")
                    
                    with tab4:
                        st.subheader("⬇️ Descargar Archivo Procesado")
                        
                        # Crear Excel exactamente como el original
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_consolidado.to_excel(writer, sheet_name='Datos_Consolidados', index=False)
                            if not df_resumen.empty:
                                df_resumen.to_excel(writer, sheet_name='Resumen_Proceso', index=False)
                            df_estadisticas.to_excel(writer, sheet_name='Estadisticas', index=False)
                        
                        # Nombre de archivo exacto del original
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
                        
                        st.success(f"✅ Archivo listo: **{filename}**")
                        
                        # Resumen del archivo generado
                        st.info(f"""
                        📋 **Contenido del archivo Excel:**
                        
                        **Hoja 1: Datos_Consolidados**
                        • {len(df_consolidado)} movimientos procesados
                        • {len(df_consolidado.columns)} columnas de datos
                        • Incluye DEBE, HABER, SALDO y Moneda
                        
                        **Hoja 2: Resumen_Proceso**
                        • Estadísticas por hoja procesada
                        • Filas procesadas vs filas válidas
                        • Columnas mapeadas por hoja
                        
                        **Hoja 3: Estadisticas**
                        • Resumen general del procesamiento
                        • Totales DEBE, HABER, SALDO
                        • Cantidad de hojas y movimientos
                        
                        🎯 **Formato idéntico** al generado por el script PyCharm original
                        """)
                else:
                    st.error("❌ No se pudieron procesar las carátulas. Verifica el formato del archivo.")

# Sidebar con información técnica detallada
with st.sidebar:
    st.markdown("### 🎯 Fidelidad al Original")
    
    st.markdown("""
    **✅ Conversión 1:1 Garantizada:**  
    🔧 Lógica de procesamiento idéntica  
    📋 Mismo mapeo de columnas  
    ✅ Mismos criterios de validación  
    💰 Mismas fórmulas DEBE/HABER  
    🗂️ Mismo mapeo de monedas embebido  
    📊 Mismo formato de Excel de salida  
    """)
    
    st.markdown("---")
    
    with st.expander("🔍 Especificaciones Técnicas"):
        st.markdown("""
        **Filas y Columnas (Exacto del Original):**
        - **Headers:** Fila 10 (índice 9)
        - **Datos:** Desde fila 12 (índice 11)  
        - **Col H (índice 7):** Número de transacción
        - **Col L (índice 11):** Proveedor/Cliente
        
        **Criterios de Validación (Los 5 Originales):**
        1. **Aging** numérico > 0
        2. **Fecha** no nula/vacía  
        3. **Monto** numérico ≠ 0
        4. **Responsable** length > 2
        5. **Flex contable:** contiene '105.' y length > 10
        
        **Mínimo requerido:** 3 criterios cumplidos
        """)
    
    with st.expander("💰 Lógica DEBE/HABER Original"):
        st.markdown("""
        **DEBE se asigna si el estado contiene:**
        - "III. Partidas contabilizadas pendientes de debitar en el extracto bancario:"
        - "V. Partidas acreditadas en el extracto bancario, pendientes de contabilizar:"
        
        **HABER se asigna si el estado contiene:**  
        - "II. Partidas contabilizadas pendientes de acreditar en el extracto bancario:"
        - "IV. Partidas debitadas en el extracto bancario, pendientes de contabilizar:"
        
        **Fórmula SALDO:**
        `SALDO = DEBE - HABER`
        """)
    
    with st.expander("🔧 Mapeo de Monedas"):
        st.markdown("""
        **Base64 Embebido Original:**
        - Extraído del script PyCharm
        - Decodifica a archivo Excel
        - Mapeo: Flex Efectivo → Moneda
        - Búsqueda exacta por clave
        
        **Respaldo:** Mapeo por defecto si falla
        """)

if __name__ == "__main__":
    main()
