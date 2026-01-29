
import streamlit as st
import pandas as pd
import numpy as np
import base64
import io
import os
from datetime import datetime, date
import warnings
import re

warnings.filterwarnings('ignore')

# Configuración de la página
st.set_page_config(
    page_title="🏦 Vaciado de Carátulas Bancarias - Despegar",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Logo de Despegar embebido en base64
LOGO_DESPEGAR_BASE64 = """iVBORw0KGgoAAAANSUhEUgAAAyEAAAMgCAYAAAA0sm1WAAAQAE..."""

# Archivo de mapeo Flex → Moneda embebido ACTUALIZADO
ARCHIVO_MAPEO_MONEDAS_BASE64 = """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"""

def mostrar_header_con_logo():
    """Muestra el header con logo de Despegar y título"""
    st.markdown(f"""
    <div style="display: flex; align-items: center; margin-bottom: 2rem; padding: 1rem; background: linear-gradient(90deg, #f8fafc 0%, #e2e8f0 100%); border-radius: 10px; border-left: 5px solid #1e3a8a;">
        <img src="data:image/png;base64,{LOGO_DESPEGAR_BASE64}" 
             style="height: 80px; margin-right: 25px;">
        <div>
            <h1 style="color: #1e3a8a; margin: 0; font-size: 2.5rem; font-weight: 700;">
                Vaciado de Carátulas Bancarias
            </h1>
            <p style="color: #475569; margin: 5px 0 0 0; font-size: 1.1rem; font-weight: 400;">
                Automatización de procesamiento financiero - Powered by Despegar
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)

@st.cache_data
def cargar_mapeo_monedas():
    """Carga el mapeo de monedas desde el archivo Excel embebido ACTUALIZADO"""
    try:
        # Decodificar el archivo base64
        archivo_bytes = base64.b64decode(ARCHIVO_MAPEO_MONEDAS_BASE64)
        
        # Leer el Excel desde memoria
        df_mapeo = pd.read_excel(io.BytesIO(archivo_bytes), engine='openpyxl')
        
        # Verificar estructura del archivo
        if 'Flex Efectivo' not in df_mapeo.columns or 'Moneda' not in df_mapeo.columns:
            raise Exception(f"Columnas incorrectas: {list(df_mapeo.columns)}")
        
        # Crear diccionario de mapeo
        mapeo_dict = dict(zip(df_mapeo['Flex Efectivo'].astype(str), df_mapeo['Moneda'].astype(str)))
        
        # Log para debug
        print(f"✅ Mapeo de monedas cargado exitosamente!")
        print(f"📊 Total entradas: {len(mapeo_dict)}")
        print(f"💰 Monedas disponibles: {sorted(df_mapeo['Moneda'].unique())}")
        print(f"🔍 Ejemplo de mapeo: {list(mapeo_dict.items())[0] if mapeo_dict else 'Vacío'}")
        
        # Verificar que no esté vacío
        if not mapeo_dict:
            raise Exception("Mapeo está vacío")
        
        return mapeo_dict, df_mapeo
        
    except Exception as e:
        error_msg = f"Error cargando mapeo actualizado: {e}"
        print(f"❌ {error_msg}")
        st.error(error_msg)
        
        # Mapeo de respaldo básico
        mapeo_respaldo = {
            'USD': 'USD', 'EUR': 'EUR', 'PEN': 'PEN', 'CLP': 'CLP', 
            'BRL': 'BRL', 'MXN': 'MXN', 'ARS': 'ARS', 'COP': 'COP', 'UYU': 'UYU'
        }
        return mapeo_respaldo, None

def limpiar_nombre_banco(nombre_hoja):
    """Extrae el nombre del banco desde el nombre de la hoja"""
    nombre_limpio = re.sub(r'\d+$', '', nombre_hoja)
    nombre_limpio = re.sub(r'\b(datos|data|info|información)\b', '', nombre_limpio, flags=re.IGNORECASE)
    nombre_limpio = nombre_limpio.strip().strip('_-. ')
    return nombre_limpio if nombre_limpio else nombre_hoja

def obtener_moneda_de_flex(flex_banco, mapeo_monedas):
    """
    Obtiene la moneda basada en el flex banco usando mapeo actualizado
    Implementa búsqueda exacta y por prefijos CON DEBUG
    """
    if pd.isna(flex_banco) or flex_banco == '':
        return ''
    
    flex_str = str(flex_banco).strip()
    
    # Debug limitado para no saturar logs
    debug_count = getattr(obtener_moneda_de_flex, 'debug_count', 0)
    if debug_count < 3:
        print(f"🔍 Buscando moneda para Flex: '{flex_str}' (Total mapeo: {len(mapeo_monedas)})")
        obtener_moneda_de_flex.debug_count = debug_count + 1
    
    # 1. Búsqueda exacta
    if flex_str in mapeo_monedas:
        moneda_encontrada = mapeo_monedas[flex_str]
        if debug_count < 3:
            print(f"✅ Coincidencia exacta: '{flex_str}' → '{moneda_encontrada}'")
        return moneda_encontrada
    
    # 2. Búsqueda por coincidencia parcial (contiene)
    for flex_clave, moneda in mapeo_monedas.items():
        if flex_str in flex_clave or flex_clave in flex_str:
            if debug_count < 3:
                print(f"✅ Coincidencia parcial: '{flex_str}' ↔ '{flex_clave}' → '{moneda}'")
            return moneda
    
    # 3. Búsqueda por prefijos jerárquicos
    partes_flex = flex_str.split('.')
    for i in range(len(partes_flex), 0, -1):
        prefijo = '.'.join(partes_flex[:i])
        for flex_clave, moneda in mapeo_monedas.items():
            if flex_clave.startswith(prefijo):
                if debug_count < 3:
                    print(f"✅ Coincidencia por prefijo: '{prefijo}' en '{flex_clave}' → '{moneda}'")
                return moneda
    
    # 4. Fallback por código de moneda
    flex_upper = flex_str.upper()
    codigos_moneda = ['USD', 'EUR', 'PEN', 'CLP', 'BRL', 'MXN', 'ARS', 'COP', 'UYU']
    for codigo in codigos_moneda:
        if codigo in flex_upper:
            if debug_count < 3:
                print(f"✅ Fallback por código: '{codigo}' en '{flex_upper}' → '{codigo}'")
            return codigo
    
    # No encontrado
    if debug_count < 3:
        print(f"⚠️ No se encontró moneda para: '{flex_str}'")
    
    return ''

def normalizar_fecha_sin_hora(fecha_value):
    """Normaliza fechas eliminando la hora"""
    try:
        if pd.isna(fecha_value) or fecha_value is None:
            return ''
        
        if hasattr(fecha_value, 'date') and callable(fecha_value.date):
            return fecha_value.date().strftime('%Y-%m-%d')
        
        if hasattr(fecha_value, 'strftime'):
            return fecha_value.strftime('%Y-%m-%d')
        
        fecha_str = str(fecha_value).strip()
        
        if ' ' in fecha_str and fecha_str != 'nan':
            try:
                fecha_parseada = pd.to_datetime(fecha_str, errors='coerce')
                if not pd.isna(fecha_parseada):
                    return fecha_parseada.strftime('%Y-%m-%d')
            except:
                pass
        
        if fecha_str and fecha_str != 'nan':
            try:
                fecha_parseada = pd.to_datetime(fecha_str, errors='coerce')
                if not pd.isna(fecha_parseada):
                    return fecha_parseada.strftime('%Y-%m-%d')
            except:
                pass
        
        return fecha_str if fecha_str != 'nan' else ''
    
    except Exception:
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
        
        # Cargar el mapeo de monedas actualizado
        mapeo_monedas, df_mapeo_original = cargar_mapeo_monedas()
        log(f"✅ Mapeo de monedas ACTUALIZADO cargado: {len(mapeo_monedas)} entradas")
        if df_mapeo_original is not None:
            monedas_disponibles = sorted(df_mapeo_original['Moneda'].unique())
            log(f"💰 Monedas disponibles: {monedas_disponibles}")
        
        # Resetear contador de debug
        if hasattr(obtener_moneda_de_flex, 'debug_count'):
            del obtener_moneda_de_flex.debug_count
        
        # Leer todas las hojas del archivo
        try:
            todas_las_hojas = pd.read_excel(archivo, sheet_name=None, header=None, engine='openpyxl')
            log(f"📄 Archivo cargado: {len(todas_las_hojas)} hojas encontradas")
        except Exception as e:
            log(f"❌ Error al leer el archivo: {e}")
            return None, None, None, None
        
        # Filtrar hojas que no son de datos
        hojas_excluidas = ['resumen', 'summary', 'índice', 'index', 'instrucciones', 'instructions', 'totales', 'total', 'config', 'configuración']
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
                    'Responsable', 'Flex contable', 'Flex banco',
                    'Moneda', 'BANCO'
                ]
                
                # Procesar todas las columnas excepto 'Moneda' y 'BANCO' que se agregan después
                for col in columnas_objetivo:
                    if col not in ['Moneda', 'BANCO']:
                        if col in mapeo_columnas:
                            valores = [df_hoja.iloc[idx, mapeo_columnas[col]] if idx < len(df_hoja) and mapeo_columnas[col] < len(df_hoja.columns) else '' for idx in filas_validas]
                            df_procesado[col] = valores
                        else:
                            df_procesado[col] = [''] * len(filas_validas)
                
                # Agregar columna de moneda usando mapeo ACTUALIZADO
                monedas_asignadas = 0
                if 'Flex banco' in mapeo_columnas:
                    log(f"💱 Aplicando mapeo de monedas a {len(filas_validas)} registros...")
                    monedas_aplicadas = []
                    for idx, flex_val in enumerate(df_procesado['Flex banco']):
                        moneda = obtener_moneda_de_flex(flex_val, mapeo_monedas)
                        monedas_aplicadas.append(moneda)
                        if moneda:
                            monedas_asignadas += 1
                        
                        # Debug solo para los primeros registros
                        if idx < 3:
                            log(f"🔍 Registro {idx+1}: Flex='{flex_val}' → Moneda='{moneda}'")
                    
                    df_procesado['Moneda'] = monedas_aplicadas
                    log(f"💰 Monedas asignadas: {monedas_asignadas} de {len(filas_validas)} registros")
                    
                    # Mostrar ejemplos de monedas encontradas
                    monedas_ejemplo = [m for m in monedas_aplicadas[:5] if m]
                    if monedas_ejemplo:
                        log(f"🔍 Ejemplos de monedas asignadas: {monedas_ejemplo}")
                    else:
                        log(f"⚠️ No se asignaron monedas en los primeros 5 registros")
                        # Mostrar ejemplos de Flex para debug
                        flex_ejemplo = [str(f)[:50] for f in df_procesado['Flex banco'].head(3)]
                        log(f"🔍 Ejemplos de Flex banco: {flex_ejemplo}")
                else:
                    log(f"⚠️ Columna 'Flex banco' no encontrada en mapeo_columnas")
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
                
                progreso = 10 + (60 * (i + 1) / total_hojas)
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
            return None, None, None, None
        
        # Consolidar todos los datos
        log("\n🔗 Consolidando datos...")
        df_final = pd.concat(datos_consolidados, ignore_index=True)
        actualizar_progreso(75)
        
        # Limpiar datos finales
        for col in df_final.columns:
            if df_final[col].dtype == 'object':
                df_final[col] = df_final[col].astype(str).str.strip()
                df_final[col] = df_final[col].replace('nan', '')
                if col not in ['Proveedor/Cliente', 'Moneda']:
                    df_final[col] = df_final[col].replace('', np.nan)
        
        # Normalizar fechas sin hora
        log("📅 Normalizando fechas sin hora...")
        if 'Fecha' in df_final.columns:
            fechas_procesadas = 0
            fechas_normalizadas = []
            
            for idx in df_final.index:
                fecha_original = df_final.at[idx, 'Fecha']
                fecha_normalizada = normalizar_fecha_sin_hora(fecha_original)
                fechas_normalizadas.append(fecha_normalizada)
                if fecha_normalizada:
                    fechas_procesadas += 1
            
            df_final['Fecha'] = fechas_normalizadas
            log(f"✅ Fechas normalizadas: {fechas_procesadas} registros procesados")
        
        actualizar_progreso(85)
        
        # Calcular DEBE/HABER/SALDO
        log("💰 Calculando DEBE, HABER y SALDO...")
        
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
            
            textos_debe = [
                "III. Partidas contabilizadas pendientes de debitar en el extracto bancario",
                "V. Partidas acreditadas en el extracto bancario, pendientes de contabilizar"
            ]
            
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
                
                if tipo_extracto == 'DEBIT':
                    df_final.at[idx, 'DEBE'] = abs(monto_val)
                    df_final.at[idx, 'HABER'] = 0.0
                    debe_count += 1
                elif tipo_extracto == 'CREDIT':
                    df_final.at[idx, 'HABER'] = abs(monto_val)
                    df_final.at[idx, 'DEBE'] = 0.0
                    haber_count += 1
                elif any(texto in estado_text for texto in textos_debe):
                    df_final.at[idx, 'DEBE'] = abs(monto_val)
                    df_final.at[idx, 'HABER'] = 0.0
                    debe_count += 1
                elif any(texto in estado_text for texto in textos_haber):
                    df_final.at[idx, 'HABER'] = abs(monto_val)
                    df_final.at[idx, 'DEBE'] = 0.0
                    haber_count += 1
                else:
                    if monto_val > 0:
                        df_final.at[idx, 'DEBE'] = monto_val
                        df_final.at[idx, 'HABER'] = 0.0
                        debe_count += 1
                    elif monto_val < 0:
                        df_final.at[idx, 'HABER'] = abs(monto_val)
                        df_final.at[idx, 'DEBE'] = 0.0
                        haber_count += 1
            
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
            'Responsable', 'Flex contable', 'Flex banco',
            'Moneda', 'BANCO', 'DEBE', 'HABER', 'SALDO'
        ]
        
        for col in df_final.columns:
            if col not in columnas_finales:
                columnas_finales.append(col)
        
        df_final = df_final[[col for col in columnas_finales if col in df_final.columns]]
        
        # Crear DataFrames de resumen
        df_resumen = pd.DataFrame(resumen_proceso)
        
        # Estadísticas de monedas
        if 'Moneda' in df_final.columns:
            monedas_stats = df_final['Moneda'].value_counts()
            df_monedas = pd.DataFrame({
                'Moneda': monedas_stats.index,
                'Registros': monedas_stats.values
            }).reset_index(drop=True)
        else:
            df_monedas = pd.DataFrame()
        
        df_estadisticas = pd.DataFrame({
            'Métrica': [
                'Total de registros procesados',
                'Número de bancos procesados', 
                'Total DEBE',
                'Total HABER',
                'SALDO FINAL',
                'Monedas diferentes encontradas',
                'Fecha de procesamiento',
                'Archivo procesado'
            ],
            'Valor': [
                len(df_final),
                len(df_final['BANCO'].unique()) if 'BANCO' in df_final.columns else 0,
                f"{df_final['DEBE'].sum():,.2f}" if 'DEBE' in df_final.columns else 0,
                f"{df_final['HABER'].sum():,.2f}" if 'HABER' in df_final.columns else 0,
                f"{df_final['SALDO'].sum():,.2f}" if 'SALDO' in df_final.columns else 0,
                df_final['Moneda'].nunique() if 'Moneda' in df_final.columns else 0,
                datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
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
        log(f"💱 Monedas procesadas: {df_final['Moneda'].nunique() if 'Moneda' in df_final.columns else 0}")
        
        # Debug final: Mostrar distribución de monedas
        if 'Moneda' in df_final.columns:
            monedas_count = df_final['Moneda'].value_counts()
            log(f"💱 Distribución final de monedas: {dict(monedas_count)}")
        
        return df_final, df_resumen, df_estadisticas, df_monedas
        
    except Exception as e:
        log(f"❌ Error crítico durante el procesamiento: {e}")
        import traceback
        log(f"📋 Traceback: {traceback.format_exc()}")
        return None, None, None, None

def main():
    """Función principal de la aplicación Streamlit"""
    
    # Header con logo de Despegar
    mostrar_header_con_logo()
    st.markdown("---")
    st.markdown("""
    ### 📋 Instrucciones de uso:
    1. **Sube tu archivo Excel** con las carátulas bancarias
    2. **Haz clic en 'Procesar archivo'** para iniciar el análisis
    3. **Descarga los resultados** en formato Excel
    
    ⚠️  **Importante**: El archivo debe tener encabezados en la fila 11 y datos a partir de la fila 13.
    """)
    
    # Sidebar con información del mapeo
    st.sidebar.header("⚙️ Configuración")
    st.sidebar.info("""
    **Formato esperado:**
    - Encabezados en fila 11
    - Datos desde fila 13
    - Columnas: Estado, Aging, Fecha, Monto, etc.
    """)
    
    # Mostrar información del mapeo de monedas
    with st.sidebar.expander("💱 Mapeo de Monedas ACTUALIZADO"):
        try:
            mapeo_monedas, df_mapeo = cargar_mapeo_monedas()
            if df_mapeo is not None:
                st.write(f"**Total de entradas:** {len(mapeo_monedas)}")
                st.write(f"**Monedas disponibles:** {len(df_mapeo['Moneda'].unique())}")
                monedas_lista = sorted(df_mapeo['Moneda'].unique())
                st.write("**Códigos:**")
                for moneda in monedas_lista:
                    st.write(f"• {moneda}")
            else:
                st.write("Mapeo básico de respaldo cargado")
        except Exception as e:
            st.write(f"Error: {e}")
    
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
                df_final, df_resumen, df_estadisticas, df_monedas = procesar_archivo_excel(
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
                    
                    # Mostrar SALDO final y monedas
                    col1, col2 = st.columns(2)
                    with col1:
                        saldo_final = suma_debe - suma_haber
                        if saldo_final > 0:
                            st.success(f"💰 **SALDO FINAL: ${saldo_final:,.2f}** (A favor)")
                        elif saldo_final < 0:
                            st.error(f"💰 **SALDO FINAL: ${abs(saldo_final):,.2f}** (En contra)")
                        else:
                            st.info(f"💰 **SALDO FINAL: $0.00** (Balanceado)")
                    
                    with col2:
                        if 'Moneda' in df_final.columns:
                            monedas_procesadas = df_final['Moneda'].nunique()
                            monedas_con_datos = len(df_final[df_final['Moneda'] != ''])
                            st.info(f"💱 **{monedas_procesadas} monedas diferentes** ({monedas_con_datos} registros con moneda)")
                    
                    # Tabs para mostrar resultados
                    tab1, tab2, tab3, tab4, tab5 = st.tabs([
                        "📋 Datos Consolidados", 
                        "💱 Análisis de Monedas", 
                        "📅 Vista de Fechas", 
                        "📊 Resumen del Proceso", 
                        "📈 Estadísticas"
                    ])
                    
                    with tab1:
                        st.subheader("📋 Datos Consolidados")
                        st.dataframe(df_final, use_container_width=True)
                        st.info(f"Total de registros: {len(df_final)}")
                        
                        # Mostrar muestra por DEBE/HABER
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            debe_registros = df_final[df_final['DEBE'] > 0]
                            if len(debe_registros) > 0:
                                st.subheader(f"💰 DEBE ({len(debe_registros)} registros)")
                                st.dataframe(
                                    debe_registros[['Estado', 'Fecha', 'Moneda', 'DEBE', 'BANCO']].head(5), 
                                    use_container_width=True
                                )
                        
                        with col2:
                            haber_registros = df_final[df_final['HABER'] > 0]
                            if len(haber_registros) > 0:
                                st.subheader(f"💰 HABER ({len(haber_registros)} registros)")
                                st.dataframe(
                                    haber_registros[['Estado', 'Fecha', 'Moneda', 'HABER', 'BANCO']].head(5), 
                                    use_container_width=True
                                )
                    
                    with tab2:
                        st.subheader("💱 Análisis de Monedas")
                        
                        if df_monedas is not None and len(df_monedas) > 0:
                            st.dataframe(df_monedas, use_container_width=True)
                            
                            # Gráfico de monedas
                            st.subheader("📊 Distribución por Moneda")
                            st.bar_chart(df_monedas.set_index('Moneda'))
                            
                            # Detalle por moneda
                            st.subheader("💰 Totales por Moneda")
                            if 'Moneda' in df_final.columns:
                                moneda_totales = df_final.groupby('Moneda').agg({
                                    'DEBE': 'sum',
                                    'HABER': 'sum',
                                    'SALDO': 'sum'
                                }).round(2)
                                st.dataframe(moneda_totales, use_container_width=True)
                        else:
                            st.warning("No se encontraron monedas asignadas")
                            
                            # Mostrar registros sin moneda
                            sin_moneda = df_final[df_final['Moneda'] == '']
                            if len(sin_moneda) > 0:
                                st.subheader("⚠️ Registros sin moneda asignada")
                                st.dataframe(sin_moneda[['Flex banco', 'BANCO', 'Monto']].head(10), use_container_width=True)
                    
                    with tab3:
                        st.subheader("📅 Vista de Fechas Normalizadas")
                        if 'Fecha' in df_final.columns:
                            fechas_df = df_final[['Fecha', 'BANCO', 'Moneda', 'Monto']].copy()
                            fechas_df = fechas_df[fechas_df['Fecha'].notna() & (fechas_df['Fecha'] != '')]
                            
                            if len(fechas_df) > 0:
                                st.dataframe(fechas_df.head(20), use_container_width=True)
                                
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("📅 Total con fecha", len(fechas_df))
                                with col2:
                                    st.metric("📅 Fechas únicas", fechas_df['Fecha'].nunique())
                                with col3:
                                    if len(fechas_df) > 0:
                                        st.metric("📅 Formato", fechas_df['Fecha'].iloc[0])
                            else:
                                st.warning("No se encontraron fechas válidas")
                        else:
                            st.warning("No se encontró columna de Fecha")
                    
                    with tab4:
                        st.subheader("📊 Resumen del Proceso")
                        if df_resumen is not None:
                            st.dataframe(df_resumen, use_container_width=True)
                        else:
                            st.warning("No hay datos de resumen disponibles")
                    
                    with tab5:
                        st.subheader("📈 Estadísticas Generales")
                        if df_estadisticas is not None:
                            st.dataframe(df_estadisticas, use_container_width=True)
                        else:
                            st.warning("No hay estadísticas disponibles")
                    
                    # Generar archivo Excel para descarga
                    st.subheader("💾 Descargar resultados")
                    
                    # Crear archivo Excel en memoria - SOLO DATOS CONSOLIDADOS
                    output = io.BytesIO()
                    
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, sheet_name='Datos_Consolidados', index=False)
                        # Solo se incluye la hoja principal según solicitud del usuario
                    
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
                import traceback
                agregar_log(f"Traceback completo: {traceback.format_exc()}")
    
    # Footer con branding de Despegar
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #6b7280; padding: 1rem;">
        <p><strong>🔧 Desarrollado para automatizar el procesamiento de carátulas bancarias</strong></p>
        <p>💱 <strong>Mapeo de monedas ACTUALIZADO</strong>: 136 entradas de Flex → Moneda</p>
        <p>📅 <strong>Fechas normalizadas</strong>: Sin hora (formato YYYY-MM-DD)</p>
        <p>✨ <strong>Powered by Despegar</strong> - Tecnología corporativa avanzada</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
