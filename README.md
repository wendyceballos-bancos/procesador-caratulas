# 🏦 Vaciado Automático - Carátulas Bancarias

Aplicación web para procesar carátulas bancarias de múltiples bancos y monedas.

## 🚀 Características

✅ **Procesa fechas de cualquier año** (sin restricciones)  
✅ **Detecta automáticamente** las hojas útiles  
✅ **Mapea columnas** automáticamente  
✅ **Asigna monedas** según Flex banco  
✅ **Calcula DEBE/HABER/SALDO** automáticamente  
✅ **Interfaz web moderna** y fácil de usar  

## 📱 Uso

1. **Sube tu archivo** Excel con carátulas bancarias
2. **Clic en "Procesar Carátulas"**  
3. **Revisa los resultados** en las pestañas
4. **Descarga** el archivo procesado

## 🔧 Instalación Local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📊 Formato de Entrada

- Archivo Excel (.xlsx, .xlsm, .xls)
- Múltiples hojas con carátulas bancarias
- Headers en filas 9-11
- Datos desde fila 12+

## 📥 Resultado

Excel con 3 hojas:
- **Datos Consolidados**: Todos los movimientos procesados
- **Resumen por Hoja**: Estadísticas por hoja
- **Estadísticas Generales**: Resumen total

## 🏦 Bancos Soportados

- Santander (EUR, USD)
- Citi (USD, EUR)
- Y cualquier banco con formato estándar

## 🌐 Versión Web

Acceso directo: [Tu App Streamlit](https://tu-app.streamlit.app)
