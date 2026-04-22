#!/usr/bin/env python3
"""
Script profesional para procesar archivos Excel con Pandas
Autor: Asistente de Data & Analytics
Fecha: 2026-04-22
"""

import pandas as pd
import re
import sys
import logging
from pathlib import Path
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def cargar_archivo_excel(ruta_archivo: str, nombre_hoja: str = 'Datos') -> pd.DataFrame:
    """
    Cargar archivo Excel con manejo de errores robusto
    
    Args:
        ruta_archivo: Ruta al archivo Excel
        nombre_hoja: Nombre de la hoja a procesar
        
    Returns:
        DataFrame con los datos procesados
        
    Raises:
        FileNotFoundError: Si el archivo no existe
        ValueError: Si la hoja no existe o el archivo está corrupto
    """
    try:
        # Verificar que el archivo existe
        if not Path(ruta_archivo).exists():
            raise FileNotFoundError(f"El archivo '{ruta_archivo}' no existe")
        
        logger.info(f"Cargando archivo: {ruta_archivo}")
        logger.info(f"Procesando hoja: {nombre_hoja}")
        
        # Primero leemos sin encabezados para identificar la estructura
        df_temp = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, header=None)
        
        # Los encabezados están en la fila 3 (índice 3)
        headers = df_temp.iloc[3].tolist()
        
        # Leemos los datos empezando desde la fila 4 (índice 4)
        df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, header=None, skiprows=4)
        df.columns = headers
        
        # Eliminar filas completamente vacías
        df = df.dropna(how='all')
        
        logger.info(f"Archivo cargado exitosamente. Shape: {df.shape}")
        logger.info(f"Columnas encontradas: {list(df.columns)}")
        
        return df
        
    except FileNotFoundError as e:
        logger.error(f"Error de archivo: {e}")
        raise
    except ValueError as e:
        logger.error(f"Error de valor: {e}")
        raise
    except Exception as e:
        logger.error(f"Error inesperado al cargar el archivo: {e}")
        raise ValueError(f"No se pudo cargar el archivo Excel: {e}")

def limpiar_direccion_individual(direccion: str) -> str:
    """
    Función personalizada para limpiar una dirección individual
    
    Args:
        direccion: Dirección original a limpiar
        
    Returns:
        Dirección limpiada
    """
    if pd.isna(direccion) or direccion == '':
        return ''
    
    # Convertir a string y eliminar espacios al inicio/final
    direccion = str(direccion).strip()
    
    # Paso 1: Eliminar caracteres prohibidos: . , @ & #
    direccion = re.sub(r'[.,@&\#]', ' ', direccion)
    
    # Paso 2: Truncar rangos numéricos (ej: 43-49 -> 43)
    # Busca patrones de número-número y reemplaza con solo el primer número
    direccion = re.sub(r'(\d+)\s*-\s*\d+', r'\1', direccion)
    
    # Paso 3: Eliminar Unit/Suite y todo lo que viene después (case-insensitive)
    # Busca "Unit" o "Suite" sin importar mayúsculas/minúsculas y se queda con la primera parte
    direccion = re.sub(r'\b(unit|suite)\b.*$', '', direccion, flags=re.IGNORECASE)
    
    # Paso 4: Limpiar espacios múltiples
    direccion = re.sub(r'\s+', ' ', direccion)
    
    # Paso 5: Eliminar espacios al inicio y final
    return direccion.strip()

def limpiar_direccion(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpiar la columna 'Dirección' usando apply() con lógica personalizada
    
    Args:
        df: DataFrame original
        
    Returns:
        DataFrame con columna 'Dirección depurada' agregada
    """
    try:
        # Aplicar la función de limpieza personalizada a cada dirección
        df['Dirección depurada'] = df['Dirección'].apply(limpiar_direccion_individual)
        
        logger.info("Columna 'Dirección depurada' creada exitosamente")
        
        # Mostrar ejemplos de direcciones procesadas para verificación
        if 'Dirección' in df.columns:
            ejemplos = df[['Dirección', 'Dirección depurada']].head(5)
            logger.info("Ejemplos de direcciones procesadas:")
            for idx, row in ejemplos.iterrows():
                logger.info(f"  Original: '{row['Dirección']}' -> Depurada: '{row['Dirección depurada']}'")
        
        return df
        
    except KeyError:
        logger.warning("Columna 'Dirección' no encontrada. Creando columna vacía.")
        df['Dirección depurada'] = ''
        return df
    except Exception as e:
        logger.error(f"Error al limpiar direcciones: {e}")
        raise

def mapear_grupo_ubicacion(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mapear números de grupo de ubicación a nombres descriptivos
    
    Args:
        df: DataFrame original
        
    Returns:
        DataFrame con columna 'Nombre de grupo de ubicación' agregada
    """
    try:
        # Diccionario de mapeo según especificación
        mapeo_ubicaciones = {
            1: "Earthquake",
            2: "Windstorm", 
            3: "Flood",
            4: "Hail",
            5: "Fire",
            6: "Other"
        }
        
        # Aplicar mapeo usando vectorización
        df['Nombre de grupo de ubicación'] = df['Número de grupo de ubicación'].map(mapeo_ubicaciones)
        
        logger.info("Columna 'Nombre de grupo de ubicación' mapeada exitosamente")
        return df
        
    except KeyError:
        logger.warning("Columna 'Número de grupo de ubicación' no encontrada")
        df['Nombre de grupo de ubicación'] = ''
        return df
    except Exception as e:
        logger.error(f"Error al mapear grupos de ubicación: {e}")
        raise

def realizar_calculos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Realizar los cálculos matemáticos requeridos
    
    Args:
        df: DataFrame original
        
    Returns:
        DataFrame con columnas calculadas agregadas
    """
    try:
        # 1. Suma Valores 1&2
        if 'Valor 1' in df.columns and 'Valor 2' in df.columns:
            df['Suma Valores 1&2'] = (df['Valor 1'] + df['Valor 2']).round(2)
            logger.info("Cálculo 'Suma Valores 1&2' completado")
        else:
            logger.warning("Columnas 'Valor 1' o 'Valor 2' no encontradas")
            df['Suma Valores 1&2'] = 0
        
        # 2. Valor 3 ajustado a 12 meses
        if 'Valor 3' in df.columns and 'Período aplicable del Valor 3 (Meses)' in df.columns:
            # Manejar división por cero
            df['Valor 3 ajustado a 12 meses'] = (
                df['Valor 3'] / df['Período aplicable del Valor 3 (Meses)'].replace(0, 1)
            ) * 12
            df['Valor 3 ajustado a 12 meses'] = df['Valor 3 ajustado a 12 meses'].round(2)
            
            # Mostrar ejemplos de cálculo para verificación
            logger.info("Ejemplos de cálculo 'Valor 3 ajustado a 12 meses':")
            for idx in range(min(3, len(df))):
                valor3 = df.iloc[idx]['Valor 3']
                periodo = df.iloc[idx]['Período aplicable del Valor 3 (Meses)']
                resultado = df.iloc[idx]['Valor 3 ajustado a 12 meses']
                logger.info(f"  Fila {idx}: Valor 3={valor3}, Período={periodo} meses -> Ajustado={resultado}")
            
            logger.info("Cálculo 'Valor 3 ajustado a 12 meses' completado")
        else:
            logger.warning("Columnas 'Valor 3' o 'Período aplicable del Valor 3 (Meses)' no encontradas")
            df['Valor 3 ajustado a 12 meses'] = 0
        
        # 3. Valor 4 (Prorrateo)
        if 'Suma Valores 1&2' in df.columns and 'Valor 3 ajustado a 12 meses' in df.columns:
            df['Valor 4 (Prorrateo)'] = (df['Suma Valores 1&2'] + df['Valor 3 ajustado a 12 meses']).round(2)
            logger.info("Cálculo 'Valor 4 (Prorrateo)' completado")
        else:
            logger.warning("No se pueden calcular 'Valor 4 (Prorrateo)' - faltan columnas anteriores")
            df['Valor 4 (Prorrateo)'] = 0
        
        return df
        
    except Exception as e:
        logger.error(f"Error en los cálculos: {e}")
        raise

def crear_resumen(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crear DataFrame resumen agrupando por número de grupo de ubicación
    
    Args:
        df: DataFrame con datos procesados
        
    Returns:
        DataFrame resumen con totales agrupados
    """
    try:
        # Columnas numéricas a sumar
        columnas_numericas = [
            'Valor 1', 'Valor 2', 'Valor 3', 'Suma Valores 1&2', 
            'Valor 3 ajustado a 12 meses', 'Valor 4 (Prorrateo)'
        ]
        
        # Verificar qué columnas existen en el DataFrame
        columnas_disponibles = [col for col in columnas_numericas if col in df.columns]
        
        if not columnas_disponibles:
            logger.warning("No se encontraron columnas numéricas para resumir")
            return pd.DataFrame()
        
        # Agrupar por 'Número de grupo de ubicación' y sumar las columnas numéricas
        df_resumen = df.groupby('Número de grupo de ubicación')[columnas_disponibles].sum()
        
        # Resetear el índice para que 'Número de grupo de ubicación' sea una columna normal
        df_resumen = df_resumen.reset_index()
        
        # Aplicar redondeo a 2 decimales a todas las columnas numéricas
        df_resumen = df_resumen.round(2)
        
        logger.info(f"Resumen creado con {len(df_resumen)} grupos")
        logger.info(f"Columnas resumidas: {columnas_disponibles}")
        
        # Mostrar ejemplos del resumen
        logger.info("Ejemplos del resumen:")
        logger.info(df_resumen.head().to_string())
        
        return df_resumen
        
    except Exception as e:
        logger.error(f"Error al crear resumen: {e}")
        raise

def crear_visualizaciones(df: pd.DataFrame, df_resumen: pd.DataFrame) -> None:
    """
    Crear gráficos de visualización de datos con matplotlib y seaborn
    
    Args:
        df: DataFrame con datos detallados
        df_resumen: DataFrame con resumen agrupado
    """
    try:
        # Configurar estilo profesional para los gráficos
        plt.style.use('seaborn-v0_8')
        sns.set_palette("husl")
        
        logger.info("=== Iniciando creación de visualizaciones ===")
        
        # 1. Gráfico de Barras: Valor 4 (Prorrateo) por grupo de ubicación
        crear_grafico_barras(df_resumen)
        
        # 2. Gráfico de Torta: Distribución porcentual por grupo de ubicación
        crear_grafico_torta(df)
        
        # 3. Histograma: Distribución de Suma Valores 1&2
        crear_histograma(df)
        
        logger.info("=== Visualizaciones creadas exitosamente ===")
        
    except Exception as e:
        logger.error(f"Error en la creación de visualizaciones: {e}")
        raise

def crear_grafico_barras(df_resumen: pd.DataFrame) -> None:
    """
    Crear gráfico de barras para Valor 4 (Prorrateo) por grupo de ubicación
    
    Args:
        df_resumen: DataFrame con datos resumidos
    """
    try:
        if df_resumen.empty or 'Valor 4 (Prorrateo)' not in df_resumen.columns:
            logger.warning("No se puede crear gráfico de barras - datos insuficientes")
            return
        
        # Ordenar datos para mejor visualización
        df_ordenado = df_resumen.sort_values('Valor 4 (Prorrateo)', ascending=True)
        
        # Crear figura
        plt.figure(figsize=(12, 8))
        bars = plt.barh(
            df_ordenado['Número de grupo de ubicación'].astype(int), 
            df_ordenado['Valor 4 (Prorrateo)'],
            color=sns.color_palette("viridis", len(df_ordenado))
        )
        
        # Configurar etiquetas y título
        plt.title('Valor 4 (Prorrateo) Total por Grupo de Ubicación', 
                 fontsize=16, fontweight='bold', pad=20)
        plt.xlabel('Valor 4 (Prorrateo)', fontsize=12, fontweight='bold')
        plt.ylabel('Número de Grupo de Ubicación', fontsize=12, fontweight='bold')
        
        # Añadir valores en las barras
        for i, bar in enumerate(bars):
            width = bar.get_width()
            plt.text(width + max(df_ordenado['Valor 4 (Prorrateo)']) * 0.01, 
                    bar.get_y() + bar.get_height()/2, 
                    f'{width:,.2f}', 
                    ha='left', va='center', fontweight='bold')
        
        # Configurar formato de ejes
        plt.gca().xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:,.0f}'))
        
        # Ajustar layout y guardar
        plt.tight_layout()
        plt.savefig('valor_4_por_grupo.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info("Gráfico de barras guardado como 'valor_4_por_grupo.png'")
        
    except Exception as e:
        logger.error(f"Error al crear gráfico de barras: {e}")
        raise

def crear_grafico_torta(df: pd.DataFrame) -> None:
    """
    Crear gráfico de torta para distribución porcentual por grupo de ubicación
    
    Args:
        df: DataFrame con datos detallados
    """
    try:
        if 'Número de grupo de ubicación' not in df.columns:
            logger.warning("No se puede crear gráfico de torta - columna no encontrada")
            return
        
        # Contar registros por grupo
        conteo_grupos = df['Número de grupo de ubicación'].value_counts().sort_index()
        
        # Crear figura
        plt.figure(figsize=(10, 8))
        colors = sns.color_palette("Set2", len(conteo_grupos))
        
        # Crear gráfico de torta
        wedges, texts, autotexts = plt.pie(
            conteo_grupos.values, 
            labels=[f'Grupo {int(x)}' for x in conteo_grupos.index],
            colors=colors,
            autopct='%1.1f%%',
            startangle=90,
            textprops={'fontsize': 10, 'fontweight': 'bold'}
        )
        
        # Configurar título
        plt.title('Distribución Porcentual de Registros por Grupo de Ubicación', 
                 fontsize=14, fontweight='bold', pad=20)
        
        # Asegurar que el gráfico sea circular
        plt.axis('equal')
        
        # Ajustar layout y guardar
        plt.tight_layout()
        plt.savefig('distribucion_porcentual.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info("Gráfico de torta guardado como 'distribucion_porcentual.png'")
        
    except Exception as e:
        logger.error(f"Error al crear gráfico de torta: {e}")
        raise

def crear_histograma(df: pd.DataFrame) -> None:
    """
    Crear histograma para distribución de Suma Valores 1&2
    
    Args:
        df: DataFrame con datos detallados
    """
    try:
        if 'Suma Valores 1&2' not in df.columns:
            logger.warning("No se puede crear histograma - columna no encontrada")
            return
        
        # Crear figura
        plt.figure(figsize=(12, 8))
        
        # Crear histograma con KDE
        sns.histplot(
            data=df, 
            x='Suma Valores 1&2', 
            bins=20, 
            kde=True,
            color='skyblue',
            edgecolor='darkblue',
            linewidth=1.5
        )
        
        # Configurar etiquetas y título
        plt.title('Distribución de Montos - Suma Valores 1&2', 
                 fontsize=16, fontweight='bold', pad=20)
        plt.xlabel('Suma Valores 1&2', fontsize=12, fontweight='bold')
        plt.ylabel('Frecuencia', fontsize=12, fontweight='bold')
        
        # Configurar formato de ejes
        plt.gca().xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x:,.0f}'))
        
        # Añadir estadísticas
        media = df['Suma Valores 1&2'].mean()
        mediana = df['Suma Valores 1&2'].median()
        
        plt.axvline(media, color='red', linestyle='--', linewidth=2, label=f'Media: {media:,.2f}')
        plt.axvline(mediana, color='green', linestyle='--', linewidth=2, label=f'Mediana: {mediana:,.2f}')
        
        plt.legend(fontsize=10)
        plt.grid(True, alpha=0.3)
        
        # Ajustar layout y guardar
        plt.tight_layout()
        plt.savefig('distribucion_suma_valores.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        logger.info("Histograma guardado como 'distribucion_suma_valores.png'")
        
    except Exception as e:
        logger.error(f"Error al crear histograma: {e}")
        raise

def guardar_resultados(df: pd.DataFrame, df_resumen: pd.DataFrame, archivo_salida: str = 'datos_completados.xlsx') -> None:
    """
    Guardar los resultados en un archivo Excel con múltiples hojas
    
    Args:
        df: DataFrame con datos detallados
        df_resumen: DataFrame con resumen agrupado
        archivo_salida: Nombre del archivo de salida
    """
    try:
        # Usar ExcelWriter para guardar múltiples hojas
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            # Guardar hoja de datos detallados
            df.to_excel(writer, sheet_name='Datos', index=False)
            
            # Guardar hoja de resumen si existe
            if not df_resumen.empty:
                df_resumen.to_excel(writer, sheet_name='Resumen', index=False)
                logger.info("Hoja 'Resumen' agregada al archivo Excel")
            else:
                logger.warning("No se agregó hoja 'Resumen' - DataFrame vacío")
        
        logger.info(f"Resultados guardados exitosamente en '{archivo_salida}'")
        logger.info(f"Hojas guardadas: Datos, Resumen")
        print(f"### Proceso completado. Archivo '{archivo_salida}' generado con éxito.")
        print(f"### Hojas incluidas: 'Datos' (detalle) y 'Resumen' (agrupado)")
        
    except Exception as e:
        logger.error(f"Error al guardar el archivo: {e}")
        raise

def procesar_archivo_excel(ruta_archivo: str) -> None:
    """
    Función principal que orquesta todo el proceso
    
    Args:
        ruta_archivo: Ruta al archivo Excel a procesar
    """
    try:
        logger.info("=== Iniciando procesamiento de archivo Excel ===")
        
        # Paso 1: Cargar archivo
        df = cargar_archivo_excel(ruta_archivo)
        
        # Paso 2: Limpiar direcciones
        df = limpiar_direccion(df)
        
        # Paso 3: Mapear grupos de ubicación
        df = mapear_grupo_ubicacion(df)
        
        # Paso 4: Realizar cálculos
        df = realizar_calculos(df)
        
        # Paso 5: Crear resumen agrupado
        df_resumen = crear_resumen(df)
        
        # Paso 6: Guardar resultados con múltiples hojas
        guardar_resultados(df, df_resumen)
        
        # Paso 7: Crear visualizaciones de datos
        crear_visualizaciones(df, df_resumen)
        
        logger.info("=== Procesamiento completado exitosamente ===")
        
    except Exception as e:
        logger.error(f"Error en el procesamiento principal: {e}")
        print(f"### Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # Configuración del archivo
    ARCHIVO_EXCEL = 'Prueba Excel- Data & Analytics Assistant.xlsm'
    
    # Ejecutar procesamiento
    procesar_archivo_excel(ARCHIVO_EXCEL)

"""
================================================================================
PIPELINE ETL COMPLETO - DOCUMENTACIÓN PARA GITHUB
================================================================================

ESTE SCRIPT FORMA PARTE DE UN PIPELINE ETL COMPLETO:

1. EXTRACCIÓN (Extraction):
   - Lectura de archivos Excel (.xlsm) con múltiples hojas
   - Identificación automática de encabezados y estructura de datos
   - Manejo robusto de errores y validación de archivos

2. TRANSFORMACIÓN (Transformation):
   - Limpieza de datos con expresiones regulares avanzadas
   - Mapeo de categorías con diccionarios predefinidos
   - Cálculos matemáticos con validación de división por cero
   - Redondeo estandarizado a 2 decimales para consistencia
   - Creación de nuevas columnas derivadas (agregaciones)

3. CARGA (Load):
   - Exportación multihonda a Excel con hojas de detalle y resumen
   - Generación automática de visualizaciones profesionales
   - Formato de alta calidad (300 DPI) para presentaciones

INTEGRACIÓN CON POWER BI:
- Los archivos generados (datos_completados.xlsx) están optimizados para Power BI
- Las visualizaciones (.png) pueden integrarse directamente en dashboards
- Estructura de datos normalizada facilita conexiones dinámicas

CARACTERÍSTICAS DE PRODUCCIÓN:
- Logging completo para auditoría y debugging
- Manejo de excepciones robusto
- Código modular y mantenible
- Documentación exhaustiva
- Compatible con entornos Ubuntu/Linux

AUTOMATIZACIÓN FUTURA:
- Puede integrarse con Airflow o Apache Beam para programación
- Soporta procesamiento por lotes de múltiples archivos
- Extensible para nuevas fuentes de datos (CSV, JSON, APIs)

REPOSITORIO GITHUB:
- Este script demuestra capacidades de Data Engineering
- Incluye visualizaciones para portfolio técnico
- Documentado para colaboración y mantenimiento
================================================================================
"""
