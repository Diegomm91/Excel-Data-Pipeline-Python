#!/usr/bin/env python3
"""
Dashboard Web con Streamlit para Visualización de Datos
Autor: Desarrollador Diego Martin
Fecha: 2026-04-22
Descripción: Aplicación web interactiva para el análisis del pipeline de datos procesados
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import warnings

# Configuración inicial
warnings.filterwarnings('ignore')

# Configuración de página
st.set_page_config(
    page_title="Dashboard de Datos - Analytics Pipeline",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Funciones de utilidad
def cargar_datos():
    """
    Cargar datos desde el archivo Excel procesado
    
    Returns:
        tuple: (df_datos, df_resumen) o (None, None) si hay error
    """
    try:
        # Cargar ambas hojas del Excel
        df_datos = pd.read_excel('datos_completados.xlsx', sheet_name='Datos', engine='openpyxl')
        df_resumen = pd.read_excel('datos_completados.xlsx', sheet_name='Resumen', engine='openpyxl')
        
        # Limpiar nombres de columnas si es necesario
        df_datos = df_datos.dropna(how='all')
        df_resumen = df_resumen.dropna(how='all')
        
        return df_datos, df_resumen
        
    except FileNotFoundError:
        st.error("### :warning: Archivo de datos no encontrado")
        st.info("""
        **Solución:**
        
        1. Ejecute primero el script de procesamiento:
        ```bash
        python procesar_excel_profesional.py
        ```
        
        2. Luego actualice esta página para ver los datos.
        """)
        return None, None
    except Exception as e:
        st.error(f"### :x: Error al cargar los datos")
        st.error(f"Error detallado: {str(e)}")
        return None, None

def calcular_kpis(df_datos, df_resumen):
    """
    Calcular KPIs principales del dashboard
    
    Args:
        df_datos: DataFrame con datos detallados
        df_resumen: DataFrame con resumen agrupado
        
    Returns:
        dict: Diccionario con KPIs calculados
    """
    if df_datos is None or df_resumen is None:
        return {}
    
    kpis = {
        'total_registros': len(df_datos),
        'suma_valor_4': df_resumen['Valor 4 (Prorrateo)'].sum() if 'Valor 4 (Prorrateo)' in df_resumen.columns else 0,
        'promedio_valor_3': df_datos['Valor 3 ajustado a 12 meses'].mean() if 'Valor 3 ajustado a 12 meses' in df_datos.columns else 0,
        'total_ubicaciones': df_datos['Nombre de grupo de ubicación'].nunique() if 'Nombre de grupo de ubicación' in df_datos.columns else 0
    }
    
    return kpis

def crear_grafico_distribucion_costos(df_resumen):
    """
    Crear gráfico interactivo de distribución de costos
    
    Args:
        df_resumen: DataFrame con datos resumidos
        
    Returns:
        plotly.graph_objects.Figure: Gráfico interactivo
    """
    if df_resumen is None or 'Valor 4 (Prorrateo)' not in df_resumen.columns:
        return None
    
    # Ordenar datos para mejor visualización
    df_ordenado = df_resumen.sort_values('Valor 4 (Prorrateo)', ascending=True)
    
    # Mapear números de grupo a nombres
    mapeo_nombres = {
        1: "Earthquake", 2: "Windstorm", 3: "Flood", 
        4: "Hail", 5: "Fire", 6: "Other"
    }
    
    df_ordenado['Nombre Grupo'] = df_ordenado['Número de grupo de ubicación'].map(mapeo_nombres)
    
    # Crear gráfico de barras horizontal
    fig = go.Figure(data=[
        go.Bar(
            x=df_ordenado['Valor 4 (Prorrateo)'],
            y=df_ordenado['Nombre Grupo'],
            orientation='h',
            marker=dict(
                color=df_ordenado['Valor 4 (Prorrateo)'],
                colorscale='Viridis',
                showscale=True
            ),
            text=df_ordenado['Valor 4 (Prorrateo)'].apply(lambda x: f'{x:,.2f}'),
            textposition='outside',
            textfont=dict(size=12, color='white')
        )
    ])
    
    fig.update_layout(
        title={
            'text': 'Distribución de Costos por Tipo de Riesgo',
            'x': 0.5,
            'xanchor': 'center',
            'font': dict(size=18, color='white')
        },
        xaxis_title='Valor 4 (Prorrateo)',
        yaxis_title='Tipo de Riesgo',
        height=500,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        margin=dict(l=20, r=20, t=60, b=20)
    )
    
    return fig

def crear_grafico_proporcion_riesgos(df_datos):
    """
    Crear gráfico de proporción de riesgos
    
    Args:
        df_datos: DataFrame con datos detallados
        
    Returns:
        plotly.graph_objects.Figure: Gráfico interactivo
    """
    if df_datos is None or 'Nombre de grupo de ubicación' not in df_datos.columns:
        return None
    
    # Contar registros por tipo de riesgo
    conteo_riesgos = df_datos['Nombre de grupo de ubicación'].value_counts()
    
    # Crear gráfico de torta
    fig = go.Figure(data=[
        go.Pie(
            labels=conteo_riesgos.index,
            values=conteo_riesgos.values,
            hole=0.3,
            marker=dict(colors=px.colors.qualitative.Set3),
            textinfo='label+percent',
            textfont=dict(size=12),
            hovertemplate='<b>%{label}</b><br>Registros: %{value}<br>Porcentaje: %{percent}<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title={
            'text': 'Proporción de Registros por Tipo de Riesgo',
            'x': 0.5,
            'xanchor': 'center',
            'font': dict(size=18, color='white')
        },
        height=400,
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        margin=dict(l=20, r=20, t=60, b=20),
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.01
        )
    )
    
    return fig

# Aplicación principal
def main():
    """
    Función principal de la aplicación Streamlit
    """
    # CSS personalizado para tema oscuro
    st.markdown("""
    <style>
        .main {
            background-color: #1e1e1e;
            color: white;
        }
        .stApp {
            background-color: #1e1e1e;
        }
        .css-1d391kg {
            background-color: #2e2e2e;
        }
        .css-1v0mbdj {
            background-color: #2e2e2e;
        }
        .metric-card {
            background-color: #2e2e2e;
            padding: 20px;
            border-radius: 10px;
            border-left: 5px solid #1f77b4;
            margin: 10px 0;
        }
        .stDataFrame {
            background-color: #2e2e2e;
        }
        .stTable {
            background-color: #2e2e2e;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Título principal
    st.title(":bar_chart: Dashboard de Analytics Pipeline")
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header(":gear: Configuración")
        
        st.markdown("""
        ### :information_source: Acerca del Proyecto
        
        Este dashboard visualiza los resultados del pipeline ETL para análisis 
        de datos de riesgos y ubicaciones. Los datos son procesados automáticamente 
        mediante scripts de Python con técnicas avanzadas de limpieza y cálculo.
        
        **Características:**
        - :white_check_mark: Limpieza de datos con regex
        - :white_check_mark: Cálculos matemáticos automatizados
        - :white_check_mark: Visualizaciones interactivas
        - :white_check_mark: Exportación a múltiples formatos
        """)
        
        st.markdown("---")
        
        # Cargar datos
        df_datos, df_resumen = cargar_datos()
        
        if df_datos is not None and df_resumen is not None:
            # Filtros interactivos
            st.subheader(":filter: Filtros")
            
            # Filtro por tipo de riesgo
            if 'Nombre de grupo de ubicación' in df_datos.columns:
                riesgos_disponibles = ['Todos'] + df_datos['Nombre de grupo de ubicación'].unique().tolist()
                riesgo_seleccionado = st.selectbox(
                    "Filtrar por Tipo de Riesgo:",
                    options=riesgos_disponibles,
                    index=0
                )
            else:
                riesgo_seleccionado = 'Todos'
            
            # Filtro por rango de valores
            if 'Valor 4 (Prorrateo)' in df_datos.columns:
                min_valor = float(df_datos['Valor 4 (Prorrateo)'].min())
                max_valor = float(df_datos['Valor 4 (Prorrateo)'].max())
                
                rango_valores = st.slider(
                    "Rango de Valor 4 (Prorrateo):",
                    min_value=min_valor,
                    max_value=max_valor,
                    value=(min_valor, max_valor),
                    format="%.2f"
                )
            else:
                rango_valores = (0, 0)
    
    # Contenido principal
    if df_datos is not None and df_resumen is not None:
        # Aplicar filtros
        datos_filtrados = df_datos.copy()
        
        if riesgo_seleccionado != 'Todos':
            datos_filtrados = datos_filtrados[
                datos_filtrados['Nombre de grupo de ubicación'] == riesgo_seleccionado
            ]
        
        if 'Valor 4 (Prorrateo)' in df_datos.columns:
            datos_filtrados = datos_filtrados[
                (datos_filtrados['Valor 4 (Prorrateo)'] >= rango_valores[0]) &
                (datos_filtrados['Valor 4 (Prorrateo)'] <= rango_valores[1])
            ]
        
        # KPIs
        kpis = calcular_kpis(datos_filtrados, df_resumen)
        
        st.subheader(":chart_with_upwards_trend: KPIs Principales")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                label=":clipboard: Total Registros",
                value=f"{kpis.get('total_registros', 0):,}",
                delta=None
            )
        
        with col2:
            st.metric(
                label=":dollar: Suma Valor 4",
                value=f"${kpis.get('suma_valor_4', 0):,.2f}",
                delta=None
            )
        
        with col3:
            st.metric(
                label=":chart: Promedio Valor 3",
                value=f"${kpis.get('promedio_valor_3', 0):,.2f}",
                delta=None
            )
        
        with col4:
            st.metric(
                label=":location: Ubicaciones",
                value=f"{kpis.get('total_ubicaciones', 0)}",
                delta=None
            )
        
        st.markdown("---")
        
        # Visualizaciones
        st.subheader(":bar_chart: Visualizaciones Interactivas")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Distribución de Costos por Riesgo")
            fig_costos = crear_grafico_distribucion_costos(df_resumen)
            if fig_costos:
                st.plotly_chart(fig_costos, use_container_width=True)
        
        with col2:
            st.markdown("#### Proporción de Riesgos")
            fig_riesgos = crear_grafico_proporcion_riesgos(df_datos)
            if fig_riesgos:
                st.plotly_chart(fig_riesgos, use_container_width=True)
        
        st.markdown("---")
        
        # Tabla de datos
        st.subheader(":table: Datos Detallados")
        
        # Mostrar información del filtro
        if riesgo_seleccionado != 'Todos' or rango_valores != (min_valor, max_valor):
            st.info(f":information_source: Mostrando {len(datos_filtrados)} registros con los filtros aplicados")
        else:
            st.info(f":information_source: Mostrando todos los {len(datos_filtrados)} registros")
        
        # Tabla con datos filtrados
        columnas_mostrar = [
            'Número de grupo de ubicación', 'Nombre de grupo de ubicación', 
            'Dirección depurada', 'Valor 1', 'Valor 2', 'Valor 3',
            'Suma Valores 1&2', 'Valor 3 ajustado a 12 meses', 'Valor 4 (Prorrateo)'
        ]
        
        columnas_disponibles = [col for col in columnas_mostrar if col in datos_filtrados.columns]
        
        if columnas_disponibles:
            # Formatear columnas numéricas para mejor visualización
            datos_mostrar = datos_filtrados[columnas_disponibles].copy()
            
            columnas_numericas = ['Valor 1', 'Valor 2', 'Valor 3', 'Suma Valores 1&2', 
                                 'Valor 3 ajustado a 12 meses', 'Valor 4 (Prorrateo)']
            
            for col in columnas_numericas:
                if col in datos_mostrar.columns:
                    datos_mostrar[col] = datos_mostrar[col].round(2)
            
            st.dataframe(
                datos_mostrar,
                use_container_width=True,
                height=400
            )
            
            # Opción de descarga
            csv = datos_mostrar.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=":download: Descargar Datos Filtrados (CSV)",
                data=csv,
                file_name='datos_filtrados_dashboard.csv',
                mime='text/csv'
            )
        
        # Footer
        st.markdown("---")
        st.markdown("""
        <div style='text-align: center; color: #888;'>
            <p>Dashboard Analytics Pipeline © 2026 | Desarrollado con Streamlit & Python</p>
            <p>:gear: Pipeline ETL: Extracción Transformación Carga Visualización</p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
