Data Pipeline: Limpieza, Procesamiento y Visualización de Activos
Este proyecto demuestra un flujo de trabajo completo (ETL) para la gestión de datos inmobiliarios y financieros. El objetivo es transformar un archivo de Excel crudo con datos "sucios" (direcciones mal formateadas, períodos de tiempo variables) en un reporte analítico listo para la toma de decisiones.

🛠️ Tecnologías Utilizadas
Python 3.12 (Procesamiento de datos)

Pandas & Openpyxl (Manipulación de DataFrames y motores de Excel)

Matplotlib & Seaborn (Generación de visualizaciones automáticas)

Power BI (Dashboard interactivo final)

Ubuntu Linux (Entorno de desarrollo)

📋 Características del Proyecto
El script de Python realiza las siguientes operaciones automáticas:

Limpieza de Direcciones mediante Regex: * Eliminación de caracteres especiales (#, @, &, etc.).

Truncado de rangos numéricos (ej: 43-49 pasa a ser 43).

Remoción de etiquetas de unidades y suites (Unit 2, Suite 6).

Mapeo de Categorías: Transformación de IDs numéricos a nombres descriptivos de riesgos (Earthquake, Flood, etc.).

Cálculos Financieros:

Suma de valores base.

Normalización temporal: Ajuste de importes de períodos variables a una base anual de 12 meses.

Prorrateo final con redondeo a 2 decimales para precisión contable.

Generación de Reportes:

Creación de una hoja de Resumen con totales agrupados.

Exportación de gráficos estadísticos en formato .png.

🚀 Cómo Ejecutarlo (en Ubuntu)
Clonar el repositorio:

Bash
git clone https://github.com/tu-usuario/nombre-del-repo.git
cd nombre-del-repo
Configurar el entorno virtual:

Bash
python3 -m venv .venv
source .venv/bin/activate
pip install pandas openpyxl matplotlib seaborn
Correr el pipeline:

Bash
python3 procesar_datos.py
📊 Visualizaciones Generadas
El script genera automáticamente insights visuales antes de la carga en Power BI:

Distribución de Prorrateo: Comparativa de costos por tipo de ubicación.

Frecuencia de Riesgos: Análisis de la composición de la cartera.

📈 Dashboard Final (Power BI)

El archivo datos_completados.xlsx generado sirve como fuente de verdad para el dashboard interactivo, permitiendo filtrar por ubicación y visualizar los valores ajustados en tiempo real.
