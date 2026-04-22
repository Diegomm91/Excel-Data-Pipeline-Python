FROM python:3.11-slim

WORKDIR /app

# Instalamos herramientas básicas
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# 1. Copiamos los requerimientos e instalamos
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# 2. Copiamos el resto de los archivos
# IMPORTANTE: Los nombres deben ser EXACTOS (mayúsculas y minúsculas)
COPY app.py .
COPY procesar_excel_profesional.py .
COPY datos_completados.xlsx .

# Exponemos el puerto de Streamlit
EXPOSE 8501

# Comando para ejecutar la aplicación
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]