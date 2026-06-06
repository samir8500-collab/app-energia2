# Dashboard Streamlit - Análisis Energía EPM

Aplicación en Streamlit para analizar tarifas de energía EPM desde un archivo Excel.

## Archivos del proyecto

```text
streamlit_app.py
requirements.txt
README.md
```

## Qué hace la app

La aplicación permite cargar un archivo Excel y analizar únicamente la hoja **Energia #**.

Incluye:

- Limpieza automática de encabezados combinados.
- Indicadores generales.
- Gráficos de barras.
- Gráficos de torta.
- Gráficos de dispersión.
- Tendencias con puntos.
- Predicción general.
- Predicción de una serie seleccionada.
- Correlación simple, sin matriz extensa.
- KMeans para agrupar tarifas según nivel, crecimiento, volatilidad y tendencia.
- Descarga de Excel con resultados.
- Descarga de reporte HTML.

## Cómo subir a GitHub

1. Crea un repositorio en GitHub.
2. Sube estos tres archivos:

```text
streamlit_app.py
requirements.txt
README.md
```

3. No subas solamente el ZIP. Streamlit necesita ver el archivo `streamlit_app.py` directamente en el repositorio.

## Cómo desplegar en Streamlit Community Cloud

En Streamlit Community Cloud crea una nueva app y configura:

```text
Repository: tu_usuario/tu_repositorio
Branch: main
Main file path: streamlit_app.py
```

Luego presiona **Deploy**.

## Cómo usar la app

1. Abre la app desplegada en Streamlit.
2. En la barra izquierda carga el Excel.
3. Ajusta los meses a proyectar.
4. Ajusta la cantidad de grupos KMeans.
5. Selecciona la serie que quieres analizar.
6. Revisa los gráficos.
7. Descarga el Excel o el reporte HTML si lo necesitas.

## Librerías usadas

Las dependencias están en `requirements.txt`:

```text
streamlit
pandas
numpy
openpyxl
plotly
scikit-learn
```

## Ejecución local opcional

Si quieres probar la app en tu computador:

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
