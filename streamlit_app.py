# ============================================================
# DASHBOARD STREAMLIT - ANÁLISIS GRÁFICO TARIFAS ENERGÍA EPM
# Adaptado desde Colab para ejecutar en Streamlit Community Cloud
# Hoja esperada: Energia #
# ============================================================

import io
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score

warnings.filterwarnings("ignore")

# ============================================================
# CONFIGURACIÓN STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Análisis Energía EPM",
    page_icon="⚡",
    layout="wide"
)

# ============================================================
# ESTILOS
# ============================================================

st.markdown(
    """
    <style>
        .main-title {
            font-size: 36px;
            font-weight: 800;
            margin-bottom: 0px;
        }
        .subtitle {
            font-size: 16px;
            color: #555;
            margin-bottom: 20px;
        }
        .info-box {
            background: #f7f7f7;
            padding: 20px;
            border-radius: 14px;
            border: 1px solid #ddd;
            margin-bottom: 20px;
        }
        .success-box {
            background: #eefbea;
            padding: 22px;
            border-radius: 14px;
            border: 1px solid #c8e6c9;
            margin-top: 20px;
            margin-bottom: 20px;
        }
        .blue-box {
            background: #eef6ff;
            padding: 18px;
            border-radius: 12px;
            border: 1px solid #c9e2ff;
            margin-top: 20px;
            margin-bottom: 20px;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# FUNCIONES BASE
# ============================================================

MESES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}


def limpiar_texto(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def quitar_acentos(texto):
    texto = str(texto).lower().strip()
    reemplazos = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n",
    }
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    return texto


def normalizar_mes(x):
    x = quitar_acentos(limpiar_texto(x))
    return MESES.get(x, np.nan)


def limpiar_numero(x):
    """Convierte valores tipo moneda/texto a número.

    Soporta formatos como:
    - 1234.56
    - 1,234.56
    - 1.234,56
    - $ 1.234,56
    """
    if isinstance(x, pd.Series):
        return np.nan

    if pd.isna(x):
        return np.nan

    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)

    x = str(x).strip()
    x = x.replace("$", "")
    x = x.replace(" ", "")
    x = x.replace("\xa0", "")

    if x == "":
        return np.nan

    # Si tiene coma y punto, identifica cuál parece ser decimal.
    if "," in x and "." in x:
        if x.rfind(",") > x.rfind("."):
            # Formato 1.234,56
            x = x.replace(".", "")
            x = x.replace(",", ".")
        else:
            # Formato 1,234.56
            x = x.replace(",", "")
    elif "," in x and "." not in x:
        x = x.replace(",", ".")

    try:
        return float(x)
    except Exception:
        return np.nan


def buscar_hoja_energia(archivo_bytes):
    excel = pd.ExcelFile(io.BytesIO(archivo_bytes), engine="openpyxl")
    hojas = excel.sheet_names

    for hoja in hojas:
        if quitar_acentos(hoja) == "energia #":
            return hoja

    for hoja in hojas:
        h = quitar_acentos(hoja)
        if "energia" in h and "#" in h:
            return hoja

    raise ValueError("No se encontró la hoja 'Energia #' en el archivo.")


@st.cache_data(show_spinner=False)
def cargar_energia(archivo_bytes):
    hoja = buscar_hoja_energia(archivo_bytes)

    raw = pd.read_excel(
        io.BytesIO(archivo_bytes),
        sheet_name=hoja,
        header=None,
        engine="openpyxl",
    )

    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    if raw.shape[0] < 5:
        raise ValueError("La hoja Energia # no tiene suficientes filas para procesar.")

    fila_grupo = raw.iloc[1].ffill()
    fila_rango = raw.iloc[2].ffill()
    fila_propiedad = raw.iloc[3].ffill()

    data = raw.iloc[4:].copy().reset_index(drop=True)
    data = data.dropna(how="all").reset_index(drop=True)

    anio = data.iloc[:, 0].ffill()
    mes = data.iloc[:, 1].ffill()

    base_fechas = pd.DataFrame({
        "Año": pd.to_numeric(anio, errors="coerce"),
        "Mes": mes,
    })

    base_fechas["Mes_num"] = base_fechas["Mes"].apply(normalizar_mes)
    base_fechas = base_fechas.dropna(subset=["Año", "Mes_num"])

    indices_validos = base_fechas.index

    data = data.loc[indices_validos].reset_index(drop=True)
    base_fechas = base_fechas.reset_index(drop=True)

    base_fechas["Año"] = base_fechas["Año"].astype(int)
    base_fechas["Mes_num"] = base_fechas["Mes_num"].astype(int)

    base_fechas["Fecha"] = pd.to_datetime(
        {
            "year": base_fechas["Año"],
            "month": base_fechas["Mes_num"],
            "day": 1,
        },
        errors="coerce",
    )

    registros = []

    for pos in range(2, raw.shape[1]):
        grupo = limpiar_texto(fila_grupo.iloc[pos])
        rango = limpiar_texto(fila_rango.iloc[pos])
        propiedad = limpiar_texto(fila_propiedad.iloc[pos])

        if grupo == "":
            grupo = "Sin grupo"
        if rango == "":
            rango = "Todo el consumo"
        if propiedad == "":
            propiedad = "Sin propiedad"

        serie_completa = f"{grupo} - {rango} - {propiedad}"

        valores = data.iloc[:, pos].apply(limpiar_numero)

        if valores.notna().sum() < 6:
            continue

        temp = base_fechas.copy()
        temp["Tarifa"] = valores.values
        temp["Grupo"] = grupo
        temp["Rango"] = rango
        temp["Propiedad"] = propiedad
        temp["Serie_completa"] = serie_completa
        temp["Columna_origen"] = pos

        temp = temp.dropna(subset=["Fecha", "Tarifa"])
        registros.append(temp)

    if len(registros) == 0:
        raise ValueError("No se encontraron columnas numéricas válidas en la hoja Energia #.")

    df = pd.concat(registros, ignore_index=True)
    df = df.sort_values(["Serie_completa", "Fecha"]).reset_index(drop=True)

    return df, hoja


def preparar_figura(fig):
    fig.update_layout(
        template="plotly_white",
        margin=dict(l=40, r=40, t=80, b=80),
    )
    return fig


def mostrar_figura(fig, titulo_reporte, figuras_reporte):
    fig = preparar_figura(fig)
    st.plotly_chart(fig, use_container_width=True)
    figuras_reporte.append((titulo_reporte, fig))


def proyectar_df(df_base, columna_valor="Tarifa", meses_futuros=6):
    df_modelo = df_base.sort_values("Fecha").copy()
    df_modelo = df_modelo.dropna(subset=[columna_valor])

    if len(df_modelo) < 6:
        raise ValueError("No hay datos suficientes para proyectar.")

    df_modelo["Periodo"] = np.arange(len(df_modelo))

    X = df_modelo[["Periodo"]]
    y = df_modelo[columna_valor]

    modelo = LinearRegression()
    modelo.fit(X, y)

    df_modelo["Tendencia_modelo"] = modelo.predict(X)
    r2 = r2_score(y, df_modelo["Tendencia_modelo"])

    ultimo_periodo = int(df_modelo["Periodo"].max())
    ultima_fecha_modelo = df_modelo["Fecha"].max()

    futuros = pd.DataFrame({
        "Periodo": np.arange(ultimo_periodo + 1, ultimo_periodo + meses_futuros + 1)
    })

    futuros["Fecha"] = pd.date_range(
        start=ultima_fecha_modelo + pd.DateOffset(months=1),
        periods=meses_futuros,
        freq="MS",
    )

    futuros["Proyeccion"] = modelo.predict(futuros[["Periodo"]])
    futuros["Proyeccion"] = futuros["Proyeccion"].clip(lower=0)

    return df_modelo, futuros, modelo, r2


def formato_pesos(x):
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"


def crear_excel_resultados(resultados):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resultados["df"].to_excel(writer, sheet_name="Base_limpia", index=False)
        resultados["df_anual"].to_excel(writer, sheet_name="Promedio_anual", index=False)
        resultados["df_mensual"].to_excel(writer, sheet_name="Tendencia_mensual", index=False)
        resultados["top_tarifas"].to_excel(writer, sheet_name="Top_tarifas", index=False)
        resultados["bottom_tarifas"].to_excel(writer, sheet_name="Tarifas_bajas", index=False)
        resultados["df_crecimientos"].to_excel(writer, sheet_name="Crecimientos", index=False)
        resultados["df_serie"].to_excel(writer, sheet_name="Serie_principal", index=False)
        resultados["proyeccion_serie"].to_excel(writer, sheet_name="Proyeccion_serie", index=False)
        resultados["proyeccion_general"].to_excel(writer, sheet_name="Proyeccion_general", index=False)

        if resultados["matriz_corr"] is not None and not resultados["matriz_corr"].empty:
            resultados["matriz_corr"].to_excel(writer, sheet_name="Matriz_correlacion")

        if resultados["df_pares_corr"] is not None and not resultados["df_pares_corr"].empty:
            resultados["df_pares_corr"].to_excel(writer, sheet_name="Top_correlaciones", index=False)

        if resultados["df_corr_modelo"] is not None and not resultados["df_corr_modelo"].empty:
            resultados["df_corr_modelo"].to_excel(writer, sheet_name="Modelo_correlacion")

    return output.getvalue()


def crear_html_reporte(resultados, figuras_reporte):
    hoja_usada = resultados["hoja_usada"]
    fecha_min = resultados["fecha_min"]
    fecha_max = resultados["fecha_max"]
    registros_total = resultados["registros_total"]
    series_total = resultados["series_total"]
    tarifa_promedio = resultados["tarifa_promedio"]
    tarifa_minima = resultados["tarifa_minima"]
    tarifa_maxima = resultados["tarifa_maxima"]
    tarifa_inicial_general = resultados["tarifa_inicial_general"]
    tarifa_final_general = resultados["tarifa_final_general"]
    crecimiento_general = resultados["crecimiento_general"]
    serie_principal = resultados["serie_principal"]
    tarifa_proyectada_final = resultados["tarifa_proyectada_final"]
    variacion_proyectada = resultados["variacion_proyectada"]
    meses_a_proyectar = resultados["meses_a_proyectar"]

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>Reporte gráfico energía EPM</title>
    </head>
    <body style="font-family:Arial, sans-serif; margin:30px;">

    <h1>Reporte gráfico tarifas energía EPM</h1>

    <h2>Problemática</h2>
    <p>
    Las tarifas de energía de EPM cambian mes a mes según estrato, rango de consumo
    y propiedad del medidor. El archivo original tiene encabezados combinados y columnas repetidas,
    lo que dificulta identificar aumentos, comparar tarifas, calcular variaciones, proyectar valores futuros
    y analizar qué tarifas se mueven de forma similar.
    </p>

    <h2>Indicadores generales</h2>
    <ul>
        <li>Hoja analizada: {hoja_usada}</li>
        <li>Fecha inicial: {fecha_min.strftime('%Y-%m')}</li>
        <li>Fecha final: {fecha_max.strftime('%Y-%m')}</li>
        <li>Registros limpios: {registros_total:,.0f}</li>
        <li>Series de tarifa: {series_total:,.0f}</li>
        <li>Tarifa promedio: ${tarifa_promedio:,.2f}</li>
        <li>Tarifa mínima: ${tarifa_minima:,.2f}</li>
        <li>Tarifa máxima: ${tarifa_maxima:,.2f}</li>
    </ul>
    """

    for titulo, fig in figuras_reporte:
        html += f"<h2>{titulo}</h2>"
        html += fig.to_html(full_html=False, include_plotlyjs="cdn")

    html += f"""
    <h2>Conclusión ejecutiva</h2>
    <p>
    La tarifa promedio general pasó de ${tarifa_inicial_general:,.2f} a ${tarifa_final_general:,.2f},
    con un crecimiento acumulado aproximado de {crecimiento_general:.2f}%.
    </p>

    <p>
    La serie principal fue: <b>{serie_principal}</b>.
    </p>

    <p>
    La tarifa proyectada a {meses_a_proyectar} meses es ${tarifa_proyectada_final:,.2f},
    con una variación estimada de {variacion_proyectada:.2f}%.
    </p>

    </body>
    </html>
    """

    return html.encode("utf-8")


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("⚡ Configuración")
st.sidebar.write("Carga el archivo Excel de tarifas EPM.")

archivo_subido = st.sidebar.file_uploader(
    "Archivo Excel",
    type=["xlsx", "xlsm", "xls"],
)

meses_a_proyectar = st.sidebar.slider(
    "Meses a proyectar",
    min_value=1,
    max_value=12,
    value=6,
    step=1,
)

mostrar_base_limpia = st.sidebar.checkbox("Mostrar base limpia", value=False)

# ============================================================
# PORTADA
# ============================================================

st.markdown('<div class="main-title">⚡ Análisis gráfico de tarifas de energía EPM</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Dashboard interactivo con tendencias, variaciones, proyecciones, correlaciones y reporte descargable.</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="info-box">
        <h3>Problemática</h3>
        <p>
        Las tarifas de energía de EPM cambian mes a mes según estrato, rango de consumo
        y propiedad del medidor. El archivo original tiene encabezados combinados y columnas repetidas,
        lo que dificulta identificar aumentos, comparar tarifas, calcular variaciones,
        proyectar valores futuros y analizar qué tarifas se mueven de forma similar.
        </p>
        <h3>Solución aplicada</h3>
        <p>
        La app toma únicamente la hoja <b>Energia #</b>, limpia la información por posición de columna
        y genera visualmente tendencias, comparativos, proyecciones, correlaciones y modelo de regresión.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

if archivo_subido is None:
    st.info("Carga el archivo Excel desde el panel izquierdo para iniciar el análisis.")
    st.stop()

# ============================================================
# CARGA DE DATOS
# ============================================================

try:
    archivo_bytes = archivo_subido.getvalue()

    with st.spinner("Leyendo y limpiando la hoja Energia #..."):
        df, hoja_usada = cargar_energia(archivo_bytes)

except Exception as e:
    st.error("No fue posible procesar el archivo.")
    st.exception(e)
    st.stop()

figuras_reporte = []

# ============================================================
# INDICADORES GENERALES
# ============================================================

fecha_min = df["Fecha"].min()
fecha_max = df["Fecha"].max()
tarifa_promedio = df["Tarifa"].mean()
tarifa_maxima = df["Tarifa"].max()
tarifa_minima = df["Tarifa"].min()
series_total = df["Serie_completa"].nunique()
registros_total = len(df)

st.subheader("Indicadores generales")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Hoja analizada", hoja_usada)
c2.metric("Fecha inicial", fecha_min.strftime("%Y-%m"))
c3.metric("Fecha final", fecha_max.strftime("%Y-%m"))
c4.metric("Registros limpios", f"{registros_total:,.0f}")

c5, c6, c7, c8 = st.columns(4)
c5.metric("Series de tarifa", f"{series_total:,.0f}")
c6.metric("Tarifa promedio", formato_pesos(tarifa_promedio))
c7.metric("Tarifa mínima", formato_pesos(tarifa_minima))
c8.metric("Tarifa máxima", formato_pesos(tarifa_maxima))

if mostrar_base_limpia:
    st.subheader("Base limpia")
    st.dataframe(df, use_container_width=True)

# ============================================================
# 1. PROMEDIO ANUAL
# ============================================================

st.subheader("1. Tarifa promedio anual")

df_anual = df.groupby("Año", as_index=False)["Tarifa"].mean()
df_anual["Tarifa"] = df_anual["Tarifa"].round(2)

fig_anual = px.bar(
    df_anual,
    x="Año",
    y="Tarifa",
    text="Tarifa",
    title="Tarifa promedio anual - Energía EPM",
)
fig_anual.update_traces(textposition="outside")
fig_anual.update_layout(xaxis_title="Año", yaxis_title="Tarifa promedio", height=520)
mostrar_figura(fig_anual, "Tarifa promedio anual", figuras_reporte)

# ============================================================
# 2. TENDENCIA MENSUAL GENERAL
# ============================================================

st.subheader("2. Tendencia mensual promedio con medias móviles")

df_mensual = df.groupby("Fecha", as_index=False)["Tarifa"].mean()
df_mensual["Variacion_mensual_%"] = df_mensual["Tarifa"].pct_change() * 100
df_mensual["Media_movil_3m"] = df_mensual["Tarifa"].rolling(3).mean()
df_mensual["Media_movil_6m"] = df_mensual["Tarifa"].rolling(6).mean()

fig_mensual = go.Figure()
fig_mensual.add_trace(go.Scatter(x=df_mensual["Fecha"], y=df_mensual["Tarifa"], mode="lines+markers", name="Tarifa promedio mensual"))
fig_mensual.add_trace(go.Scatter(x=df_mensual["Fecha"], y=df_mensual["Media_movil_3m"], mode="lines", name="Media móvil 3 meses"))
fig_mensual.add_trace(go.Scatter(x=df_mensual["Fecha"], y=df_mensual["Media_movil_6m"], mode="lines", name="Media móvil 6 meses"))
fig_mensual.update_layout(title="Tendencia mensual promedio con medias móviles", xaxis_title="Fecha", yaxis_title="Tarifa promedio", height=580)
mostrar_figura(fig_mensual, "Tendencia mensual promedio", figuras_reporte)

# ============================================================
# 3. VARIACIÓN MENSUAL %
# ============================================================

st.subheader("3. Variación mensual promedio %")

fig_var_general = px.bar(df_mensual, x="Fecha", y="Variacion_mensual_%", title="Variación mensual promedio %")
fig_var_general.update_layout(xaxis_title="Fecha", yaxis_title="Variación mensual %", height=520)
mostrar_figura(fig_var_general, "Variación mensual promedio", figuras_reporte)

# ============================================================
# 4. MAPA DE CALOR AÑO VS MES
# ============================================================

st.subheader("4. Mapa de calor de tarifa promedio por año y mes")

df_heat = df_mensual.copy()
df_heat["Año"] = df_heat["Fecha"].dt.year
df_heat["Mes"] = df_heat["Fecha"].dt.month
heat_pivot = df_heat.pivot_table(index="Año", columns="Mes", values="Tarifa", aggfunc="mean")

fig_heat = px.imshow(heat_pivot, text_auto=".1f", aspect="auto", title="Mapa de calor de tarifa promedio por año y mes")
fig_heat.update_layout(xaxis_title="Mes", yaxis_title="Año", height=520)
mostrar_figura(fig_heat, "Mapa de calor año mes", figuras_reporte)

# ============================================================
# 5. TENDENCIA POR GRUPO / ESTRATO
# ============================================================

st.subheader("5. Tendencia mensual por grupo / estrato")

df_grupo = df.groupby(["Fecha", "Grupo"], as_index=False)["Tarifa"].mean()
fig_grupo = px.line(df_grupo, x="Fecha", y="Tarifa", color="Grupo", title="Tendencia mensual por grupo / estrato")
fig_grupo.update_layout(xaxis_title="Fecha", yaxis_title="Tarifa promedio", height=650)
mostrar_figura(fig_grupo, "Tendencia por grupo", figuras_reporte)

# ============================================================
# 6. TENDENCIA POR PROPIEDAD
# ============================================================

st.subheader("6. Tendencia mensual por propiedad del medidor")

df_propiedad = df.groupby(["Fecha", "Propiedad"], as_index=False)["Tarifa"].mean()
fig_propiedad = px.line(df_propiedad, x="Fecha", y="Tarifa", color="Propiedad", title="Tendencia mensual por propiedad del medidor")
fig_propiedad.update_layout(xaxis_title="Fecha", yaxis_title="Tarifa promedio", height=600)
mostrar_figura(fig_propiedad, "Tendencia por propiedad", figuras_reporte)

# ============================================================
# 7. DISTRIBUCIÓN POR GRUPO
# ============================================================

st.subheader("7. Distribución de tarifas por grupo / estrato")

fig_box = px.box(df, x="Grupo", y="Tarifa", title="Distribución de tarifas por grupo / estrato")
fig_box.update_layout(xaxis_title="Grupo", yaxis_title="Tarifa", height=650)
mostrar_figura(fig_box, "Distribución por grupo", figuras_reporte)

# ============================================================
# 8. HISTOGRAMA GENERAL
# ============================================================

st.subheader("8. Distribución general de tarifas")

fig_hist = px.histogram(df, x="Tarifa", nbins=40, title="Distribución general de tarifas")
fig_hist.update_layout(xaxis_title="Tarifa", yaxis_title="Cantidad de registros", height=520)
mostrar_figura(fig_hist, "Histograma de tarifas", figuras_reporte)

# ============================================================
# 9. TOP TARIFAS ALTAS ÚLTIMO MES
# ============================================================

ultima_fecha = df["Fecha"].max()
df_ultimo = df[df["Fecha"] == ultima_fecha].copy().sort_values("Tarifa", ascending=False)
top_tarifas = df_ultimo.head(15)

st.subheader(f"9. Top 15 tarifas más altas - {ultima_fecha.strftime('%Y-%m')}")

fig_top = px.bar(top_tarifas, x="Tarifa", y="Serie_completa", orientation="h", text="Tarifa", title=f"Top 15 tarifas más altas - {ultima_fecha.strftime('%Y-%m')}")
fig_top.update_layout(xaxis_title="Tarifa", yaxis_title="Serie", height=750)
fig_top.update_yaxes(autorange="reversed")
mostrar_figura(fig_top, "Top tarifas más altas", figuras_reporte)

# ============================================================
# 10. TOP TARIFAS BAJAS ÚLTIMO MES
# ============================================================

bottom_tarifas = df_ultimo.sort_values("Tarifa", ascending=True).head(15)

st.subheader(f"10. Top 15 tarifas más bajas - {ultima_fecha.strftime('%Y-%m')}")

fig_bottom = px.bar(bottom_tarifas, x="Tarifa", y="Serie_completa", orientation="h", text="Tarifa", title=f"Top 15 tarifas más bajas - {ultima_fecha.strftime('%Y-%m')}")
fig_bottom.update_layout(xaxis_title="Tarifa", yaxis_title="Serie", height=750)
fig_bottom.update_yaxes(autorange="reversed")
mostrar_figura(fig_bottom, "Top tarifas más bajas", figuras_reporte)

# ============================================================
# 11. RANKING POR GRUPO ÚLTIMO MES
# ============================================================

st.subheader(f"11. Ranking de tarifa promedio por grupo - {ultima_fecha.strftime('%Y-%m')}")

df_ranking_grupo = df_ultimo.groupby("Grupo", as_index=False)["Tarifa"].mean().sort_values("Tarifa", ascending=False)
fig_ranking_grupo = px.bar(df_ranking_grupo, x="Grupo", y="Tarifa", text="Tarifa", title=f"Ranking de tarifa promedio por grupo - {ultima_fecha.strftime('%Y-%m')}")
fig_ranking_grupo.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_ranking_grupo.update_layout(xaxis_title="Grupo", yaxis_title="Tarifa promedio", height=560)
mostrar_figura(fig_ranking_grupo, "Ranking por grupo último mes", figuras_reporte)

# ============================================================
# 12. SERIES CON MAYOR CRECIMIENTO ACUMULADO
# ============================================================

st.subheader("12. Top 15 series con mayor crecimiento acumulado")

crecimientos = []
for serie, temp in df.groupby("Serie_completa"):
    temp = temp.sort_values("Fecha")
    if len(temp) >= 6:
        inicial = temp["Tarifa"].iloc[0]
        final = temp["Tarifa"].iloc[-1]
        if inicial > 0:
            crecimiento = ((final / inicial) - 1) * 100
            crecimientos.append({
                "Serie_completa": serie,
                "Tarifa_inicial": inicial,
                "Tarifa_final": final,
                "Crecimiento_%": crecimiento,
            })

df_crecimientos = pd.DataFrame(crecimientos)
if not df_crecimientos.empty:
    df_crecimientos = df_crecimientos.sort_values("Crecimiento_%", ascending=False)
    top_crecimientos = df_crecimientos.head(15)

    fig_crec = px.bar(top_crecimientos, x="Crecimiento_%", y="Serie_completa", orientation="h", text="Crecimiento_%", title="Top 15 series con mayor crecimiento acumulado")
    fig_crec.update_traces(texttemplate="%{text:.2f}%")
    fig_crec.update_layout(xaxis_title="Crecimiento acumulado %", yaxis_title="Serie", height=750)
    fig_crec.update_yaxes(autorange="reversed")
    mostrar_figura(fig_crec, "Top crecimiento acumulado", figuras_reporte)
else:
    st.warning("No hay datos suficientes para calcular crecimiento acumulado por serie.")
    top_crecimientos = pd.DataFrame()

# ============================================================
# 13. SERIE PRINCIPAL AUTOMÁTICA
# ============================================================

serie_principal = df_ultimo.iloc[0]["Serie_completa"]
df_serie = df[df["Serie_completa"] == serie_principal].copy().sort_values("Fecha")
df_serie["Variacion_mensual_%"] = df_serie["Tarifa"].pct_change() * 100
df_serie["Media_movil_3m"] = df_serie["Tarifa"].rolling(3).mean()

st.markdown(
    f"""
    <div class="blue-box">
        <h3>Serie principal seleccionada automáticamente</h3>
        <p>Se seleccionó la tarifa más alta del último mes disponible:</p>
        <h4>{serie_principal}</h4>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 14. TENDENCIA SERIE PRINCIPAL
# ============================================================

st.subheader("13. Tendencia histórica de la serie principal")

fig_serie = go.Figure()
fig_serie.add_trace(go.Scatter(x=df_serie["Fecha"], y=df_serie["Tarifa"], mode="lines+markers", name="Tarifa histórica"))
fig_serie.add_trace(go.Scatter(x=df_serie["Fecha"], y=df_serie["Media_movil_3m"], mode="lines", name="Media móvil 3 meses"))
fig_serie.update_layout(title="Tendencia histórica de la serie principal", xaxis_title="Fecha", yaxis_title="Tarifa", height=580)
mostrar_figura(fig_serie, "Serie principal tendencia", figuras_reporte)

# ============================================================
# 15. VARIACIÓN SERIE PRINCIPAL
# ============================================================

st.subheader("14. Variación mensual % de la serie principal")

fig_serie_var = px.bar(df_serie, x="Fecha", y="Variacion_mensual_%", title="Variación mensual % de la serie principal")
fig_serie_var.update_layout(xaxis_title="Fecha", yaxis_title="Variación mensual %", height=520)
mostrar_figura(fig_serie_var, "Serie principal variación mensual", figuras_reporte)

# ============================================================
# 16. PROYECCIÓN GENERAL Y SERIE PRINCIPAL
# ============================================================

st.subheader(f"15. Proyección general promedio a {meses_a_proyectar} meses")

historico_general, proyeccion_general, modelo_general, r2_general = proyectar_df(
    df_mensual[["Fecha", "Tarifa"]],
    columna_valor="Tarifa",
    meses_futuros=meses_a_proyectar,
)

df_general_hist = historico_general[["Fecha", "Tarifa", "Tendencia_modelo"]].copy()
df_general_hist = df_general_hist.rename(columns={"Tarifa": "Histórico promedio", "Tendencia_modelo": "Tendencia modelo"})
df_general_melt = df_general_hist.melt(id_vars="Fecha", var_name="Tipo", value_name="Valor")

df_general_fut = proyeccion_general[["Fecha", "Proyeccion"]].copy()
df_general_fut = df_general_fut.rename(columns={"Proyeccion": "Valor"})
df_general_fut["Tipo"] = "Proyección promedio"

df_general_plot = pd.concat([df_general_melt, df_general_fut[["Fecha", "Tipo", "Valor"]]], ignore_index=True)

fig_proy_general = px.line(df_general_plot, x="Fecha", y="Valor", color="Tipo", markers=True, title=f"Proyección general promedio a {meses_a_proyectar} meses")
fig_proy_general.update_layout(xaxis_title="Fecha", yaxis_title="Tarifa promedio", height=600)
mostrar_figura(fig_proy_general, "Proyección general promedio", figuras_reporte)

st.subheader(f"16. Proyección de serie principal a {meses_a_proyectar} meses")

historico_serie, proyeccion_serie, modelo_serie, r2_serie = proyectar_df(
    df_serie[["Fecha", "Tarifa"]],
    columna_valor="Tarifa",
    meses_futuros=meses_a_proyectar,
)

df_hist_plot = historico_serie[["Fecha", "Tarifa", "Tendencia_modelo"]].copy()
df_hist_plot = df_hist_plot.rename(columns={"Tarifa": "Histórico", "Tendencia_modelo": "Tendencia modelo"})
df_hist_melt = df_hist_plot.melt(id_vars="Fecha", var_name="Tipo", value_name="Valor")

df_fut_plot = proyeccion_serie[["Fecha", "Proyeccion"]].copy()
df_fut_plot = df_fut_plot.rename(columns={"Proyeccion": "Valor"})
df_fut_plot["Tipo"] = "Proyección"

df_total_proy = pd.concat([df_hist_melt, df_fut_plot[["Fecha", "Tipo", "Valor"]]], ignore_index=True)

fig_proy = px.line(df_total_proy, x="Fecha", y="Valor", color="Tipo", markers=True, title=f"Proyección de serie principal a {meses_a_proyectar} meses")
fig_proy.update_layout(xaxis_title="Fecha", yaxis_title="Tarifa", height=600)
mostrar_figura(fig_proy, "Proyección serie principal", figuras_reporte)

tarifa_actual = df_serie["Tarifa"].iloc[-1]
tarifa_proyectada_final = proyeccion_serie["Proyeccion"].iloc[-1]
variacion_proyectada = ((tarifa_proyectada_final / tarifa_actual) - 1) * 100 if tarifa_actual != 0 else 0

p1, p2, p3, p4 = st.columns(4)
p1.metric("Tarifa actual serie principal", formato_pesos(tarifa_actual))
p2.metric("Tarifa proyectada", formato_pesos(tarifa_proyectada_final))
p3.metric("Variación proyectada", f"{variacion_proyectada:.2f}%")
p4.metric("R² modelo serie", f"{r2_serie:.3f}")

# ============================================================
# 17. MATRIZ DE CORRELACIÓN
# ============================================================

st.subheader("17. Matriz de correlación - Top 25 series del último mes")

pivot = df.pivot_table(index="Fecha", columns="Serie_completa", values="Tarifa", aggfunc="mean").sort_index()
minimo_datos = max(6, int(len(pivot) * 0.5))
pivot = pivot.dropna(axis=1, thresh=minimo_datos)

matriz_corr = pd.DataFrame()
matriz_corr_top = pd.DataFrame()
df_pares_corr = pd.DataFrame()
df_corr_modelo = pd.DataFrame()

if pivot.shape[1] >= 2:
    matriz_corr = pivot.corr()
    series_top_corr = list(dict.fromkeys(df_ultimo.head(25)["Serie_completa"].tolist()))
    series_top_corr = [s for s in series_top_corr if s in matriz_corr.columns]

    if len(series_top_corr) >= 2:
        matriz_corr_top = matriz_corr.loc[series_top_corr, series_top_corr]
        fig_corr = px.imshow(matriz_corr_top, text_auto=".2f", aspect="auto", title="Matriz de correlación - Top 25 series del último mes")
        fig_corr.update_layout(height=900)
        mostrar_figura(fig_corr, "Matriz de correlación", figuras_reporte)
    else:
        st.warning("No hay suficientes series válidas para construir el mapa de correlación.")

    # Top correlaciones
    st.subheader("18. Top 20 relaciones más correlacionadas")

    pares = []
    cols = list(matriz_corr.columns)
    for i in range(len(cols)):
        for j in range(i + 1, len(cols)):
            corr_val = matriz_corr.iloc[i, j]
            if pd.notna(corr_val):
                pares.append({
                    "Serie_1": cols[i],
                    "Serie_2": cols[j],
                    "Correlacion": corr_val,
                })

    df_pares_corr = pd.DataFrame(pares)
    if not df_pares_corr.empty:
        df_pares_corr["Correlacion_abs"] = df_pares_corr["Correlacion"].abs()
        df_pares_corr = df_pares_corr.sort_values("Correlacion_abs", ascending=False)

        top_pares_corr = df_pares_corr.head(20).copy()
        top_pares_corr["Par"] = top_pares_corr["Serie_1"] + " ↔ " + top_pares_corr["Serie_2"]

        fig_top_corr = px.bar(top_pares_corr, x="Correlacion", y="Par", orientation="h", text="Correlacion", title="Top 20 relaciones más correlacionadas")
        fig_top_corr.update_traces(texttemplate="%{text:.3f}")
        fig_top_corr.update_layout(xaxis_title="Correlación Pearson", yaxis_title="Par de series", height=850)
        fig_top_corr.update_yaxes(autorange="reversed")
        mostrar_figura(fig_top_corr, "Top correlaciones", figuras_reporte)
    else:
        st.warning("No se pudieron calcular pares de correlación.")

    # Modelo de correlación lineal
    st.subheader("19. Modelo de correlación lineal entre dos tarifas")

    if len(series_top_corr) >= 2:
        variable_x = series_top_corr[0]
        variable_y = series_top_corr[1]
        df_corr_modelo = pivot[[variable_x, variable_y]].dropna().copy()

        if len(df_corr_modelo) >= 2:
            X = df_corr_modelo[[variable_x]]
            y = df_corr_modelo[variable_y]

            modelo_corr = LinearRegression()
            modelo_corr.fit(X, y)

            df_corr_modelo["Prediccion"] = modelo_corr.predict(X)
            r2_corr = r2_score(y, df_corr_modelo["Prediccion"])
            pearson = df_corr_modelo[variable_x].corr(df_corr_modelo[variable_y])
            pendiente = modelo_corr.coef_[0]
            intercepto = modelo_corr.intercept_

            fig_reg = px.scatter(df_corr_modelo, x=variable_x, y=variable_y, title="Modelo de correlación lineal entre dos tarifas")
            df_linea = df_corr_modelo.sort_values(variable_x)
            fig_reg.add_scatter(x=df_linea[variable_x], y=df_linea["Prediccion"], mode="lines", name="Línea de regresión")
            fig_reg.update_layout(height=650)
            mostrar_figura(fig_reg, "Modelo de correlación lineal", figuras_reporte)

            st.markdown(
                f"""
                <div class="info-box">
                    <h3>Resultado del modelo de correlación lineal</h3>
                    <p><b>Variable X:</b> {variable_x}</p>
                    <p><b>Variable Y:</b> {variable_y}</p>
                    <p><b>Correlación Pearson:</b> {pearson:.4f}</p>
                    <p><b>R² regresión:</b> {r2_corr:.4f}</p>
                    <p><b>Pendiente:</b> {pendiente:.4f}</p>
                    <p><b>Intercepto:</b> {intercepto:.4f}</p>
                    <p><b>Modelo:</b> {variable_y} = {intercepto:.2f} + {pendiente:.4f} * {variable_x}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.warning("No hay suficientes datos cruzados para el modelo de correlación lineal.")
else:
    st.warning("No hay suficientes series para calcular correlaciones.")

# ============================================================
# 20. COMPARACIÓN TOP 10 SERIES EN EL TIEMPO
# ============================================================

st.subheader("20. Tendencia histórica de las 10 tarifas más altas del último mes")

top_10_series = list(dict.fromkeys(df_ultimo.head(10)["Serie_completa"].tolist()))
df_top10_tiempo = df[df["Serie_completa"].isin(top_10_series)].copy()

fig_top10_tiempo = px.line(df_top10_tiempo, x="Fecha", y="Tarifa", color="Serie_completa", title="Tendencia histórica de las 10 tarifas más altas del último mes")
fig_top10_tiempo.update_layout(xaxis_title="Fecha", yaxis_title="Tarifa", height=700)
mostrar_figura(fig_top10_tiempo, "Tendencia top 10 series", figuras_reporte)

# ============================================================
# CONCLUSIÓN EJECUTIVA
# ============================================================

tarifa_inicial_general = df_mensual["Tarifa"].iloc[0]
tarifa_final_general = df_mensual["Tarifa"].iloc[-1]
crecimiento_general = ((tarifa_final_general / tarifa_inicial_general) - 1) * 100 if tarifa_inicial_general != 0 else 0

tarifa_inicial_serie = df_serie["Tarifa"].iloc[0]
tarifa_final_serie = df_serie["Tarifa"].iloc[-1]
crecimiento_serie = ((tarifa_final_serie / tarifa_inicial_serie) - 1) * 100 if tarifa_inicial_serie != 0 else 0

st.markdown(
    f"""
    <div class="success-box">
        <h2>Conclusión ejecutiva</h2>
        <p>
        La hoja analizada fue <b>{hoja_usada}</b>. El periodo disponible va desde
        <b>{fecha_min.strftime('%Y-%m')}</b> hasta <b>{fecha_max.strftime('%Y-%m')}</b>.
        </p>
        <p>
        Se limpiaron <b>{registros_total:,.0f}</b> registros y se identificaron
        <b>{series_total:,.0f}</b> series de tarifa.
        </p>
        <p>
        La tarifa promedio general pasó de <b>${tarifa_inicial_general:,.2f}</b> a
        <b>${tarifa_final_general:,.2f}</b>, con un crecimiento acumulado aproximado de
        <b>{crecimiento_general:.2f}%</b>.
        </p>
        <p>
        La serie principal analizada fue:<br>
        <b>{serie_principal}</b>
        </p>
        <p>
        Esta serie pasó de <b>${tarifa_inicial_serie:,.2f}</b> a
        <b>${tarifa_final_serie:,.2f}</b>, con un crecimiento acumulado de
        <b>{crecimiento_serie:.2f}%</b>.
        </p>
        <p>
        Según el modelo lineal, la tarifa proyectada a <b>{meses_a_proyectar}</b> meses es de
        <b>${tarifa_proyectada_final:,.2f}</b>, con una variación estimada de
        <b>{variacion_proyectada:.2f}%</b> frente al último valor histórico.
        </p>
        <p>
        La matriz de correlación evidencia que varias tarifas se mueven de forma similar,
        especialmente aquellas que dependen de la misma base tarifaria y cambian por factores
        de subsidio, contribución, estrato o tipo de usuario.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# DESCARGAS
# ============================================================

resultados = {
    "df": df,
    "hoja_usada": hoja_usada,
    "df_anual": df_anual,
    "df_mensual": df_mensual,
    "top_tarifas": top_tarifas,
    "bottom_tarifas": bottom_tarifas,
    "df_crecimientos": df_crecimientos,
    "df_serie": df_serie,
    "proyeccion_serie": proyeccion_serie,
    "proyeccion_general": proyeccion_general,
    "matriz_corr": matriz_corr,
    "df_pares_corr": df_pares_corr,
    "df_corr_modelo": df_corr_modelo,
    "fecha_min": fecha_min,
    "fecha_max": fecha_max,
    "registros_total": registros_total,
    "series_total": series_total,
    "tarifa_promedio": tarifa_promedio,
    "tarifa_minima": tarifa_minima,
    "tarifa_maxima": tarifa_maxima,
    "tarifa_inicial_general": tarifa_inicial_general,
    "tarifa_final_general": tarifa_final_general,
    "crecimiento_general": crecimiento_general,
    "serie_principal": serie_principal,
    "tarifa_proyectada_final": tarifa_proyectada_final,
    "variacion_proyectada": variacion_proyectada,
    "meses_a_proyectar": meses_a_proyectar,
}

excel_bytes = crear_excel_resultados(resultados)
html_bytes = crear_html_reporte(resultados, figuras_reporte)

st.subheader("Descargar resultados")

d1, d2 = st.columns(2)
with d1:
    st.download_button(
        label="⬇️ Descargar Excel del análisis",
        data=excel_bytes,
        file_name="analisis_grafico_energia_epm.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with d2:
    st.download_button(
        label="⬇️ Descargar reporte HTML",
        data=html_bytes,
        file_name="reporte_grafico_energia_epm.html",
        mime="text/html",
    )
