from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st


# ---------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ---------------------------------------------------------

st.set_page_config(
    page_title="Dashboard de Compras CAPEX",
    page_icon="📊",
    layout="wide",
)


# ---------------------------------------------------------
# ESTILOS
# ---------------------------------------------------------

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }

        [data-testid="stMetric"] {
            background-color: #EAF2F8;
            border-left: 6px solid #00A6A6;
            padding: 16px;
            border-radius: 6px;
        }

        [data-testid="stMetricLabel"] {
            color: #1F4E78;
            font-weight: 700;
        }

        h1, h2, h3 {
            color: #17365D;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# CONSTANTES
# ---------------------------------------------------------

MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

ORDEN_MESES = list(MESES.values())


# ---------------------------------------------------------
# FUNCIONES
# ---------------------------------------------------------

def limpiar_texto(serie):
    """
    Convierte una columna en texto, elimina espacios
    y reemplaza valores vacíos.
    """
    return (
        serie.fillna("Sin información")
        .astype(str)
        .str.strip()
        .replace("", "Sin información")
    )


@st.cache_data(show_spinner=False)
def cargar_archivo(archivo):
    """
    Lee la hoja Data del archivo Excel.

    header=5 significa que la fila 6 del Excel contiene
    los encabezados.
    """
    archivo.seek(0)

    df = pd.read_excel(
        archivo,
        sheet_name="Data",
        header=5,
        engine="openpyxl",
    )

    # Elimina columnas totalmente vacías
    df = df.dropna(axis=1, how="all")

    # Limpia espacios en los encabezados
    df.columns = [str(col).strip() for col in df.columns]

    return df


def localizar_columna(df, nombres_posibles):
    """
    Busca una columna utilizando diferentes nombres posibles.
    Esto evita errores por espacios o pequeñas variaciones.
    """
    columnas_normalizadas = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    for nombre in nombres_posibles:
        nombre_normalizado = nombre.strip().lower()

        if nombre_normalizado in columnas_normalizadas:
            return columnas_normalizadas[nombre_normalizado]

    return None


def preparar_datos(df):
    """
    Identifica las columnas necesarias, limpia las fechas
    y crea los campos utilizados por el dashboard.
    """

    columnas = {
        "tienda": localizar_columna(
            df,
            ["Proyecto", "Tienda", "Tienda / Proyecto"],
        ),
        "centro": localizar_columna(
            df,
            ["Centro"],
        ),
        "supranumero": localizar_columna(
            df,
            ["Supranumero", "Supranúmero"],
        ),
        "agrupador": localizar_columna(
            df,
            ["Agrupador"],
        ),
        "detalle": localizar_columna(
            df,
            ["Detalles de compra", "Detalle de compra"],
        ),
        "cantidad": localizar_columna(
            df,
            ["Cant.", "Cantidad"],
        ),
        "monto_usd": localizar_columna(
            df,
            [
                "Ppto. Orden Interna [USD]",
                "Monto USD",
            ],
        ),
        "fecha_compra": localizar_columna(
            df,
            ["Compra", "Compra "],
        ),
    }

    columnas_faltantes = [
        nombre
        for nombre, columna in columnas.items()
        if columna is None
    ]

    if columnas_faltantes:
        faltantes = ", ".join(columnas_faltantes)

        raise ValueError(
            f"No se encontraron estas columnas: {faltantes}"
        )

    resultado = pd.DataFrame()

    resultado["Tienda"] = limpiar_texto(df[columnas["tienda"]])
    resultado["Centro"] = limpiar_texto(df[columnas["centro"]])
    resultado["Supranúmero"] = limpiar_texto(
        df[columnas["supranumero"]]
    )
    resultado["Agrupador"] = limpiar_texto(
        df[columnas["agrupador"]]
    )
    resultado["Detalle de compra"] = limpiar_texto(
        df[columnas["detalle"]]
    )

    resultado["Cantidad"] = pd.to_numeric(
        df[columnas["cantidad"]],
        errors="coerce",
    ).fillna(0)

    resultado["Monto USD"] = pd.to_numeric(
        df[columnas["monto_usd"]],
        errors="coerce",
    ).fillna(0)

    resultado["Fecha de compra"] = pd.to_datetime(
        df[columnas["fecha_compra"]],
        errors="coerce",
    )

    # Se consideran válidas fechas entre 2020 y 2100.
    fecha_valida = (
        resultado["Fecha de compra"].notna()
        & resultado["Fecha de compra"].dt.year.between(2020, 2100)
    )

    resultado["Fecha válida"] = fecha_valida

    resultado["Estado fecha"] = "OK"

    resultado.loc[
        ~fecha_valida & resultado["Monto USD"].gt(0),
        "Estado fecha",
    ] = "Revisar fecha"

    resultado.loc[
        ~fecha_valida & resultado["Monto USD"].le(0),
        "Estado fecha",
    ] = "Sin fecha"

    resultado["Año"] = resultado["Fecha de compra"].dt.year
    resultado["Mes número"] = resultado["Fecha de compra"].dt.month
    resultado["Mes"] = resultado["Mes número"].map(MESES)

    # Mantiene únicamente filas con información relevante.
    resultado = resultado[
        resultado["Tienda"].ne("Sin información")
        | resultado["Detalle de compra"].ne("Sin información")
        | resultado["Monto USD"].ne(0)
    ].copy()

    return resultado


def aplicar_filtro_multiple(df, columna, seleccion):
    """
    Aplica un filtro solo cuando existe al menos
    una selección.
    """
    if seleccion:
        return df[df[columna].isin(seleccion)]

    return df


def crear_excel_descarga(df):
    """
    Genera un Excel descargable con los datos filtrados.
    """
    salida = BytesIO()

    columnas = [
        "Tienda",
        "Centro",
        "Supranúmero",
        "Agrupador",
        "Detalle de compra",
        "Cantidad",
        "Monto USD",
        "Fecha de compra",
        "Año",
        "Mes",
        "Estado fecha",
    ]

    with pd.ExcelWriter(
        salida,
        engine="xlsxwriter",
        datetime_format="dd/mm/yyyy",
    ) as writer:

        df[columnas].to_excel(
            writer,
            sheet_name="Compras filtradas",
            index=False,
        )

        workbook = writer.book
        worksheet = writer.sheets["Compras filtradas"]

        formato_encabezado = workbook.add_format(
            {
                "bold": True,
                "font_color": "white",
                "bg_color": "#17365D",
                "align": "center",
                "border": 0,
            }
        )

        formato_moneda = workbook.add_format(
            {
                "num_format": '$#,##0;$#,##0;-'
            }
        )

        formato_cantidad = workbook.add_format(
            {
                "num_format": '#,##0.00'
            }
        )

        for columna_numero, columna_nombre in enumerate(columnas):
            worksheet.write(
                0,
                columna_numero,
                columna_nombre,
                formato_encabezado,
            )

        worksheet.set_column("A:A", 25)
        worksheet.set_column("B:B", 15)
        worksheet.set_column("C:C", 30)
        worksheet.set_column("D:D", 35)
        worksheet.set_column("E:E", 55)
        worksheet.set_column("F:F", 13, formato_cantidad)
        worksheet.set_column("G:G", 16, formato_moneda)
        worksheet.set_column("H:H", 16)
        worksheet.set_column("I:J", 12)
        worksheet.set_column("K:K", 18)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(
            0,
            0,
            len(df),
            len(columnas) - 1,
        )

    salida.seek(0)

    return salida.getvalue()


# ---------------------------------------------------------
# ENCABEZADO
# ---------------------------------------------------------

st.title("Dashboard de planificación de compras")

st.caption(
    "Consolidado mensual por tienda, centro, agrupador "
    "y supranúmero. Montos expresados en USD."
)


# ---------------------------------------------------------
# CARGA DEL ARCHIVO
# ---------------------------------------------------------

archivo = st.file_uploader(
    "Carga el archivo de planificación de compras",
    type=["xlsm", "xlsx"],
    help="El archivo debe contener una hoja llamada Data.",
)

if archivo is None:
    st.info(
        "Carga el archivo Excel para comenzar el análisis."
    )
    st.stop()


# ---------------------------------------------------------
# LECTURA Y VALIDACIÓN
# ---------------------------------------------------------

try:
    df_original = cargar_archivo(archivo)
    df = preparar_datos(df_original)

except ValueError as error:
    st.error(str(error))
    st.stop()

except Exception as error:
    st.error(
        "No fue posible leer el archivo. "
        "Verifica que tenga una hoja llamada Data."
    )

    with st.expander("Ver detalle técnico"):
        st.exception(error)

    st.stop()


# ---------------------------------------------------------
# FILTROS
# ---------------------------------------------------------

st.sidebar.header("Filtros de análisis")

tiendas = sorted(
    df["Tienda"]
    .dropna()
    .unique()
    .tolist()
)

centros = sorted(
    df["Centro"]
    .dropna()
    .unique()
    .tolist()
)

agrupadores = sorted(
    df["Agrupador"]
    .dropna()
    .unique()
    .tolist()
)

anios = sorted(
    df.loc[df["Fecha válida"], "Año"]
    .dropna()
    .astype(int)
    .unique()
    .tolist()
)

filtro_tienda = st.sidebar.multiselect(
    "Tienda o proyecto",
    options=tiendas,
    placeholder="Todas las tiendas",
)

filtro_centro = st.sidebar.multiselect(
    "Centro",
    options=centros,
    placeholder="Todos los centros",
)

filtro_agrupador = st.sidebar.multiselect(
    "Agrupador",
    options=agrupadores,
    placeholder="Todos los agrupadores",
)

filtro_anio = st.sidebar.multiselect(
    "Año",
    options=anios,
    default=anios,
)

# Primero filtra toda la base para calcular pendientes.
df_general = aplicar_filtro_multiple(
    df,
    "Tienda",
    filtro_tienda,
)

df_general = aplicar_filtro_multiple(
    df_general,
    "Centro",
    filtro_centro,
)

df_general = aplicar_filtro_multiple(
    df_general,
    "Agrupador",
    filtro_agrupador,
)

pendientes_fecha = df_general[
    df_general["Estado fecha"].eq("Revisar fecha")
].copy()

# La planificación mensual utiliza fechas válidas.
df_filtrado = df_general[
    df_general["Fecha válida"]
].copy()

if filtro_anio:
    df_filtrado = df_filtrado[
        df_filtrado["Año"].isin(filtro_anio)
    ]


# ---------------------------------------------------------
# INDICADORES
# ---------------------------------------------------------

monto_total = df_filtrado["Monto USD"].sum()
lineas_compra = len(df_filtrado)
cantidad_activos = df_filtrado["Cantidad"].sum()
cantidad_pendientes = len(pendientes_fecha)

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "Monto planificado",
    f"US${monto_total:,.0f}",
)

col2.metric(
    "Líneas de compra",
    f"{lineas_compra:,.0f}",
)

col3.metric(
    "Cantidad de activos",
    f"{cantidad_activos:,.0f}",
)

col4.metric(
    "Fechas pendientes",
    f"{cantidad_pendientes:,.0f}",
)


if cantidad_pendientes > 0:
    st.warning(
        f"Existen {cantidad_pendientes:,.0f} registros con monto "
        "que no tienen una fecha de compra válida. Estos registros "
        "no están incluidos en el calendario mensual."
    )


# ---------------------------------------------------------
# VALIDACIÓN DE RESULTADOS
# ---------------------------------------------------------

if df_filtrado.empty:
    st.warning(
        "No existen compras con fechas válidas para "
        "los filtros seleccionados."
    )
    st.stop()


# ---------------------------------------------------------
# RESUMEN MENSUAL
# ---------------------------------------------------------

resumen_mensual = (
    df_filtrado
    .groupby(
        ["Mes número", "Mes"],
        as_index=False,
        observed=True,
    )
    .agg(
        Monto_USD=("Monto USD", "sum"),
        Lineas=("Detalle de compra", "size"),
        Cantidad=("Cantidad", "sum"),
    )
)

calendario = pd.DataFrame(
    {
        "Mes número": range(1, 13),
        "Mes": ORDEN_MESES,
    }
)

resumen_mensual = calendario.merge(
    resumen_mensual,
    on=["Mes número", "Mes"],
    how="left",
)

resumen_mensual[
    ["Monto_USD", "Lineas", "Cantidad"]
] = resumen_mensual[
    ["Monto_USD", "Lineas", "Cantidad"]
].fillna(0)


# ---------------------------------------------------------
# GRÁFICOS PRINCIPALES
# ---------------------------------------------------------

st.subheader("Plan mensual de compras")

grafico_mensual = px.bar(
    resumen_mensual,
    x="Mes",
    y="Monto_USD",
    text_auto=".2s",
    color="Monto_USD",
    color_continuous_scale=[
        "#DDEBF7",
        "#5B9BD5",
        "#17365D",
    ],
    labels={
        "Mes": "",
        "Monto_USD": "Monto USD",
    },
)

grafico_mensual.update_layout(
    coloraxis_showscale=False,
    height=430,
    margin=dict(l=10, r=10, t=20, b=10),
)

grafico_mensual.update_traces(
    hovertemplate=(
        "<b>%{x}</b><br>"
        "Monto: US$%{y:,.2f}"
        "<extra></extra>"
    )
)

st.plotly_chart(
    grafico_mensual,
    width="stretch",
)


# ---------------------------------------------------------
# DISTRIBUCIÓN POR SUPRANÚMERO
# ---------------------------------------------------------

st.subheader("Distribución por supranúmero")

resumen_supranumero = (
    df_filtrado
    .groupby(
        "Supranúmero",
        as_index=False,
        observed=True,
    )
    .agg(
        Monto_USD=("Monto USD", "sum"),
        Cantidad=("Cantidad", "sum"),
        Lineas=("Detalle de compra", "size"),
    )
    .sort_values(
        "Monto_USD",
        ascending=True,
    )
)

grafico_supranumero = px.bar(
    resumen_supranumero.tail(20),
    x="Monto_USD",
    y="Supranúmero",
    orientation="h",
    text_auto=".2s",
    color="Monto_USD",
    color_continuous_scale=[
        "#DDEBF7",
        "#00A6A6",
        "#17365D",
    ],
    labels={
        "Monto_USD": "Monto USD",
        "Supranúmero": "",
    },
)

grafico_supranumero.update_layout(
    coloraxis_showscale=False,
    height=650,
    margin=dict(l=10, r=10, t=20, b=10),
)

grafico_supranumero.update_traces(
    hovertemplate=(
        "<b>%{y}</b><br>"
        "Monto: US$%{x:,.2f}"
        "<extra></extra>"
    )
)

st.plotly_chart(
    grafico_supranumero,
    width="stretch",
)


# ---------------------------------------------------------
# MATRIZ SUPRANÚMERO POR MES
# ---------------------------------------------------------

st.subheader("Monto por supranúmero y mes")

matriz = pd.pivot_table(
    df_filtrado,
    index="Supranúmero",
    columns="Mes",
    values="Monto USD",
    aggfunc="sum",
    fill_value=0,
)

# Garantiza el orden correcto de los meses.
for mes in ORDEN_MESES:
    if mes not in matriz.columns:
        matriz[mes] = 0

matriz = matriz[ORDEN_MESES]

matriz["Total"] = matriz.sum(axis=1)

matriz = matriz.sort_values(
    "Total",
    ascending=False,
)

columnas_formato = {
    columna: st.column_config.NumberColumn(
        columna,
        format="US$ %.2f",
    )
    for columna in matriz.columns
}

st.dataframe(
    matriz.style.background_gradient(
        cmap="Blues",
        subset=ORDEN_MESES,
    ),
    width="stretch",
    height=600,
    column_config=columnas_formato,
)


# ---------------------------------------------------------
# DETALLE DE COMPRAS
# ---------------------------------------------------------

st.subheader("Detalle de activos planificados")

detalle = df_filtrado[
    [
        "Tienda",
        "Centro",
        "Agrupador",
        "Supranúmero",
        "Detalle de compra",
        "Cantidad",
        "Monto USD",
        "Fecha de compra",
        "Mes",
    ]
].sort_values(
    ["Fecha de compra", "Tienda", "Supranúmero"]
)

st.dataframe(
    detalle,
    width="stretch",
    hide_index=True,
    height=500,
    column_config={
        "Cantidad": st.column_config.NumberColumn(
            "Cantidad",
            format="%.2f",
        ),
        "Monto USD": st.column_config.NumberColumn(
            "Monto USD",
            format="US$ %.2f",
        ),
        "Fecha de compra": st.column_config.DateColumn(
            "Fecha de compra",
            format="DD/MM/YYYY",
        ),
    },
)


# ---------------------------------------------------------
# REGISTROS CON FECHAS PENDIENTES
# ---------------------------------------------------------

if not pendientes_fecha.empty:

    with st.expander(
        f"Revisar registros con fecha pendiente "
        f"({len(pendientes_fecha):,.0f})"
    ):

        st.dataframe(
            pendientes_fecha[
                [
                    "Tienda",
                    "Centro",
                    "Agrupador",
                    "Supranúmero",
                    "Detalle de compra",
                    "Cantidad",
                    "Monto USD",
                    "Fecha de compra",
                    "Estado fecha",
                ]
            ],
            width="stretch",
            hide_index=True,
            column_config={
                "Monto USD": st.column_config.NumberColumn(
                    "Monto USD",
                    format="US$ %.2f",
                ),
            },
        )


# ---------------------------------------------------------
# DESCARGA
# ---------------------------------------------------------

st.subheader("Descargar información")

archivo_descarga = crear_excel_descarga(df_filtrado)

st.download_button(
    label="Descargar compras filtradas en Excel",
    data=archivo_descarga,
    file_name="compras_planificadas_filtradas.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)