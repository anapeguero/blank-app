from io import BytesIO
from datetime import date
import re
import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

st.set_page_config(
    page_title="Dashboard Gerencial de Compras CAPEX",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }

        [data-testid="stMetric"] {
            background-color: #EAF2F8;
            border-left: 6px solid #00A6A6;
            padding: 15px;
            border-radius: 7px;
        }

        [data-testid="stMetricLabel"] {
            color: #1F4E78;
            font-weight: 700;
        }

        h1, h2, h3 {
            color: #17365D;
        }

        div[data-testid="stDownloadButton"] button {
            background-color: #17365D;
            color: white;
            border-radius: 6px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


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
HOY = pd.Timestamp(date.today())


# =========================================================
# FUNCIONES GENERALES
# =========================================================

def normalizar_texto(valor):
    valor = str(valor).strip().lower()

    valor = "".join(
        caracter
        for caracter in unicodedata.normalize("NFD", valor)
        if unicodedata.category(caracter) != "Mn"
    )

    return re.sub(r"\s+", " ", valor)


def localizar_columna(df, opciones):
    mapa = {
        normalizar_texto(columna): columna
        for columna in df.columns
    }

    for opcion in opciones:
        opcion_normalizada = normalizar_texto(opcion)

        if opcion_normalizada in mapa:
            return mapa[opcion_normalizada]

    return None


def obtener_serie(df, columna, valor_defecto=None):
    if columna is not None:
        return df[columna]

    return pd.Series(
        [valor_defecto] * len(df),
        index=df.index,
    )


def limpiar_texto(serie, valor_vacio="Sin información"):
    return (
        serie.fillna(valor_vacio)
        .astype(str)
        .str.strip()
        .replace("", valor_vacio)
        .replace("nan", valor_vacio)
        .replace("None", valor_vacio)
    )


def convertir_numero(serie):
    return pd.to_numeric(
        serie,
        errors="coerce",
    ).fillna(0)


def convertir_fecha(serie):
    fechas = pd.to_datetime(
        serie,
        errors="coerce",
    )

    fechas_validas = (
        fechas.notna()
        & fechas.dt.year.between(2020, 2100)
    )

    return fechas.where(fechas_validas)


@st.cache_data(show_spinner=False)
def cargar_archivo(contenido_archivo):
    df = pd.read_excel(
        BytesIO(contenido_archivo),
        sheet_name="Data",
        header=5,
        engine="openpyxl",
    )

    df = df.dropna(axis=1, how="all")
    df.columns = [
        str(columna).strip()
        for columna in df.columns
    ]

    return df


def preparar_datos(df):
    columnas = {
        "formato": localizar_columna(
            df,
            ["Formato"],
        ),
        "iniciativa": localizar_columna(
            df,
            ["Iniciativa"],
        ),
        "tienda": localizar_columna(
            df,
            ["Proyecto", "Tienda", "Tienda / Proyecto"],
        ),
        "centro": localizar_columna(
            df,
            ["Centro"],
        ),
        "estatus_iniciativa": localizar_columna(
            df,
            ["Estatus Inic", "Estatus iniciativa"],
        ),
        "prioridad_proyecto": localizar_columna(
            df,
            ["Prioridad del proyecto"],
        ),
        "estatus_compra": localizar_columna(
            df,
            ["Estatus de compra"],
        ),
        "prioridad_embarque": localizar_columna(
            df,
            ["Prioridad embarque"],
        ),
        "estado_tarea": localizar_columna(
            df,
            ["Estado de tarea", "Estado tarea"],
        ),
        "prioridad_tarea": localizar_columna(
            df,
            ["Prioridad tarea"],
        ),
        "su": localizar_columna(
            df,
            ["SU"],
        ),
        "supranumero": localizar_columna(
            df,
            ["Supranumero", "Supranúmero"],
        ),
        "partida": localizar_columna(
            df,
            ["Partidas", "Partida"],
        ),
        "agrupador": localizar_columna(
            df,
            ["Agrupador"],
        ),
        "tipo_capex": localizar_columna(
            df,
            ["Tipo de Capex", "Tipo CAPEX"],
        ),
        "detalle": localizar_columna(
            df,
            ["Detalles de compra", "Detalle de compra"],
        ),
        "cantidad": localizar_columna(
            df,
            ["Cant.", "Cantidad"],
        ),
        "moneda": localizar_columna(
            df,
            ["Mon", "Moneda"],
        ),
        "tasa": localizar_columna(
            df,
            ["Tasa"],
        ),
        "precio_unitario": localizar_columna(
            df,
            ["Precio Und.", "Precio unidad"],
        ),
        "valor_neto_dop": localizar_columna(
            df,
            ["Valor Neto [DOP]"],
        ),
        "presupuesto_dop": localizar_columna(
            df,
            ["Ppto. Orden Interna [DOP]"],
        ),
        "presupuesto_usd": localizar_columna(
            df,
            ["Ppto. Orden Interna [USD]", "Monto USD"],
        ),
        "periodo_1": localizar_columna(
            df,
            ["Per.P1"],
        ),
        "pago_1": localizar_columna(
            df,
            ["Pago 1"],
        ),
        "periodo_2": localizar_columna(
            df,
            ["Per.P2"],
        ),
        "pago_2": localizar_columna(
            df,
            ["Pago 2"],
        ),
        "periodo_3": localizar_columna(
            df,
            ["Per.P3"],
        ),
        "pago_3": localizar_columna(
            df,
            ["Pago 3"],
        ),
        "periodo_4": localizar_columna(
            df,
            ["Per.P4"],
        ),
        "pago_4": localizar_columna(
            df,
            ["Pago 4"],
        ),
        "periodo_5": localizar_columna(
            df,
            ["Per.P5"],
        ),
        "pago_5": localizar_columna(
            df,
            ["Pago 5"],
        ),
        "total_flujo": localizar_columna(
            df,
            ["Total Flujo Caja"],
        ),
        "lead_time": localizar_columna(
            df,
            ["Lead time", "Lead Time"],
        ),
        "apertura": localizar_columna(
            df,
            ["Apertura"],
        ),
        "compra": localizar_columna(
            df,
            ["Compra", "Compra "],
        ),
        "entrega": localizar_columna(
            df,
            ["Entrega"],
        ),
        "fin": localizar_columna(
            df,
            ["Fin"],
        ),
        "almacenaje": localizar_columna(
            df,
            ["Almacenaje"],
        ),
        "ubicacion_almacen": localizar_columna(
            df,
            ["Ubicación ALM", "Ubicacion ALM"],
        ),
        "tipo_almacen": localizar_columna(
            df,
            ["Tipo Alm"],
        ),
        "condicion_almacen": localizar_columna(
            df,
            ["Condición Alm", "Condicion Alm"],
        ),
    }

    obligatorias = [
        "tienda",
        "centro",
        "supranumero",
        "agrupador",
        "detalle",
        "cantidad",
        "presupuesto_usd",
        "compra",
    ]

    faltantes = [
        nombre
        for nombre in obligatorias
        if columnas[nombre] is None
    ]

    if faltantes:
        raise ValueError(
            "No se encontraron estas columnas obligatorias: "
            + ", ".join(faltantes)
        )

    resultado = pd.DataFrame(index=df.index)

    campos_texto = {
        "Formato": "formato",
        "Iniciativa": "iniciativa",
        "Tienda": "tienda",
        "Centro": "centro",
        "Estatus iniciativa": "estatus_iniciativa",
        "Prioridad proyecto": "prioridad_proyecto",
        "Estatus compra": "estatus_compra",
        "Prioridad embarque": "prioridad_embarque",
        "Estado tarea": "estado_tarea",
        "Prioridad tarea": "prioridad_tarea",
        "SU": "su",
        "Supranúmero": "supranumero",
        "Partida": "partida",
        "Agrupador": "agrupador",
        "Tipo CAPEX": "tipo_capex",
        "Detalle de compra": "detalle",
        "Moneda": "moneda",
        "Almacenaje": "almacenaje",
        "Ubicación almacén": "ubicacion_almacen",
        "Tipo almacén": "tipo_almacen",
        "Condición almacén": "condicion_almacen",
    }

    for salida, origen in campos_texto.items():
        resultado[salida] = limpiar_texto(
            obtener_serie(
                df,
                columnas[origen],
                "Sin información",
            )
        )

    campos_numericos = {
        "Cantidad": "cantidad",
        "Tasa": "tasa",
        "Precio unitario": "precio_unitario",
        "Valor neto DOP": "valor_neto_dop",
        "Presupuesto DOP": "presupuesto_dop",
        "Presupuesto USD": "presupuesto_usd",
        "Total flujo": "total_flujo",
        "Lead time": "lead_time",
    }

    for salida, origen in campos_numericos.items():
        resultado[salida] = convertir_numero(
            obtener_serie(
                df,
                columnas[origen],
                0,
            )
        )

    resultado["Fecha apertura"] = convertir_fecha(
        obtener_serie(
            df,
            columnas["apertura"],
        )
    )

    resultado["Fecha compra"] = convertir_fecha(
        obtener_serie(
            df,
            columnas["compra"],
        )
    )

    resultado["Fecha entrega"] = convertir_fecha(
        obtener_serie(
            df,
            columnas["entrega"],
        )
    )

    resultado["Fecha fin"] = convertir_fecha(
        obtener_serie(
            df,
            columnas["fin"],
        )
    )

    for numero in range(1, 6):
        resultado[f"Periodo {numero}"] = convertir_fecha(
            obtener_serie(
                df,
                columnas[f"periodo_{numero}"],
            )
        )

        resultado[f"Pago {numero}"] = convertir_numero(
            obtener_serie(
                df,
                columnas[f"pago_{numero}"],
                0,
            )
        )

    resultado["Total pagos calculado"] = sum(
        resultado[f"Pago {numero}"]
        for numero in range(1, 6)
    )

    resultado["Total flujo utilizado"] = resultado[
        "Total flujo"
    ].where(
        resultado["Total flujo"].ne(0),
        resultado["Total pagos calculado"],
    )

    resultado["Fecha válida"] = resultado[
        "Fecha compra"
    ].notna()

    resultado["Año compra"] = resultado[
        "Fecha compra"
    ].dt.year

    resultado["Mes número compra"] = resultado[
        "Fecha compra"
    ].dt.month

    resultado["Mes compra"] = resultado[
        "Mes número compra"
    ].map(MESES)

    resultado["Año entrega"] = resultado[
        "Fecha entrega"
    ].dt.year

    resultado["Mes número entrega"] = resultado[
        "Fecha entrega"
    ].dt.month

    resultado["Mes entrega"] = resultado[
        "Mes número entrega"
    ].map(MESES)

    resultado["Días para compra"] = (
        resultado["Fecha compra"] - HOY
    ).dt.days

    resultado["Días entrega a apertura"] = (
        resultado["Fecha apertura"]
        - resultado["Fecha entrega"]
    ).dt.days

    resultado["Semáforo"] = "Sin fecha"

    resultado.loc[
        resultado["Días para compra"] <= 15,
        "Semáforo",
    ] = "Crítico"

    resultado.loc[
        resultado["Días para compra"].between(16, 30),
        "Semáforo",
    ] = "Atención"

    resultado.loc[
        resultado["Días para compra"].between(31, 60),
        "Semáforo",
    ] = "Planificar"

    resultado.loc[
        resultado["Días para compra"] > 60,
        "Semáforo",
    ] = "En tiempo"

    estatus_normalizado = resultado[
        "Estatus compra"
    ].map(normalizar_texto)

    compra_completada = estatus_normalizado.str.contains(
        "realizada|completada|comprada|cerrada|finalizada",
        regex=True,
        na=False,
    )

    resultado.loc[
        compra_completada,
        "Semáforo",
    ] = "Completada"

    resultado["Diferencia flujo DOP"] = (
        resultado["Total flujo utilizado"]
        - resultado["Presupuesto DOP"]
    )

    resultado["Flujo coincide"] = (
        resultado["Diferencia flujo DOP"].abs() <= 1
    )

    resultado["Error fecha compra"] = (
        resultado["Fecha compra"].isna()
        & resultado["Presupuesto USD"].gt(0)
    )

    resultado["Error fecha entrega"] = (
        resultado["Fecha entrega"].isna()
        & resultado["Presupuesto USD"].gt(0)
    )

    resultado["Error monto cero"] = resultado[
        "Presupuesto USD"
    ].le(0)

    resultado["Error cantidad cero"] = resultado[
        "Cantidad"
    ].le(0)

    resultado["Error supranúmero"] = resultado[
        "Supranúmero"
    ].eq("Sin información")

    resultado["Error agrupador"] = resultado[
        "Agrupador"
    ].eq("Sin información")

    resultado["Error lead time"] = resultado[
        "Lead time"
    ].le(0)

    resultado["Compra posterior a apertura"] = (
        resultado["Fecha compra"].notna()
        & resultado["Fecha apertura"].notna()
        & (
            resultado["Fecha compra"]
            > resultado["Fecha apertura"]
        )
    )

    resultado["Entrega posterior a apertura"] = (
        resultado["Fecha entrega"].notna()
        & resultado["Fecha apertura"].notna()
        & (
            resultado["Fecha entrega"]
            > resultado["Fecha apertura"]
        )
    )

    resultado["Entrega anterior a compra"] = (
        resultado["Fecha entrega"].notna()
        & resultado["Fecha compra"].notna()
        & (
            resultado["Fecha entrega"]
            < resultado["Fecha compra"]
        )
    )

    resultado["Problema flujo"] = (
        resultado["Presupuesto DOP"].gt(0)
        & ~resultado["Flujo coincide"]
    )

    columnas_error = [
        "Error fecha compra",
        "Error fecha entrega",
        "Error monto cero",
        "Error cantidad cero",
        "Error supranúmero",
        "Error agrupador",
        "Error lead time",
        "Compra posterior a apertura",
        "Entrega posterior a apertura",
        "Entrega anterior a compra",
        "Problema flujo",
    ]

    resultado["Cantidad errores"] = resultado[
        columnas_error
    ].sum(axis=1)

    resultado["Calidad registro"] = resultado[
        "Cantidad errores"
    ].apply(
        lambda valor: "Correcto"
        if valor == 0
        else "Revisar"
    )

    resultado = resultado[
        resultado["Detalle de compra"].ne("Sin información")
        | resultado["Presupuesto USD"].ne(0)
        | resultado["Cantidad"].ne(0)
    ].copy()

    return resultado


def aplicar_multifiltro(df, columna, seleccion):
    if seleccion:
        return df[df[columna].isin(seleccion)]

    return df


def crear_excel_descarga(df, nombre_hoja):
    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="xlsxwriter",
        datetime_format="dd/mm/yyyy",
    ) as writer:
        df.to_excel(
            writer,
            sheet_name=nombre_hoja[:31],
            index=False,
        )

        workbook = writer.book
        worksheet = writer.sheets[nombre_hoja[:31]]

        formato_encabezado = workbook.add_format(
            {
                "bold": True,
                "font_color": "white",
                "bg_color": "#17365D",
                "align": "center",
            }
        )

        formato_moneda = workbook.add_format(
            {
                "num_format": '$#,##0.00;$#,##0.00;-'
            }
        )

        for numero_columna, nombre_columna in enumerate(
            df.columns
        ):
            worksheet.write(
                0,
                numero_columna,
                nombre_columna,
                formato_encabezado,
            )

            ancho = min(
                max(
                    len(str(nombre_columna)) + 3,
                    14,
                ),
                45,
            )

            worksheet.set_column(
                numero_columna,
                numero_columna,
                ancho,
            )

            if any(
                palabra in nombre_columna.lower()
                for palabra in [
                    "presupuesto",
                    "pago",
                    "flujo",
                    "monto",
                    "valor",
                    "precio",
                    "diferencia",
                ]
            ):
                worksheet.set_column(
                    numero_columna,
                    numero_columna,
                    18,
                    formato_moneda,
                )

        worksheet.freeze_panes(1, 0)

        if len(df) > 0:
            worksheet.autofilter(
                0,
                0,
                len(df),
                len(df.columns) - 1,
            )

    salida.seek(0)

    return salida.getvalue()


def mostrar_tabla(df, altura=480):
    configuracion = {}

    for columna in df.columns:
        if "Fecha" in columna or "Periodo" in columna:
            configuracion[columna] = (
                st.column_config.DateColumn(
                    columna,
                    format="DD/MM/YYYY",
                )
            )

        if any(
            palabra in columna
            for palabra in [
                "Presupuesto",
                "Pago",
                "Flujo",
                "Diferencia",
                "Precio",
                "Valor",
                "Monto",
            ]
        ):
            configuracion[columna] = (
                st.column_config.NumberColumn(
                    columna,
                    format="%.2f",
                )
            )

    st.dataframe(
        df,
        width="stretch",
        hide_index=True,
        height=altura,
        column_config=configuracion,
    )


# =========================================================
# ENCABEZADO Y CARGA
# =========================================================

st.title("Dashboard Gerencial de Compras CAPEX")

st.caption(
    "Planificación consolidada, seguimiento de riesgos, "
    "flujo de caja, entregas, almacenamiento y calidad de datos."
)

archivo = st.file_uploader(
    "Carga el archivo de planificación de compras",
    type=["xlsm", "xlsx"],
    help="El archivo debe contener una hoja llamada Data.",
)

if archivo is None:
    st.info(
        "Carga el archivo Excel para generar el dashboard."
    )
    st.stop()

try:
    df_original = cargar_archivo(
        archivo.getvalue()
    )

    df = preparar_datos(
        df_original
    )

except ValueError as error:
    st.error(str(error))
    st.stop()

except Exception as error:
    st.error(
        "No fue posible procesar el archivo. "
        "Verifica que la hoja Data tenga los encabezados en la fila 6."
    )

    with st.expander("Detalle técnico"):
        st.exception(error)

    st.stop()


# =========================================================
# FILTROS
# =========================================================

st.sidebar.header("Filtros generales")


def crear_opciones(columna):
    return sorted(
        df[columna]
        .dropna()
        .unique()
        .tolist()
    )


filtro_tienda = st.sidebar.multiselect(
    "Tienda o proyecto",
    crear_opciones("Tienda"),
    placeholder="Todas",
)

filtro_centro = st.sidebar.multiselect(
    "Centro",
    crear_opciones("Centro"),
    placeholder="Todos",
)

filtro_agrupador = st.sidebar.multiselect(
    "Agrupador",
    crear_opciones("Agrupador"),
    placeholder="Todos",
)

filtro_estatus = st.sidebar.multiselect(
    "Estatus de compra",
    crear_opciones("Estatus compra"),
    placeholder="Todos",
)

filtro_prioridad = st.sidebar.multiselect(
    "Prioridad del proyecto",
    crear_opciones("Prioridad proyecto"),
    placeholder="Todas",
)

opciones_anio = sorted(
    df["Año compra"]
    .dropna()
    .astype(int)
    .unique()
    .tolist()
)

filtro_anio = st.sidebar.multiselect(
    "Año de compra",
    opciones_anio,
    default=opciones_anio,
)

df_filtrado = df.copy()

filtros = [
    ("Tienda", filtro_tienda),
    ("Centro", filtro_centro),
    ("Agrupador", filtro_agrupador),
    ("Estatus compra", filtro_estatus),
    ("Prioridad proyecto", filtro_prioridad),
]

for columna, seleccion in filtros:
    df_filtrado = aplicar_multifiltro(
        df_filtrado,
        columna,
        seleccion,
    )

df_planificado = df_filtrado[
    df_filtrado["Fecha válida"]
].copy()

if filtro_anio:
    df_planificado = df_planificado[
        df_planificado["Año compra"].isin(
            filtro_anio
        )
    ]


# =========================================================
# PESTAÑAS
# =========================================================

(
    tab_resumen,
    tab_plan,
    tab_seguimiento,
    tab_flujo,
    tab_entregas,
    tab_calidad,
) = st.tabs(
    [
        "📊 Resumen ejecutivo",
        "📅 Plan mensual",
        "🚦 Seguimiento",
        "💵 Flujo de caja",
        "🚚 Entregas y almacén",
        "✅ Calidad de datos",
    ]
)


# =========================================================
# RESUMEN EJECUTIVO
# =========================================================

with tab_resumen:
    st.subheader("Resumen ejecutivo")

    criticas = df_filtrado[
        df_filtrado["Semáforo"].eq("Crítico")
    ]

    pendientes_fecha = df_filtrado[
        df_filtrado["Error fecha compra"]
    ]

    indicadores = st.columns(5)

    indicadores[0].metric(
        "Monto planificado",
        f"US${df_planificado['Presupuesto USD'].sum():,.0f}",
    )

    indicadores[1].metric(
        "Líneas de compra",
        f"{len(df_planificado):,.0f}",
    )

    indicadores[2].metric(
        "Cantidad de activos",
        f"{df_planificado['Cantidad'].sum():,.0f}",
    )

    indicadores[3].metric(
        "Tiendas involucradas",
        f"{df_planificado['Tienda'].nunique():,.0f}",
    )

    indicadores[4].metric(
        "Compras críticas",
        f"{len(criticas):,.0f}",
    )

    if len(pendientes_fecha) > 0:
        st.warning(
            f"Existen {len(pendientes_fecha):,.0f} registros "
            "con monto que no tienen una fecha de compra válida."
        )

    if df_planificado.empty:
        st.info(
            "No existen compras con fechas válidas "
            "para los filtros seleccionados."
        )

    else:
        resumen_mensual = (
            df_planificado
            .groupby(
                [
                    "Mes número compra",
                    "Mes compra",
                ],
                as_index=False,
            )
            .agg(
                Monto_USD=(
                    "Presupuesto USD",
                    "sum",
                )
            )
        )

        calendario = pd.DataFrame(
            {
                "Mes número compra": range(1, 13),
                "Mes compra": ORDEN_MESES,
            }
        )

        resumen_mensual = calendario.merge(
            resumen_mensual,
            on=[
                "Mes número compra",
                "Mes compra",
            ],
            how="left",
        ).fillna(0)

        grafico_mensual = px.bar(
            resumen_mensual,
            x="Mes compra",
            y="Monto_USD",
            text_auto=".2s",
            color="Monto_USD",
            color_continuous_scale=[
                "#DDEBF7",
                "#5B9BD5",
                "#17365D",
            ],
            title="Monto de compras por mes",
        )

        grafico_mensual.update_layout(
            coloraxis_showscale=False,
            height=430,
        )

        st.plotly_chart(
            grafico_mensual,
            width="stretch",
        )

        izquierda, derecha = st.columns(2)

        resumen_tienda = (
            df_planificado
            .groupby(
                "Tienda",
                as_index=False,
            )
            .agg(
                Monto_USD=(
                    "Presupuesto USD",
                    "sum",
                )
            )
            .sort_values("Monto_USD")
        )

        grafico_tienda = px.bar(
            resumen_tienda,
            x="Monto_USD",
            y="Tienda",
            orientation="h",
            color="Monto_USD",
            color_continuous_scale="Blues",
            title="Presupuesto por tienda",
        )

        grafico_tienda.update_layout(
            coloraxis_showscale=False,
            height=500,
        )

        izquierda.plotly_chart(
            grafico_tienda,
            width="stretch",
        )

        resumen_supranumero = (
            df_planificado
            .groupby(
                "Supranúmero",
                as_index=False,
            )
            .agg(
                Monto_USD=(
                    "Presupuesto USD",
                    "sum",
                )
            )
            .nlargest(
                15,
                "Monto_USD",
            )
            .sort_values(
                "Monto_USD"
            )
        )

        grafico_supranumero = px.bar(
            resumen_supranumero,
            x="Monto_USD",
            y="Supranúmero",
            orientation="h",
            color="Monto_USD",
            color_continuous_scale="Teal",
            title="Top 15 supranúmeros por inversión",
        )

        grafico_supranumero.update_layout(
            coloraxis_showscale=False,
            height=500,
        )

        derecha.plotly_chart(
            grafico_supranumero,
            width="stretch",
        )


# =========================================================
# PLAN MENSUAL
# =========================================================

with tab_plan:
    st.subheader(
        "Plan mensual por supranúmero"
    )

    if df_planificado.empty:
        st.info(
            "No existen datos planificados "
            "para los filtros seleccionados."
        )

    else:
        matriz = pd.pivot_table(
            df_planificado,
            index="Supranúmero",
            columns="Mes compra",
            values="Presupuesto USD",
            aggfunc="sum",
            fill_value=0,
        )

        for mes in ORDEN_MESES:
            if mes not in matriz.columns:
                matriz[mes] = 0

        matriz = matriz[ORDEN_MESES]
        matriz["Total"] = matriz.sum(axis=1)

        matriz = matriz.sort_values(
            "Total",
            ascending=False,
        )

        st.dataframe(
            matriz.style.background_gradient(
                cmap="Blues",
                subset=ORDEN_MESES,
            ),
            width="stretch",
            height=620,
            column_config={
                columna: st.column_config.NumberColumn(
                    columna,
                    format="US$ %.2f",
                )
                for columna in matriz.columns
            },
        )

        detalle_plan = df_planificado[
            [
                "Tienda",
                "Centro",
                "Agrupador",
                "Supranúmero",
                "Detalle de compra",
                "Cantidad",
                "Moneda",
                "Precio unitario",
                "Presupuesto USD",
                "Fecha compra",
                "Fecha entrega",
                "Estatus compra",
                "Prioridad proyecto",
                "Semáforo",
            ]
        ].sort_values(
            [
                "Fecha compra",
                "Tienda",
                "Supranúmero",
            ]
        )

        st.subheader(
            "Detalle de activos planificados"
        )

        mostrar_tabla(
            detalle_plan
        )

        st.download_button(
            "Descargar plan mensual filtrado",
            data=crear_excel_descarga(
                detalle_plan,
                "Plan mensual",
            ),
            file_name="plan_mensual_compras.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )


# =========================================================
# SEGUIMIENTO
# =========================================================

with tab_seguimiento:
    st.subheader(
        "Seguimiento y riesgo de compras"
    )

    vencidas = df_filtrado[
        (
            df_filtrado["Días para compra"] < 0
        )
        & ~df_filtrado["Semáforo"].eq(
            "Completada"
        )
    ]

    proximos_30 = df_filtrado[
        df_filtrado[
            "Días para compra"
        ].between(0, 30)
    ]

    proximos_60 = df_filtrado[
        df_filtrado[
            "Días para compra"
        ].between(31, 60)
    ]

    proximos_90 = df_filtrado[
        df_filtrado[
            "Días para compra"
        ].between(61, 90)
    ]

    indicadores = st.columns(4)

    indicadores[0].metric(
        "Compras vencidas",
        f"{len(vencidas):,.0f}",
        f"US${vencidas['Presupuesto USD'].sum():,.0f}",
        delta_color="inverse",
    )

    indicadores[1].metric(
        "Próximos 30 días",
        f"{len(proximos_30):,.0f}",
        f"US${proximos_30['Presupuesto USD'].sum():,.0f}",
    )

    indicadores[2].metric(
        "Entre 31 y 60 días",
        f"{len(proximos_60):,.0f}",
        f"US${proximos_60['Presupuesto USD'].sum():,.0f}",
    )

    indicadores[3].metric(
        "Entre 61 y 90 días",
        f"{len(proximos_90):,.0f}",
        f"US${proximos_90['Presupuesto USD'].sum():,.0f}",
    )

    resumen_semaforo = (
        df_filtrado
        .groupby(
            "Semáforo",
            as_index=False,
        )
        .agg(
            Monto_USD=(
                "Presupuesto USD",
                "sum",
            )
        )
    )

    colores_semaforo = {
        "Crítico": "#C00000",
        "Atención": "#ED7D31",
        "Planificar": "#FFD966",
        "En tiempo": "#70AD47",
        "Completada": "#5B9BD5",
        "Sin fecha": "#A5A5A5",
    }

    grafico_semaforo = px.bar(
        resumen_semaforo,
        x="Semáforo",
        y="Monto_USD",
        color="Semáforo",
        text_auto=".2s",
        color_discrete_map=colores_semaforo,
        title="Monto por nivel de riesgo",
    )

    grafico_semaforo.update_layout(
        showlegend=False,
        height=430,
    )

    st.plotly_chart(
        grafico_semaforo,
        width="stretch",
    )

    izquierda, derecha = st.columns(2)

    estado_compra = (
        df_filtrado
        .groupby(
            "Estatus compra",
            as_index=False,
        )
        .agg(
            Monto_USD=(
                "Presupuesto USD",
                "sum",
            )
        )
        .sort_values(
            "Monto_USD"
        )
    )

    grafico_estado = px.bar(
        estado_compra,
        x="Monto_USD",
        y="Estatus compra",
        orientation="h",
        color="Monto_USD",
        color_continuous_scale="Blues",
        title="Monto por estatus de compra",
    )

    grafico_estado.update_layout(
        coloraxis_showscale=False,
        height=430,
    )

    izquierda.plotly_chart(
        grafico_estado,
        width="stretch",
    )

    prioridad = (
        df_filtrado
        .groupby(
            "Prioridad proyecto",
            as_index=False,
        )
        .agg(
            Monto_USD=(
                "Presupuesto USD",
                "sum",
            )
        )
    )

    grafico_prioridad = px.pie(
        prioridad,
        names="Prioridad proyecto",
        values="Monto_USD",
        hole=0.55,
        title="Distribución por prioridad",
    )

    grafico_prioridad.update_layout(
        height=430
    )

    derecha.plotly_chart(
        grafico_prioridad,
        width="stretch",
    )

    tabla_critica = df_filtrado[
        df_filtrado[
            "Semáforo"
        ].isin(
            [
                "Crítico",
                "Atención",
            ]
        )
    ][
        [
            "Tienda",
            "Centro",
            "Agrupador",
            "Supranúmero",
            "Detalle de compra",
            "Presupuesto USD",
            "Fecha compra",
            "Días para compra",
            "Estatus compra",
            "Prioridad proyecto",
            "Semáforo",
        ]
    ].sort_values(
        [
            "Días para compra",
            "Presupuesto USD",
        ],
        ascending=[
            True,
            False,
        ],
    )

    st.subheader(
        "Compras críticas y próximas"
    )

    mostrar_tabla(
        tabla_critica
    )


# =========================================================
# FLUJO DE CAJA
# =========================================================

with tab_flujo:
    st.subheader(
        "Flujo de caja planificado"
    )

    lista_pagos = []

    for numero in range(1, 6):
        temporal = df_filtrado[
            [
                "Tienda",
                "Centro",
                "Agrupador",
                "Supranúmero",
                "Detalle de compra",
                "Moneda",
                f"Periodo {numero}",
                f"Pago {numero}",
            ]
        ].copy()

        temporal = temporal.rename(
            columns={
                f"Periodo {numero}": "Fecha pago",
                f"Pago {numero}": "Monto pago DOP",
            }
        )

        temporal["Número pago"] = numero

        temporal = temporal[
            temporal["Fecha pago"].notna()
            & temporal["Monto pago DOP"].ne(0)
        ]

        lista_pagos.append(temporal)

    if lista_pagos:
        df_pagos = pd.concat(
            lista_pagos,
            ignore_index=True,
        )

    else:
        df_pagos = pd.DataFrame()

    if df_pagos.empty:
        st.info(
            "No existen pagos con fechas válidas "
            "para los filtros seleccionados."
        )

    else:
        df_pagos["Periodo"] = df_pagos[
            "Fecha pago"
        ].dt.strftime("%b-%Y")

        indicadores = st.columns(3)

        indicadores[0].metric(
            "Flujo planificado",
            f"RD${df_pagos['Monto pago DOP'].sum():,.0f}",
        )

        indicadores[1].metric(
            "Meses con pagos",
            f"{df_pagos['Periodo'].nunique():,.0f}",
        )

        indicadores[2].metric(
            "Tiendas con pagos",
            f"{df_pagos['Tienda'].nunique():,.0f}",
        )

        flujo_mensual = (
            df_pagos
            .groupby(
                pd.Grouper(
                    key="Fecha pago",
                    freq="MS",
                ),
                as_index=False,
            )
            .agg(
                Monto_pago_DOP=(
                    "Monto pago DOP",
                    "sum",
                )
            )
        )

        flujo_mensual["Periodo"] = flujo_mensual[
            "Fecha pago"
        ].dt.strftime("%b-%Y")

        grafico_flujo = px.bar(
            flujo_mensual,
            x="Periodo",
            y="Monto_pago_DOP",
            text_auto=".2s",
            color="Monto_pago_DOP",
            color_continuous_scale="Greens",
            title="Flujo de pagos por mes",
        )

        grafico_flujo.update_layout(
            coloraxis_showscale=False,
            height=450,
        )

        st.plotly_chart(
            grafico_flujo,
            width="stretch",
        )

        flujo_tienda = (
            df_pagos
            .groupby(
                "Tienda",
                as_index=False,
            )
            .agg(
                Monto_pago_DOP=(
                    "Monto pago DOP",
                    "sum",
                )
            )
            .sort_values(
                "Monto_pago_DOP"
            )
        )

        grafico_flujo_tienda = px.bar(
            flujo_tienda,
            x="Monto_pago_DOP",
            y="Tienda",
            orientation="h",
            color="Monto_pago_DOP",
            color_continuous_scale="Greens",
            title="Flujo planificado por tienda",
        )

        grafico_flujo_tienda.update_layout(
            coloraxis_showscale=False,
            height=450,
        )

        st.plotly_chart(
            grafico_flujo_tienda,
            width="stretch",
        )

        st.subheader(
            "Detalle de pagos"
        )

        mostrar_tabla(
            df_pagos.sort_values(
                [
                    "Fecha pago",
                    "Tienda",
                ]
            )
        )

    control_flujo = df_filtrado[
        df_filtrado["Problema flujo"]
    ][
        [
            "Tienda",
            "Centro",
            "Agrupador",
            "Supranúmero",
            "Detalle de compra",
            "Presupuesto DOP",
            "Total flujo utilizado",
            "Diferencia flujo DOP",
        ]
    ].copy()

    control_flujo["Diferencia absoluta"] = control_flujo[
        "Diferencia flujo DOP"
    ].abs()

    control_flujo = control_flujo.sort_values(
        "Diferencia absoluta",
        ascending=False,
    )

    st.subheader(
        "Control presupuesto versus flujo"
    )

    indicadores_control = st.columns(3)

    indicadores_control[0].metric(
        "Registros con diferencias",
        f"{len(control_flujo):,.0f}",
    )

    indicadores_control[1].metric(
        "Diferencia absoluta",
        f"RD${control_flujo['Diferencia absoluta'].sum():,.0f}",
    )

    indicadores_control[2].metric(
        "Presupuesto afectado",
        f"RD${control_flujo['Presupuesto DOP'].sum():,.0f}",
    )

    mostrar_tabla(
        control_flujo
    )


# =========================================================
# ENTREGAS Y ALMACÉN
# =========================================================

with tab_entregas:
    st.subheader(
        "Entregas y necesidad de almacenamiento"
    )

    entregas_validas = df_filtrado[
        df_filtrado["Fecha entrega"].notna()
    ].copy()

    entregas_riesgo = df_filtrado[
        df_filtrado[
            "Entrega posterior a apertura"
        ]
    ]

    sin_entrega = df_filtrado[
        df_filtrado[
            "Error fecha entrega"
        ]
    ]

    requiere_almacen = df_filtrado[
        df_filtrado[
            "Almacenaje"
        ].ne("Sin información")
        | df_filtrado[
            "Ubicación almacén"
        ].ne("Sin información")
    ]

    indicadores = st.columns(4)

    indicadores[0].metric(
        "Entregas planificadas",
        f"{len(entregas_validas):,.0f}",
    )

    indicadores[1].metric(
        "Posteriores a apertura",
        f"{len(entregas_riesgo):,.0f}",
    )

    indicadores[2].metric(
        "Sin fecha de entrega",
        f"{len(sin_entrega):,.0f}",
    )

    indicadores[3].metric(
        "Con almacenaje",
        f"{len(requiere_almacen):,.0f}",
    )

    if not entregas_validas.empty:
        entregas_mes = (
            entregas_validas
            .groupby(
                [
                    "Mes número entrega",
                    "Mes entrega",
                ],
                as_index=False,
            )
            .agg(
                Monto_USD=(
                    "Presupuesto USD",
                    "sum",
                )
            )
        )

        calendario_entrega = pd.DataFrame(
            {
                "Mes número entrega": range(1, 13),
                "Mes entrega": ORDEN_MESES,
            }
        )

        entregas_mes = calendario_entrega.merge(
            entregas_mes,
            on=[
                "Mes número entrega",
                "Mes entrega",
            ],
            how="left",
        ).fillna(0)

        grafico_entrega = px.bar(
            entregas_mes,
            x="Mes entrega",
            y="Monto_USD",
            text_auto=".2s",
            color="Monto_USD",
            color_continuous_scale="Oranges",
            title="Valor de entregas por mes",
        )

        grafico_entrega.update_layout(
            coloraxis_showscale=False,
            height=430,
        )

        st.plotly_chart(
            grafico_entrega,
            width="stretch",
        )

    izquierda, derecha = st.columns(2)

    resumen_ubicacion = (
        requiere_almacen
        .groupby(
            "Ubicación almacén",
            as_index=False,
        )
        .agg(
            Monto_USD=(
                "Presupuesto USD",
                "sum",
            )
        )
        .sort_values(
            "Monto_USD"
        )
    )

    if not resumen_ubicacion.empty:
        grafico_ubicacion = px.bar(
            resumen_ubicacion,
            x="Monto_USD",
            y="Ubicación almacén",
            orientation="h",
            color="Monto_USD",
            color_continuous_scale="Blues",
            title="Monto por ubicación de almacén",
        )

        grafico_ubicacion.update_layout(
            coloraxis_showscale=False,
            height=450,
        )

        izquierda.plotly_chart(
            grafico_ubicacion,
            width="stretch",
        )

    resumen_tipo = (
        requiere_almacen
        .groupby(
            "Tipo almacén",
            as_index=False,
        )
        .agg(
            Monto_USD=(
                "Presupuesto USD",
                "sum",
            )
        )
    )

    if not resumen_tipo.empty:
        grafico_tipo = px.pie(
            resumen_tipo,
            names="Tipo almacén",
            values="Monto_USD",
            hole=0.5,
            title="Distribución por tipo de almacén",
        )

        grafico_tipo.update_layout(
            height=450
        )

        derecha.plotly_chart(
            grafico_tipo,
            width="stretch",
        )

    tabla_riesgo = df_filtrado[
        df_filtrado[
            "Entrega posterior a apertura"
        ]
        | df_filtrado[
            "Entrega anterior a compra"
        ]
    ][
        [
            "Tienda",
            "Centro",
            "Agrupador",
            "Supranúmero",
            "Detalle de compra",
            "Presupuesto USD",
            "Fecha compra",
            "Fecha entrega",
            "Fecha apertura",
            "Días entrega a apertura",
            "Ubicación almacén",
            "Tipo almacén",
        ]
    ].sort_values(
        "Días entrega a apertura"
    )

    st.subheader(
        "Entregas con riesgo frente a apertura"
    )

    mostrar_tabla(
        tabla_riesgo
    )


# =========================================================
# CALIDAD DE DATOS
# =========================================================

with tab_calidad:
    st.subheader(
        "Control de calidad de la información"
    )

    total_registros = len(
        df_filtrado
    )

    registros_correctos = len(
        df_filtrado[
            df_filtrado[
                "Calidad registro"
            ].eq("Correcto")
        ]
    )

    registros_revisar = len(
        df_filtrado[
            df_filtrado[
                "Calidad registro"
            ].eq("Revisar")
        ]
    )

    porcentaje_correcto = (
        registros_correctos
        / total_registros
        * 100
        if total_registros > 0
        else 0
    )

    indicadores = st.columns(4)

    indicadores[0].metric(
        "Registros analizados",
        f"{total_registros:,.0f}",
    )

    indicadores[1].metric(
        "Registros correctos",
        f"{registros_correctos:,.0f}",
    )

    indicadores[2].metric(
        "Registros para revisar",
        f"{registros_revisar:,.0f}",
    )

    indicadores[3].metric(
        "Calidad general",
        f"{porcentaje_correcto:,.1f}%",
    )

    resumen_errores = pd.DataFrame(
        {
            "Validación": [
                "Fecha de compra pendiente",
                "Fecha de entrega pendiente",
                "Monto en cero",
                "Cantidad en cero",
                "Supranúmero pendiente",
                "Agrupador pendiente",
                "Lead time pendiente",
                "Compra posterior a apertura",
                "Entrega posterior a apertura",
                "Entrega anterior a compra",
                "Diferencia presupuesto vs. flujo",
            ],
            "Registros": [
                int(
                    df_filtrado[
                        "Error fecha compra"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Error fecha entrega"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Error monto cero"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Error cantidad cero"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Error supranúmero"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Error agrupador"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Error lead time"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Compra posterior a apertura"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Entrega posterior a apertura"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Entrega anterior a compra"
                    ].sum()
                ),
                int(
                    df_filtrado[
                        "Problema flujo"
                    ].sum()
                ),
            ],
        }
    ).sort_values(
        "Registros"
    )

    grafico_calidad = px.bar(
        resumen_errores,
        x="Registros",
        y="Validación",
        orientation="h",
        text="Registros",
        color="Registros",
        color_continuous_scale="OrRd",
        title="Incidencias identificadas",
    )

    grafico_calidad.update_layout(
        coloraxis_showscale=False,
        height=560,
    )

    st.plotly_chart(
        grafico_calidad,
        width="stretch",
    )

    registros_error = df_filtrado[
        df_filtrado[
            "Calidad registro"
        ].eq("Revisar")
    ][
        [
            "Tienda",
            "Centro",
            "Agrupador",
            "Supranúmero",
            "Detalle de compra",
            "Cantidad",
            "Presupuesto USD",
            "Fecha compra",
            "Fecha entrega",
            "Fecha apertura",
            "Lead time",
            "Total flujo utilizado",
            "Diferencia flujo DOP",
            "Cantidad errores",
        ]
    ].sort_values(
        [
            "Cantidad errores",
            "Presupuesto USD",
        ],
        ascending=[
            False,
            False,
        ],
    )

    st.subheader(
        "Registros que requieren corrección"
    )

    mostrar_tabla(
        registros_error
    )

    st.download_button(
        "Descargar registros para corregir",
        data=crear_excel_descarga(
            registros_error,
            "Datos para corregir",
        ),
        file_name="registros_para_corregir.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# =========================================================
# PIE DE PÁGINA
# =========================================================

st.divider()

st.caption(
    "Fuente: hoja Data del archivo cargado. "
    "Las compras sin fecha válida no se incluyen en el calendario mensual, "
    "pero sí se muestran en el control de calidad."
)