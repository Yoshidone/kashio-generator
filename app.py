# app.py

import streamlit as st
import pandas as pd
from datetime import timedelta
import os
import re

# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="GENERADOR KASHIO",
    page_icon="📄",
    layout="wide"
)

HISTORIAL_FILE = "historial_ids.xlsx"
EXPIRACION_FIJA = "31/12/2040"

# =========================================================
# FUNCIONES
# =========================================================

def limpiar_texto(texto):

    if pd.isna(texto):
        return ""

    texto = str(texto)

    texto = texto.upper()

    texto = texto.strip()

    texto = texto.replace("\n", " ")

    texto = texto.replace("\r", " ")

    texto = texto.replace("\t", " ")

    texto = re.sub(r"\s+", " ", texto)

    texto = texto.encode("ascii", "ignore").decode()

    return texto


def extraer_nombre_descripcion(texto):

    if pd.isna(texto):
        return ""

    texto = str(texto).upper()

    patrones = [
        "LICENCIA RECAUDOS",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS",
        "LIC DE PLATAFORMA KASHIO RECAUDOS",
        "PLATAFORMA KASHIO RECAUDOS",
        "RECAUDOS ABR 26 -",
        "RECAUDOS MAY 26 -",
        "RECAUDOS JUN 26 -",
        "RECAUDOS JUL 26 -",
        "RECAUDOS AGO 26 -",
        "RECAUDOS SEP 26 -",
        "RECAUDOS OCT 26 -",
        "RECAUDOS NOV 26 -",
        "RECAUDOS DIC 26 -",
        "RECAUDOS ENE 26 -",
        "RECAUDOS FEB 26 -",
        "RECAUDOS MAR 26 -",
        "-",
    ]

    for patron in patrones:
        texto = texto.replace(patron, "")

    texto = texto.strip()

    return limpiar_texto(texto)


def detectar_columna(df, opciones):

    columnas = [c.upper().strip() for c in df.columns]

    for opcion in opciones:

        for col in columnas:

            if opcion.upper() in col:

                return df.columns[columnas.index(col)]

    return None


def obtener_ultimo_correlativo():

    if not os.path.exists(HISTORIAL_FILE):
        return 0

    try:

        historial = pd.read_excel(HISTORIAL_FILE)

        if historial.empty:
            return 0

        return historial["CORRELATIVO"].max()

    except:
        return 0


def guardar_historial(ids_generados):

    nuevo = pd.DataFrame(ids_generados)

    if os.path.exists(HISTORIAL_FILE):

        historial = pd.read_excel(HISTORIAL_FILE)

        historial = pd.concat(
            [historial, nuevo],
            ignore_index=True
        )

    else:

        historial = nuevo

    historial.to_excel(
        HISTORIAL_FILE,
        index=False
    )


def generar_id(mes, anio, correlativo):

    return f"KSH{mes}{anio}{str(correlativo).zfill(5)}"


# =========================================================
# TITULO
# =========================================================

st.title("📄 GENERADOR KASHIO")

st.markdown(
    "### Generador Automático de Plantillas Kashio"
)

# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.header("⚙️ CONFIGURACION")

mes = st.sidebar.selectbox(
    "MES",
    [
        "ENERO",
        "FEBRERO",
        "MARZO",
        "ABRIL",
        "MAYO",
        "JUNIO",
        "JULIO",
        "AGOSTO",
        "SEPTIEMBRE",
        "OCTUBRE",
        "NOVIEMBRE",
        "DICIEMBRE"
    ]
)

anio = st.sidebar.selectbox(
    "AÑO",
    [2025, 2026, 2027, 2028, 2029, 2030]
)

nombre_periodo = f"{mes} {anio}"

prefijo_descripcion = st.sidebar.text_input(
    "PREFIJO DESCRIPCION",
    value="LICENCIA RECAUDOS"
)

dias_vencimiento = st.sidebar.number_input(
    "DIAS VENCIMIENTO",
    min_value=1,
    max_value=365,
    value=30
)

# =========================================================
# SUBIR ARCHIVOS
# =========================================================

st.subheader("📂 SUBIR ARCHIVOS")

archivo_maestro = st.file_uploader(
    "📥 SUBIR BASE MAESTRA CLIENTES",
    type=["xlsx"]
)

archivo_reporte = st.file_uploader(
    "📥 SUBIR REPORTE MENSUAL",
    type=["xlsx"]
)

# =========================================================
# PROCESAMIENTO
# =========================================================

if archivo_maestro and archivo_reporte:

    try:

        maestro = pd.read_excel(archivo_maestro)
        reporte = pd.read_excel(archivo_reporte)

        st.success("✅ Archivos cargados correctamente")

        # =================================================
        # DETECTAR COLUMNAS REPORTE
        # =================================================

        col_fecha = detectar_columna(
            reporte,
            ["FECHA"]
        )

        col_tipo_cpe = detectar_columna(
            reporte,
            ["TIPO CPE"]
        )

        col_nro_cp = detectar_columna(
            reporte,
            [
                "N° COMPROBANTE",
                "NRO COMPROBANTE",
                "COMPROBANTE"
            ]
        )

        col_descripcion = detectar_columna(
            reporte,
            [
                "DESCRIPCION",
                "DESCRIPCIÓN",
                "PRODUCTOS/SERVICIOS"
            ]
        )

        col_moneda = detectar_columna(
            reporte,
            [
                "MONEDA",
                "MO",
                "MON",
                "DIVISA",
                "CURRENCY"
            ]
        )

        col_monto = detectar_columna(
            reporte,
            [
                "PRECIO VEN",
                "PRECIO VENTA",
                "VALOR VEN",
                "VALOR VENTA",
                "IMPORTE"
            ]
        )

        # =================================================
        # DETECTAR COLUMNAS BASE MAESTRA
        # =================================================

        col_id_cliente = detectar_columna(
            maestro,
            ["ID CLIENTE"]
        )

        col_correo = detectar_columna(
            maestro,
            [
                "CORREO",
                "EMAIL"
            ]
        )

        col_nombre_conta = detectar_columna(
            maestro,
            [
                "NOMBRE CONTABILIDAD"
            ]
        )

        # =================================================
        # VALIDACIONES
        # =================================================

        columnas_faltantes = []

        if not col_fecha:
            columnas_faltantes.append("FECHA")

        if not col_tipo_cpe:
            columnas_faltantes.append("TIPO CPE")

        if not col_nro_cp:
            columnas_faltantes.append("NRO COMPROBANTE")

        if not col_descripcion:
            columnas_faltantes.append("DESCRIPCION")

        if not col_moneda:
            columnas_faltantes.append("MONEDA")

        if not col_monto:
            columnas_faltantes.append("MONTO")

        if not col_id_cliente:
            columnas_faltantes.append("ID CLIENTE")

        if not col_correo:
            columnas_faltantes.append("CORREO")

        if not col_nombre_conta:
            columnas_faltantes.append("NOMBRE CONTABILIDAD")

        if columnas_faltantes:

            st.error(
                f"❌ Faltan columnas: {', '.join(columnas_faltantes)}"
            )

            st.stop()

        # =================================================
        # MATCH INTELIGENTE
        # =================================================

        reporte["MATCH"] = reporte[
            col_descripcion
        ].apply(extraer_nombre_descripcion)

        maestro["MATCH"] = maestro[
            col_nombre_conta
        ].apply(limpiar_texto)

        # =================================================
        # MERGE
        # =================================================

        final = reporte.merge(
            maestro,
            on="MATCH",
            how="left"
        )

        # =================================================
        # REFERENCIA
        # =================================================

        final["REFERENCIA_FINAL"] = (

            final[col_tipo_cpe]
            .astype(str)
            .str.upper()
            .str.strip()

            + " "

            + final[col_nro_cp]
            .astype(str)
            .str.strip()
        )

        # =================================================
        # DESCRIPCION FINAL
        # =================================================

        final["DESCRIPCION_FINAL"] = (

            prefijo_descripcion

            + " "

            + nombre_periodo

            + " "

            + final["REFERENCIA_FINAL"]
        )

        # =================================================
        # FECHAS
        # =================================================

        final[col_fecha] = pd.to_datetime(
            final[col_fecha]
        )

        final["VENCIMIENTO"] = (
            final[col_fecha]
            + timedelta(days=dias_vencimiento)
        ).dt.strftime("%d/%m/%Y")

        # =================================================
        # IDS UNICOS
        # =================================================

        ultimo = obtener_ultimo_correlativo()

        ids_finales = []
        historial_ids = []

        for i in range(len(final)):

            correlativo = ultimo + i + 1

            nuevo_id = generar_id(
                mes,
                anio,
                correlativo
            )

            ids_finales.append(nuevo_id)

            historial_ids.append({
                "ID": nuevo_id,
                "CORRELATIVO": correlativo
            })

        # =================================================
        # PLANTILLA FINAL
        # =================================================

        plantilla = pd.DataFrame({

            "ID ORDEN DE PAGO":
                ids_finales,

            "REFERENCIA":
                final["REFERENCIA_FINAL"],

            "NOMBRE":
                nombre_periodo,

            "DESCRIPCION":
                final["DESCRIPCION_FINAL"],

            "ID CLIENTE (*)":
                final[col_id_cliente],

            "EMAIL DEL CLIENTE (*)":
                final[col_correo],

            "MONEDA":
                final[col_moneda],

            "MONTO":
                final[col_monto],

            "VENCIMIENTO":
                final["VENCIMIENTO"],

            "EXPIRACION":
                EXPIRACION_FIJA
        })

        # =================================================
        # CLIENTES SIN MATCH
        # =================================================

        errores = plantilla[
            plantilla["ID CLIENTE (*)"].isna()
        ]

        if not errores.empty:

            st.warning(
                "⚠️ Algunos clientes no tuvieron match"
            )

            st.dataframe(
                errores[
                    [
                        "REFERENCIA",
                        "DESCRIPCION"
                    ]
                ],
                use_container_width=True
            )

        # =================================================
        # VISTA PREVIA
        # =================================================

        st.subheader("📋 VISTA PREVIA")

        st.dataframe(
            plantilla,
            use_container_width=True
        )

        # =================================================
        # GUARDAR HISTORIAL
        # =================================================

        guardar_historial(
            historial_ids
        )

        # =================================================
        # EXPORTAR
        # =================================================

        nombre_salida = (
            f"KASHIO_{mes}_{anio}.xlsx"
        )

        plantilla.to_excel(
            nombre_salida,
            index=False
        )

        # =================================================
        # DESCARGA
        # =================================================

        with open(nombre_salida, "rb") as file:

            st.download_button(
                label="⬇️ DESCARGAR PLANTILLA KASHIO",
                data=file,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:

        st.error(
            f"❌ ERROR: {str(e)}"
        )
