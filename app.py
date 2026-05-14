# app.py

import streamlit as st
import pandas as pd
from datetime import timedelta
import re
import random
import string

# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="GENERADOR KASHIO",
    page_icon="📄",
    layout="wide"
)

EXPIRACION_FIJA = "31/12/2040"

# =========================================================
# LIMPIAR TEXTO
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


# =========================================================
# NORMALIZAR MONEDA
# =========================================================

def normalizar_moneda(moneda):

    if pd.isna(moneda):
        return ""

    moneda = limpiar_texto(moneda)

    if moneda in ["SOLES", "SOL", "PEN"]:
        return "PEN"

    if moneda in ["DOLARES", "DÓLARES", "USD", "US$"]:
        return "USD"

    return moneda


# =========================================================
# EXTRAER NOMBRE DESDE DESCRIPCION
# =========================================================

def extraer_nombre_descripcion(texto):

    if pd.isna(texto):
        return ""

    texto = limpiar_texto(texto)

    patrones = [

        "LIC. DE PLATAFORMA KASHIO RECAUDOS ABR 26 -",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS MAY 26 -",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS JUN 26 -",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS JUL 26 -",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS AGO 26 -",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS SEP 26 -",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS OCT 26 -",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS NOV 26 -",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS DIC 26 -",

        "LICENCIA RECAUDOS ABR 26 -",
        "LICENCIA RECAUDOS MAY 26 -",
        "LICENCIA RECAUDOS JUN 26 -",
        "LICENCIA RECAUDOS JUL 26 -",
        "LICENCIA RECAUDOS AGO 26 -",
        "LICENCIA RECAUDOS SEP 26 -",
        "LICENCIA RECAUDOS OCT 26 -",
        "LICENCIA RECAUDOS NOV 26 -",
        "LICENCIA RECAUDOS DIC 26 -",

        "LICENCIA RECAUDOS",
        "LIC. DE PLATAFORMA KASHIO RECAUDOS",
        "PLATAFORMA KASHIO RECAUDOS",

        "-"
    ]

    for patron in patrones:

        texto = texto.replace(
            limpiar_texto(patron),
            ""
        )

    texto = limpiar_texto(texto)

    return texto


# =========================================================
# DETECTAR COLUMNAS
# =========================================================

def detectar_columna(df, opciones):

    columnas = [c.upper().strip() for c in df.columns]

    for opcion in opciones:

        for col in columnas:

            if opcion.upper() in col:

                return df.columns[columnas.index(col)]

    return None


# =========================================================
# GENERAR IDS UNICOS
# =========================================================

def generar_id():

    caracteres = string.ascii_uppercase + string.digits

    return "KSH" + ''.join(
        random.choices(caracteres, k=10)
    )


# =========================================================
# TITULO
# =========================================================

st.title("📄 GENERADOR KASHIO")

st.markdown(
    "### Generador Automático de Plantillas"
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

meses_cortos = {
    "ENERO": "ENE",
    "FEBRERO": "FEB",
    "MARZO": "MAR",
    "ABRIL": "ABR",
    "MAYO": "MAY",
    "JUNIO": "JUN",
    "JULIO": "JUL",
    "AGOSTO": "AGO",
    "SEPTIEMBRE": "SEP",
    "OCTUBRE": "OCT",
    "NOVIEMBRE": "NOV",
    "DICIEMBRE": "DIC"
}

# NOMBRE COMPLETO
nombre_periodo = f"{mes} {anio}"

# DESCRIPCION CORTA
descripcion_periodo = f"{meses_cortos[mes]} {str(anio)[-2:]}"

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
                "MON",
                "MO",
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
        # DETECTAR COLUMNAS MAESTRA
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
        # MATCH
        # =================================================

        reporte["MATCH"] = reporte[
            col_descripcion
        ].apply(extraer_nombre_descripcion)

        maestro["MATCH"] = maestro[
            col_nombre_conta
        ].apply(extraer_nombre_descripcion)

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

            final[col_nro_cp]
            .astype(str)
            .str.strip()
        )

        # =================================================
        # DESCRIPCION FINAL
        # =================================================

        final["DESCRIPCION_FINAL"] = (

            prefijo_descripcion

            + " "

            + descripcion_periodo

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
        # IDS
        # =================================================

        ids_finales = []

        for i in range(len(final)):

            nuevo_id = generar_id()

            ids_finales.append(nuevo_id)

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

            "ID CLIENTE":
                final[col_id_cliente],

            "EMAIL DEL CLIENTE":
                final[col_correo],

            "MONEDA":
                final[col_moneda].apply(normalizar_moneda),

            "MONTO":
                final[col_monto],

            "VENCIMIENTO":
                final["VENCIMIENTO"],

            "EXPIRACION":
                EXPIRACION_FIJA
        })

        # =================================================
        # LIMPIAR ARCHIVO FINAL
        # =================================================

        plantilla = plantilla.fillna("")

        plantilla = plantilla.replace(
            r'[\n\r\t]',
            ' ',
            regex=True
        )

        # =================================================
        # VALIDAR CLIENTES SIN MATCH
        # =================================================

        errores = plantilla[
            plantilla["ID CLIENTE"] == ""
        ]

        if not errores.empty:

            st.warning(
                f"⚠️ {len(errores)} clientes no tuvieron match"
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
        # MOSTRAR REGISTROS VALIDOS
        # =================================================

        registros_validos = plantilla[
            plantilla["ID CLIENTE"] != ""
        ]

        st.success(
            f"✅ Registros válidos: {len(registros_validos)}"
        )

        # =================================================
        # VISTA PREVIA
        # =================================================

        st.subheader("📋 VISTA PREVIA")

        st.dataframe(
            registros_validos.head(1000),
            use_container_width=True
        )

        st.info(
            f"Mostrando primeras {min(len(registros_validos),1000)} filas de {len(registros_validos)} registros."
        )

        # =================================================
        # EXPORTAR SOLO VALIDOS
        # =================================================

        nombre_salida = (
            f"KASHIO_{mes}_{anio}.xlsx"
        )

        registros_validos.to_excel(
            nombre_salida,
            index=False
        )

        # =================================================
        # DESCARGAR
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
