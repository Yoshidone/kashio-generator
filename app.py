# app.py

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font
import os
import re

# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="Kashio Generator",
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
    
    texto = str(texto).upper().strip()
    texto = re.sub(r"\s+", " ", texto)
    
    return texto


def obtener_mes(fecha):
    meses = {
        1: "ENERO",
        2: "FEBRERO",
        3: "MARZO",
        4: "ABRIL",
        5: "MAYO",
        6: "JUNIO",
        7: "JULIO",
        8: "AGOSTO",
        9: "SEPTIEMBRE",
        10: "OCTUBRE",
        11: "NOVIEMBRE",
        12: "DICIEMBRE"
    }
    
    return meses[fecha.month]


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
        historial = pd.concat([historial, nuevo], ignore_index=True)
    else:
        historial = nuevo
    
    historial.to_excel(HISTORIAL_FILE, index=False)


def generar_id(fecha, correlativo):
    
    mes = obtener_mes(fecha)
    anio = fecha.year
    
    return f"KSH{mes}{anio}{str(correlativo).zfill(5)}"


def detectar_columna(df, opciones):
    
    columnas = [c.upper().strip() for c in df.columns]
    
    for opcion in opciones:
        for col in columnas:
            if opcion.upper() in col:
                return df.columns[columnas.index(col)]
    
    return None


# =========================================================
# TITULO
# =========================================================

st.title("📄 GENERADOR KASHIO")
st.markdown("### Subir Reporte + Base Maestra")

# =========================================================
# UPLOADS
# =========================================================

archivo_reporte = st.file_uploader(
    "📥 Subir Reporte Mensual",
    type=["xlsx"]
)

archivo_maestro = st.file_uploader(
    "📥 Subir Base Maestra Clientes",
    type=["xlsx"]
)

# =========================================================
# PROCESAMIENTO
# =========================================================

if archivo_reporte and archivo_maestro:
    
    try:
        
        reporte = pd.read_excel(archivo_reporte)
        maestro = pd.read_excel(archivo_maestro)
        
        st.success("✅ Archivos cargados correctamente")
        
        # =================================================
        # DETECTAR COLUMNAS
        # =================================================
        
        col_fecha = detectar_columna(
            reporte,
            ["FECHA"]
        )
        
        col_nombre = detectar_columna(
            reporte,
            ["RAZON SOCIAL", "NOMBRE"]
        )
        
        col_descripcion = detectar_columna(
            reporte,
            ["DESCRIPCION"]
        )
        
        col_moneda = detectar_columna(
            reporte,
            ["MONEDA", "MO"]
        )
        
        col_monto = detectar_columna(
            reporte,
            ["PRECIO VEN", "VALOR VEN"]
        )
        
        # =================================================
        # MAESTRO
        # =================================================
        
        col_id_cliente = detectar_columna(
            maestro,
            ["ID CLIENTE"]
        )
        
        col_correo = detectar_columna(
            maestro,
            ["CORREO", "EMAIL"]
        )
        
        col_nombre_conta = detectar_columna(
            maestro,
            ["NOMBRE CONTABILIDAD"]
        )
        
        # =================================================
        # LIMPIEZA
        # =================================================
        
        reporte["MATCH"] = reporte[col_nombre].apply(limpiar_texto)
        maestro["MATCH"] = maestro[col_nombre_conta].apply(limpiar_texto)
        
        # =================================================
        # MERGE
        # =================================================
        
        final = reporte.merge(
            maestro,
            on="MATCH",
            how="left"
        )
        
        # =================================================
        # FECHAS
        # =================================================
        
        final[col_fecha] = pd.to_datetime(final[col_fecha])
        
        # =================================================
        # IDS
        # =================================================
        
        ultimo = obtener_ultimo_correlativo()
        
        ids_historial = []
        ids_finales = []
        
        for i in range(len(final)):
            
            correlativo = ultimo + i + 1
            
            fecha = final.iloc[i][col_fecha]
            
            nuevo_id = generar_id(fecha, correlativo)
            
            ids_finales.append(nuevo_id)
            
            ids_historial.append({
                "ID": nuevo_id,
                "CORRELATIVO": correlativo
            })
        
        # =================================================
        # VENCIMIENTO
        # =================================================
        
        final["VENCIMIENTO"] = (
            final[col_fecha] + timedelta(days=30)
        ).dt.strftime("%d/%m/%Y")
        
        # =================================================
        # PLANTILLA FINAL
        # =================================================
        
        plantilla = pd.DataFrame({
            "ID ORDEN DE PAGO": ids_finales,
            "REFERENCIA": ids_finales,
            "NOMBRE": final[col_nombre],
            "DESCRIPCION": final[col_descripcion],
            "ID CLIENTE (*)": final[col_id_cliente],
            "EMAIL DEL CLIENTE (*)": final[col_correo],
            "MONEDA": final[col_moneda],
            "MONTO": final[col_monto],
            "VENCIMIENTO": final["VENCIMIENTO"],
            "EXPIRACION": EXPIRACION_FIJA
        })
        
        # =================================================
        # VALIDACIONES
        # =================================================
        
        errores = plantilla[
            plantilla["ID CLIENTE (*)"].isna()
        ]
        
        if not errores.empty:
            
            st.warning("⚠️ Algunos clientes no tuvieron match")
            
            st.dataframe(
                errores[[
                    "NOMBRE"
                ]]
            )
        
        # =================================================
        # PREVIEW
        # =================================================
        
        st.subheader("📋 Vista Previa")
        st.dataframe(plantilla)
        
        # =================================================
        # GUARDAR HISTORIAL
        # =================================================
        
        guardar_historial(ids_historial)
        
        # =================================================
        # EXPORTAR
        # =================================================
        
        nombre_salida = (
            f"KASHIO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        plantilla.to_excel(
            nombre_salida,
            index=False
        )
        
        # =================================================
        # DOWNLOAD
        # =================================================
        
        with open(nombre_salida, "rb") as file:
            
            st.download_button(
                label="⬇️ DESCARGAR PLANTILLA KASHIO",
                data=file,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    except Exception as e:
        
        st.error(f"❌ Error: {str(e)}")
