#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime
from mailmerge import MailMerge

st.set_page_config(page_title="GAS OIL - Generador de Informes", layout="centered")

# -----------------------------
# Utilidades (idénticas a tu lógica)
# -----------------------------
def extraer_valor(df, nombre_celda, n_muestra):
    col = 3 + n_muestra
    fila = df[df[1] == nombre_celda]
    if fila.empty:
        return "#########"  # El análisis no existe
    valor = fila.iloc[0, col]
    if pd.isna(valor) or str(valor).strip() == "":
        return "-----"  # El análisis existe pero no tiene dato
    return str(valor)

def extraer_valor_float(df, nombre_celda, n_muestra):
    valor = extraer_valor(df, nombre_celda, n_muestra)
    if valor in ["-----", "#########"]:
        return valor
    # Mantener comparadores si vinieran como texto
    if isinstance(valor, str) and (valor.startswith(">") or valor.startswith("<") or valor.startswith("&gt;") or valor.startswith("&lt;")):
        return valor
    try:
        return float(str(valor).replace(",", "."))
    except ValueError:
        return "#########"

def formatear(valor, decimales):
    if valor in ["#########", "-----"]:
        return valor
    if isinstance(valor, float):
        return f"{valor:.{decimales}f}".replace(".", ",")
    return str(valor)

def extraer_datos(df, nombre_celda, n_muestra):
    fila = df[df[0] == nombre_celda]
    if fila.empty:
        return "#########"  # El análisis no existe
    valor = fila.iloc[0, 3 + n_muestra]
    if pd.isna(valor) or str(valor).strip() == "":
        return "-----"  # El análisis existe pero no tiene dato
    return str(valor)

def formatear_fecha_datos(fecha_str, incluir_hora=True):
    if fecha_str in ["-----", "#########"]:
        return fecha_str
    try:
        dt = datetime.strptime(fecha_str, "%d/%b/%Y %H:%M:%S")
        return dt.strftime("%d/%m/%Y %H:%M") if incluir_hora else dt.strftime("%d/%m/%Y")
    except ValueError:
        return "#########"

def extraer_valor_norma(df, nombre):
    resultado = df.loc[df[1] == nombre, 2]
    return resultado.values[0] if not resultado.empty else ""


# -----------------------------
# UI
# -----------------------------
st.title("🛢️ GAS OIL | Generador de informe Word (Streamlit)")

st.markdown(
    """
Subí el **CSV de LIMS**, completá **cliente**, **prioridad** y **cantidad de muestras**.  
La app generará el **Word** usando tu plantilla `GASOIL {n}M.docx`.
"""
)

uploaded = st.file_uploader("📄 Cargar CSV de LIMS", type=["csv"])

col1, col2 = st.columns(2)
with col1:
    n_muestras = st.number_input("¿Cuántas muestras desea procesar? (1 a 6)", min_value=1, max_value=6, value=1, step=1)
with col2:
    prioridad = st.text_input("Prioridad (días hábiles)", value="")

cliente = st.text_input("Nombre del cliente", value="")

# Nombre archivo (para A.C / C.M)
nombre_archivo_LIMS = st.text_input(
    "Nombre del archivo LIMS (solo para el número de comunicado, ej: AC-1234 o CM-5678)",
    value=""
)

st.divider()

# -----------------------------
# Procesamiento
# -----------------------------
if uploaded is not None:
    try:
        df_a = pd.read_csv(uploaded, encoding="latin1", sep=";", header=None)
        st.success("✅ CSV leído correctamente.")
    except Exception as e:
        st.error(f"❌ No pude leer el CSV: {e}")
        st.stop()

    # Determinar número de comunicado desde el "nombre_archivo_LIMS"
    numero_comunicado = ""
    if len(nombre_archivo_LIMS) >= 2:
        pref = nombre_archivo_LIMS[:2].lower()
        # Mantengo tu lógica de [3:] (asumiendo formato tipo 'AC-1234')
        if pref == "ac":
            numero_comunicado = "A.C. " + str(nombre_archivo_LIMS[3:])
        elif pref == "cm":
            numero_comunicado = "C.M. " + str(nombre_archivo_LIMS[3:])
        else:
            numero_comunicado = nombre_archivo_LIMS  # fallback
    else:
        numero_comunicado = "#########"

    # Selección de plantilla desde repo
    ruta_plantillas = os.path.join(os.getcwd(), "Formularios")
    plantilla = os.path.join(ruta_plantillas, f"GASOIL {int(n_muestras)}M.docx")

    if not os.path.isfile(plantilla):
        st.error(f"❌ No encuentro la plantilla: {plantilla}\n\nChequeá que exista en `Formularios/`.")
        st.stop()

    # Extracción por muestra
    datos_muestras = {}
    obs = ""
    notas_agua = set()
    agua_reportada = False

    for i in range(1, int(n_muestras) + 1):
        datos_muestras[f"lims_m{i}"] = extraer_datos(df_a, "Número de Muestra", i)
        datos_muestras["motivo"] = extraer_datos(df_a, "Observaciones", 1)
        datos_muestras[f"fecha_ext_m{i}"] = formatear_fecha_datos(extraer_datos(df_a, "Fecha de Extracción", i), incluir_hora=False)
        datos_muestras[f"lug_ext_m{i}"] = extraer_datos(df_a, "Lugar de Extracción", i)
        datos_muestras[f"vol_m{i}"] = extraer_datos(df_a, "Volumen de Muestra", i).replace(".", ",")

        datos_muestras[f"dens_15_m{i}"] = formatear(extraer_valor_float(df_a, "Densidad a 15ºC", i), 4)
        datos_muestras[f"dens_20_m{i}"] = formatear(extraer_valor_float(df_a, "Densidad a 20ºC", i), 4)

        # Aspecto GO y Fase no miscible
        asp_val = extraer_valor(df_a, "Aspecto GO", i)
        datos_muestras[f"asp_GO_m{i}"] = asp_val

        fase_val = extraer_valor(df_a, "Fase No Miscible", i)
        datos_muestras[f"fase_m{i}"] = fase_val

        datos_muestras[f"temp_m{i}"] = formatear(extraer_valor_float(df_a, "Temperatura", i), 1)
        datos_muestras[f"color_vis_m{i}"] = extraer_valor(df_a, "Color", i)
        datos_muestras[f"particulas_m{i}"] = extraer_valor(df_a, "Partículas", i)
        datos_muestras[f"aspecto_m{i}"] = extraer_valor(df_a, "Aspecto", i)
        datos_muestras[f"color_m{i}"] = extraer_valor(df_a, "Color ASTM 1500", i)
        datos_muestras[f"PM_m{i}"] = formatear(extraer_valor_float(df_a, "Punto de Inflamación Pensky Martens", i), 1)

        # Agua Karl Fisher + reglas
        agua_val = extraer_valor(df_a, "Agua por Karl Fisher", i)

        if str(fase_val).strip() != "No se observa":
            datos_muestras[f"agua_m{i}"] = "(**)"
            notas_agua.add("(**) El análisis de Agua por Karl Fischer no se realiza dado que la muestra presenta una fase no miscible con el combustible (presumiblemente agua), lo cual no permite extraer un alícuota representativa.")
        elif str(asp_val).strip() in ["1", "2"] and str(agua_val).strip() in ["", "-----", "#########"]:
            datos_muestras[f"agua_m{i}"] = "(*)"
            notas_agua.add("(*) Dado el resultado de Aspecto por la norma ASTM D 4176 y la ausencia de fase no miscible visible en la muestra, se puede considerar que la misma se encuentra en especificación de contenido de Agua.")
        else:
            datos_muestras[f"agua_m{i}"] = agua_val
            if str(agua_val).strip() not in ["", "-----", "#########", "(*)", "(**)"]:
                agua_reportada = True

        datos_muestras[f"azufre_m{i}"] = formatear(extraer_valor_float(df_a, "Azufre", i), 1)
        datos_muestras[f"biodiesel_m{i}"] = formatear(extraer_valor_float(df_a, "Contenido de Biodiesel", i), 1)

        comentario = str(extraer_valor(df_a, "Comentarios", i))
        datos_muestras[f"comentario_m{i}"] = comentario

        if comentario and comentario.strip() not in ["", "-----", "#########"]:
            obs += f"Muestra con LIMS {datos_muestras[f'lims_m{i}']}: {comentario}\n"
            datos_muestras[f"m{i}"] = f"(#{i})"
        else:
            datos_muestras[f"m{i}"] = ""

    datos_muestras["fecha_rec"] = formatear_fecha_datos(extraer_datos(df_a, "Fecha de Recepción", 1), incluir_hora=True)

    producto = extraer_datos(df_a, "Producto", 1)
    esp_azufre = "#########"
    nombre_archivo_nuevo = f"{numero_comunicado} Gas Oil - {cliente}".strip()

    if producto == "GAS_OIL_50S":
        esp_azufre = "50"
        nombre_archivo_nuevo = f"{numero_comunicado} Gas Oil 50S - {cliente}".strip()
    elif producto == "GAS_OIL_10S":
        esp_azufre = "10"
        nombre_archivo_nuevo = f"{numero_comunicado} Gas Oil 10S - {cliente}".strip()

    # Notas de agua ordenadas
    notas_ordenadas = []
    for prefijo in ["(*)", "(**)"]:
        for nota in sorted(notas_agua):
            if nota.startswith(prefijo):
                notas_ordenadas.append(nota)

    # Datos de fusión (merge)
    datos_fusion = {
        "num_com": numero_comunicado,
        "esp_azufre": esp_azufre,
        "cliente": cliente,
        "prioridad": prioridad,
        "n_dens": str(extraer_valor_norma(df_a, "Densidad a 20ºC")).replace("_", " ") if extraer_valor_norma(df_a, "Densidad a 20ºC") else "#########",
        "n_asp_GO": str(extraer_valor_norma(df_a, "Aspecto GO")).replace("_", " ") if extraer_valor_norma(df_a, "Aspecto GO") else "#########",
        "n_aspecto": str(extraer_valor_norma(df_a, "Aspecto")).replace("_", " ") if extraer_valor_norma(df_a, "Aspecto") else "#########",
        "n_color": str(extraer_valor_norma(df_a, "Color ASTM 1500")).replace("_", " ") if extraer_valor_norma(df_a, "Color ASTM 1500") else "#########",
        "n_PM": str(extraer_valor_norma(df_a, "Punto de Inflamación Pensky Martens")).replace("_", " ") if extraer_valor_norma(df_a, "Punto de Inflamación Pensky Martens") else "#########",
        "n_azufre": str(extraer_valor_norma(df_a, "Azufre")).replace("_", " ") if extraer_valor_norma(df_a, "Azufre") else "#########",
        "n_bio": str(extraer_valor_norma(df_a, "Contenido de Biodiesel")).replace("_", " ") if extraer_valor_norma(df_a, "Contenido de Biodiesel") else "#########",
        "obs": obs,
        "nota_agua": "\n".join(notas_ordenadas) if notas_ordenadas else ""
    }

    # Norma agua según si hubo datos válidos o no
    if agua_reportada:
        norma_agua = extraer_valor_norma(df_a, "Agua por Karl Fisher")
        datos_fusion["n_agua"] = str(norma_agua).replace("_", " ") if norma_agua else "#########"
    else:
        datos_fusion["n_agua"] = "ASTM D 6304"

    datos_fusion.update(datos_muestras)

    # Botón para generar
    if st.button("📝 Generar informe Word", type="primary"):
        try:
            doc = MailMerge(plantilla)
            doc.merge(**{k: str(v) for k, v in datos_fusion.items()})
            buffer = BytesIO()
            doc.write(buffer)
            buffer.seek(0)

            st.success("✅ Informe generado.")
            st.download_button(
                "⬇️ Descargar informe .docx",
                data=buffer,
                file_name=f"{nombre_archivo_nuevo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            # Opcional: mostrar observaciones en pantalla
            if obs.strip():
                with st.expander("📌 Observaciones detectadas (Comentarios por muestra)"):
                    st.text(obs)

            if datos_fusion.get("nota_agua", "").strip():
                with st.expander("💧 Notas de Agua"):
                    st.text(datos_fusion["nota_agua"])

        except Exception as e:
            st.error(f"❌ Error generando el Word: {e}")
else:
    st.info("📥 Subí un CSV para comenzar.")

