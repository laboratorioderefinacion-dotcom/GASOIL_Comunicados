#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO
from datetime import datetime
from mailmerge import MailMerge

st.set_page_config(page_title="GAS OIL Comunicados - Generador de Informes", layout="centered")

# -----------------------------
# Utilidades (misma lógica que tu script)
# -----------------------------
def extraer_valor(df, nombre_celda, n_muestra):
    col = 3 + n_muestra
    fila = df[df[1] == nombre_celda]
    if fila.empty:
        return "#########"
    valor = fila.iloc[0, col]
    if pd.isna(valor) or str(valor).strip() == "":
        return "-----"
    return str(valor)

def extraer_valor_float(df, nombre_celda, n_muestra):
    valor = extraer_valor(df, nombre_celda, n_muestra)
    if valor in ["-----", "#########"]:
        return valor
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
        return "#########"
    valor = fila.iloc[0, 3 + n_muestra]
    if pd.isna(valor) or str(valor).strip() == "":
        return "-----"
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

def construir_numero_comunicado(nombre_base):
    """
    Toma nombre del archivo (sin .csv) y devuelve:
    - "A.C. xxxx" si empieza con AC
    - "C.M. xxxx" si empieza con CM
    - sino devuelve el nombre tal cual
    Soporta separadores: '-', '_', ' ', '.', etc.
    """
    if not nombre_base:
        return "#########"

    nb = nombre_base.strip()

    # Detecta prefijo AC/CM al inicio
    m = re.match(r"^(AC|CM)\b[\s\-_\.]*?(.*)$", nb, flags=re.IGNORECASE)
    if not m:
        return nb

    pref = m.group(1).upper()
    resto = m.group(2).strip()

    # Si resto quedó vacío, igual devolvemos algo
    if pref == "AC":
        return "A.C. " + (resto if resto else "#########")
    if pref == "CM":
        return "C.M. " + (resto if resto else "#########")
    return nb


# -----------------------------
# UI
# -----------------------------
st.title("🛢️ GAS OIL | Generador de informe Asistencia Comercial")

uploaded = st.file_uploader("📄 Cargar CSV de LIMS", type=["csv"])

# Nombre LIMS automático desde el CSV
nombre_archivo_LIMS = os.path.splitext(uploaded.name)[0]
numero_comunicado = construir_numero_comunicado(nombre_archivo_LIMS)

st.caption(f"📌 Archivo cargado: **{uploaded.name}**  |  Comunicado: **{numero_comunicado}**")

st.divider()

# ORDEN: Cliente -> Muestras -> Prioridad
cliente = st.text_input("Nombre del cliente", value="")

n_muestras = st.number_input(
    "¿Cuántas muestras desea procesar? (1 a 6)",
    min_value=1, max_value=6, value=1, step=1
)

prioridad = st.text_input("Prioridad (días hábiles)", value="")

st.divider()

# -----------------------------
# Lectura CSV
# -----------------------------
try:
    df_a = pd.read_csv(uploaded, encoding="latin1", sep=";", header=None)
    st.success("✅ CSV leído correctamente.")
except Exception as e:
    st.error(f"❌ No pude leer el CSV: {e}")
    st.stop()

# -----------------------------
# Plantilla (en la raíz del repo)
# -----------------------------
ruta_plantillas = os.getcwd()
plantilla = os.path.join(ruta_plantillas, f"GASOIL {int(n_muestras)}M.docx")

if not os.path.isfile(plantilla):
    st.error(
        f"❌ No encuentro la plantilla: **{os.path.basename(plantilla)}**\n\n"
        "📌 Asegurate que esté en el repo, en el **mismo nivel** que `app.py`."
    )
    st.stop()

# -----------------------------
# Extracción por muestra
# -----------------------------
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

    # Agua + reglas
    agua_val = extraer_valor(df_a, "Agua por Karl Fisher", i)
    if str(fase_val).strip() != "No se observa":
        datos_muestras[f"agua_m{i}"] = "(**)"
        notas_agua.add(
            "(**) El análisis de Agua por Karl Fischer no se realiza dado que la muestra presenta una fase no miscible con el combustible (presumiblemente agua), lo cual no permite extraer un alícuota representativa."
        )
    elif str(asp_val).strip() in ["1", "2"] and str(agua_val).strip() in ["", "-----", "#########"]:
        datos_muestras[f"agua_m{i}"] = "(*)"
        notas_agua.add(
            "(*) Dado el resultado de Aspecto por la norma ASTM D 4176 y la ausencia de fase no miscible visible en la muestra, se puede considerar que la misma se encuentra en especificación de contenido de Agua."
        )
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

# Notas agua ordenadas
notas_ordenadas = []
for prefijo in ["(*)", "(**)"]:
    for nota in sorted(notas_agua):
        if nota.startswith(prefijo):
            notas_ordenadas.append(nota)

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

if agua_reportada:
    norma_agua = extraer_valor_norma(df_a, "Agua por Karl Fisher")
    datos_fusion["n_agua"] = str(norma_agua).replace("_", " ") if norma_agua else "#########"
else:
    datos_fusion["n_agua"] = "ASTM D 6304"

datos_fusion.update(datos_muestras)

# -----------------------------
# Generación Word
# -----------------------------
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
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        if obs.strip():
            with st.expander("📌 Observaciones detectadas (Comentarios por muestra)"):
                st.text(obs)

        if datos_fusion.get("nota_agua", "").strip():
            with st.expander("💧 Notas de Agua"):
                st.text(datos_fusion["nota_agua"])

    except Exception as e:
        st.error(f"❌ Error generando el Word: {e}")

