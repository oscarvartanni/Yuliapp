import streamlit as st
from docx import Document
import io
import re
import os

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="CRM Document Cloud", layout="wide")

# Mapeo exacto de tus archivos fuente
TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

def extraer_datos(file):
    doc = Document(file)
    full_text = "\n".join([p.text for p in doc.paragraphs])
    return {
        "Fecha": re.search(r"Fecha:\s*(.*)", full_text, re.I).group(1) if re.search(r"Fecha:\s*(.*)", full_text, re.I) else "",
        "Objetivo": re.search(r"Objetivo:\s*(.*)", full_text, re.I).group(1) if re.search(r"Objetivo:\s*(.*)", full_text, re.I) else ""
    }

st.title("🚀 Generador CRM Online")

opcion = st.selectbox("Selecciona la plantilla:", list(TEMPLATES.keys()))
archivo_info = st.file_uploader("Sube el Word con la nueva información", type=["docx"])

if archivo_info:
    datos_sugeridos = extraer_datos(archivo_info)
    
    with st.form("editor_form"):
        col1, col2 = st.columns(2)
        with col1:
            f_val = st.text_input("Fecha:", value=datos_sugeridos["Fecha"])
            o_val = st.text_area("Objetivo:", value=datos_sugeridos["Objetivo"])
        with col2:
            st.info("Nota: Los datos se insertarán respetando el formato original del template.")
            omitir = st.checkbox("Omitir campos vacíos")
            
        generar = st.form_submit_button("Generar Word Final")

    if generar:
        if os.path.exists(TEMPLATES[opcion]):
            doc = Document(TEMPLATES[opcion])
            
            # Reemplazo básico en párrafos
            for p in doc.paragraphs:
                if "Fecha:" in p.text: p.text = f"Fecha: {f_val}"
                if "Objetivo:" in p.text: p.text = f"Objetivo: {o_val}"
            
            # Guardar y Descargar
            target = io.BytesIO()
            doc.save(target)
            target.seek(0)
            
            st.success("¡Documento procesado!")
            st.download_button(
                label="⬇️ Descargar Archivo",
                data=target,
                file_name=f"Generado_{opcion}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error(f"Error: No se encontró el archivo base {TEMPLATES[opcion]} en el servidor.")