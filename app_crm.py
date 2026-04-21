import streamlit as st
from docx import Document
import io
import re

st.set_page_config(page_title="CRM Document Cloud", layout="wide")

TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

def extraer_datos_inteligente(file):
    doc = Document(file)
    texto_completo = "\n".join([p.text for p in doc.paragraphs])
    # Buscar también dentro de tablas del archivo subido
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto_completo += "\n" + celda.text

    datos = {
        "Fecha": re.search(r"Fecha:\s*(.*)", texto_completo, re.I).group(1) if re.search(r"Fecha:\s*(.*)", texto_completo, re.I) else "",
        "Objetivo": re.search(r"Objetivo:\s*(.*)", texto_completo, re.I).group(1) if re.search(r"Objetivo:\s*(.*)", texto_completo, re.I) else "",
        "Asistentes": []
    }
    
    # Extraer asistentes (busca líneas con nombre y puesto o de la tabla de asistentes)
    lineas = texto_completo.split('\n')
    for linea in lineas:
        if "," in linea and len(linea.split(",")) == 2:
            datos["Asistentes"].append(linea.strip())
            
    return datos

def procesar_documento(template_path, f_val, o_val, asistentes_lista):
    doc = Document(template_path)
    
    # 1. Reemplazar texto en párrafos y tablas
    for p in doc.paragraphs:
        if "Fecha:" in p.text: p.text = f"Fecha: {f_val}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {o_val}"

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if "Fecha:" in celda.text: celda.text = celda.text.replace("Fecha:", f"Fecha: {f_val}")
                if "Objetivo:" in celda.text: celda.text = celda.text.replace("Objetivo:", f"Objetivo: {o_val}")

    # 2. Llenar tabla de Asistentes correctamente
    for tabla in doc.tables:
        # Identificar la tabla de asistentes por su cabecera
        if "Nombre" in tabla.cell(0, 0).text and "Puesto" in tabla.cell(0, 1).text:
            # ELIMINAR filas existentes excepto la cabecera para evitar duplicados
            while len(tabla.rows) > 1:
                tbl = tabla._tbl
                tr = tabla.rows[-1]._tr
                tbl.remove(tr)
            
            # Agregar los nuevos datos
            for asis in asistentes_lista:
                if asis.strip():
                    nueva_fila = tabla.add_row().cells
                    partes = asis.split(",")
                    nueva_fila[0].text = partes[0].strip()
                    if len(partes) > 1:
                        nueva_fila[1].text = partes[1].strip()
            break
            
    return doc

st.title("🚀 Generador CRM Profesional")

opcion = st.selectbox("Selecciona la plantilla:", list(TEMPLATES.keys()))
archivo_info = st.file_uploader("Sube el Word con la información", type=["docx"])

if archivo_info:
    datos = extraer_datos_inteligente(archivo_info)
    
    with st.form("editor"):
        col1, col2 = st.columns(2)
        with col1:
            f_input = st.text_input("Fecha encontrada:", value=datos["Fecha"])
            o_input = st.text_area("Objetivo encontrado:", value=datos["Objetivo"])
        with col2:
            asis_input = st.text_area("Asistentes (Nombre, Puesto):", 
                                     value="\n".join(datos["Asistentes"]),
                                     help="Escribe un nombre y puesto por línea separados por coma.")
        
        if st.form_submit_button("Generar y Descargar"):
            doc_final = procesar_documento(TEMPLATES[opcion], f_input, o_input, asis_input.split('\n'))
            
            target = io.BytesIO()
            doc_final.save(target)
            target.seek(0)
            
            st.success("✅ Archivo corregido generado con éxito")
            st.download_button("Descargar Word", target, file_name=f"Final_{opcion}.docx")