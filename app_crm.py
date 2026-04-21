import streamlit as st
from docx import Document
import io
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="CRM Pro - Generador", layout="wide")
st.title("🚀 Generador de Documentos CRM")

# Nombres de tus archivos en GitHub
TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

def extraer_datos_base(file):
    doc = Document(file)
    texto = "\n".join([p.text for p in doc.paragraphs])
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto += "\n" + celda.text
    return {
        "Fecha": re.search(r"Fecha:\s*(.*)", texto, re.I).group(1) if re.search(r"Fecha:\s*(.*)", texto, re.I) else "",
        "Objetivo": re.search(r"Objetivo:\s*(.*)", texto, re.I).group(1) if re.search(r"Objetivo:\s*(.*)", texto, re.I) else ""
    }

def generar_documento(template_path, fecha, objetivo, asistentes_texto):
    doc = Document(template_path)
    # Reemplazo en párrafos
    for p in doc.paragraphs:
        if "Fecha:" in p.text: p.text = f"Fecha: {fecha}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {objetivo}"
    
    # Manejo de tabla de asistentes (Nombre, Puesto)
    asistentes = [a.strip() for a in asistentes_texto.split('\n') if a.strip()]
    for tabla in doc.tables:
        # Detectamos la tabla de asistentes por sus encabezados [cite: 5, 25]
        if len(tabla.rows) > 0 and "Nombre" in tabla.cell(0, 0).text:
            # Borrar filas de ejemplo (excepto la cabecera)
            while len(tabla.rows) > 1:
                tr = tabla.rows[-1]._tr
                tabla._tbl.remove(tr)
            # Insertar nuevos datos
            for asis in asistentes:
                nueva_fila = tabla.add_row().cells
                partes = asis.split(',')
                nueva_fila[0].text = partes[0].strip()
                if len(partes) > 1:
                    nueva_fila[1].text = partes[1].strip()
            break
    return doc

# --- INTERFAZ ---
opcion = st.selectbox("Selecciona plantilla:", list(TEMPLATES.keys()))
archivo_subido = st.file_uploader("Sube el archivo con información:", type=["docx"])

if archivo_subido:
    info = extraer_datos_base(archivo_subido)
    
    with st.form("form_datos"):
        col1, col2 = st.columns(2)
        with col1:
            f_val = st.text_input("Fecha:", value=info["Fecha"])
            o_val = st.text_area("Objetivo:", value=info["Objetivo"])
        with col2:
            asis_val = st.text_area("Asistentes (Nombre, Puesto separados por coma):")
        
        procesar = st.form_submit_button("Preparar Documento")

    # AQUÍ ESTABA EL ERROR: Ahora el espaciado es correcto y está fuera del 'with st.form'
    if procesar:
        doc_final = generar_documento(TEMPLATES[opcion], f_val, o_val, asis_val)
        
        buffer = io.BytesIO()
        doc_final.save(buffer)
        buffer.seek(0)
        
        st.success("✅ ¡Documento listo!")
        st.download_button(
            label="📥 Descargar Word Final",
            data=buffer,
            file_name=f"Final_{opcion.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
