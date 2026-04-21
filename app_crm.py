import streamlit as st
from docx import Document
import io
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="CRM Pro - Generador", layout="wide")
st.title("🚀 Generador de Documentos CRM")

TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# Definición de campos por cada tipo de documento
CAMPOS_CONFIG = {
    "M102 Gap Analysis": ["Fecha", "Objetivo", "Asistentes", "Módulos", "Pendientes"],
    "M100 Minuta": ["Fecha", "Objetivo", "Asistentes", "Puntos discutidos", "Acuerdos"],
    "M101 Escenarios": ["Fecha", "Objetivo", "Escenarios de prueba"]
}

def extraer_datos_base(file):
    doc = Document(file)
    texto = "\n".join([p.text for p in doc.paragraphs])
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto += "\n" + celda.text
    return texto

def generar_documento(template_path, datos_finales):
    doc = Document(template_path)
    
    # 1. Reemplazo de texto en párrafos
    for p in doc.paragraphs:
        for campo, valor in datos_finales.items():
            if f"{campo}:" in p.text:
                p.text = f"{campo}: {valor}"

    # 2. Lógica para la tabla de Asistentes (si el campo existe)
    if "Asistentes" in datos_finales and datos_finales["Asistentes"] != "[OMITIDO]":
        asistentes = [a.strip() for a in datos_finales["Asistentes"].split('\n') if a.strip()]
        for tabla in doc.tables:
            if len(tabla.rows) > 0 and "Nombre" in tabla.cell(0, 0).text:
                while len(tabla.rows) > 1:
                    tabla._tbl.remove(tabla.rows[-1]._tr)
                for asis in asistentes:
                    nueva_fila = tabla.add_row().cells
                    partes = asis.split(',')
                    nueva_fila[0].text = partes[0].strip()
                    if len(partes) > 1:
                        nueva_fila[1].text = partes[1].strip()
                break
    return doc

# --- INTERFAZ ---
opcion = st.selectbox("1. Selecciona plantilla:", list(TEMPLATES.keys()))
archivo_subido = st.file_uploader("2. Sube el archivo con información:", type=["docx"])

if archivo_subido:
    texto_extraido = extraer_datos_base(archivo_subido)
    datos_para_procesar = {}
    
    st.subheader("3. Validar información campo por campo")
    
    with st.form("form_detallado"):
        for campo in CAMPOS_CONFIG[opcion]:
            st.markdown(f"**Configuración para: {campo}**")
            col_input, col_check = st.columns([4, 1])
            
            # Intento de pre-rellenado básico
            busqueda = re.search(rf"{campo}:\s*(.*)", texto_extraido, re.I)
            valor_default = busqueda.group(1).strip() if busqueda else ""
            
            with col_input:
                # Usamos text_area para permitir varias líneas en campos como 'Objetivo'
                texto_usuario = st.text_area(f"Contenido de {campo}", value=valor_default, key=f"txt_{campo}", height=100)
            
            with col_check:
                omitir = st.checkbox("Omitir", key=f"omit_{campo}")
            
            datos_para_procesar[campo] = "[OMITIDO]" if omitir else texto_usuario
            st.divider()
            
        procesar = st.form_submit_button("🔨 Preparar Documento Final")

    if procesar:
        doc_final = generar_documento(TEMPLATES[opcion], datos_para_procesar)
        
        buffer = io.BytesIO()
        doc_final.save(buffer)
        buffer.seek(0)
        
        st.success("✅ ¡Documento generado con las opciones seleccionadas!")
        st.download_button(
            label="📥 Descargar Word Final",
            data=buffer,
            file_name=f"Final_{opcion.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        )
