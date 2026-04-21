import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="CRM Pro - Generador con Logo", layout="wide")
st.title("🚀 Generador de Documentos CRM")

TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

CAMPOS_CONFIG = {
    "M102 Gap Analysis": ["Fecha", "Objetivo", "Asistentes", "Módulos", "Pendientes"],
    "M100 Minuta": ["Fecha", "Objetivo", "Asistentes", "Puntos discutidos", "Acuerdos"],
    "M101 Escenarios": ["Fecha", "Objetivo", "Escenarios de prueba"]
}

def extraer_texto_base(file):
    doc = Document(file)
    texto = "\n".join([p.text for p in doc.paragraphs])
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto += "\n" + celda.text
    return texto

def generar_documento(template_path, datos_finales, logo_file=None):
    doc = Document(template_path)
    
    # 1. Insertar Logo en el Encabezado (si se subió uno)
    if logo_file:
        section = doc.sections[0]
        header = section.header
        # Limpiar encabezado previo si existe y añadir el nuevo
        paragraph = header.paragraphs[0]
        paragraph.alignment = 1 # 1 = Centrado
        run = paragraph.add_run()
        run.add_picture(logo_file, width=Inches(1.5)) # Ajusta el tamaño aquí

    # 2. Reemplazo de texto en párrafos
    for p in doc.paragraphs:
        for campo, valor in datos_finales.items():
            etiqueta = f"{campo}:"
            if etiqueta in p.text:
                texto_sustituto = "" if valor == "[OMITIDO]" else valor
                p.text = f"{etiqueta} {texto_sustituto}"

    # 3. Lógica para tabla de Asistentes
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
st.sidebar.header("⚙️ Configuración Visual")
logo_subido = st.sidebar.file_uploader("Insertar Logo (Imagen):", type=["png", "jpg", "jpeg"])

opcion = st.selectbox("1. Selecciona plantilla:", list(TEMPLATES.keys()))
archivo_subido = st.file_uploader("2. Sube el archivo con información:", type=["docx"])

if archivo_subido:
    texto_fuente = extraer_texto_base(archivo_subido)
    datos_recolectados = {}
    
    with st.form("form_campos"):
        for campo in CAMPOS_CONFIG[opcion]:
            st.markdown(f"**{campo}**")
            col_txt, col_chk = st.columns([4, 1])
            
            match = re.search(rf"{campo}:\s*(.*)", texto_fuente, re.I)
            sugerencia = match.group(1).strip() if match else ""
            
            with col_txt:
                val_usuario = st.text_area(f"Editar {campo}", value=sugerencia, key=f"in_{campo}", height=80)
            with col_chk:
                marcado_omitir = st.checkbox("Omitir", key=f"om_{campo}")
            
            datos_recolectados[campo] = "[OMITIDO]" if marcado_omitir else val_usuario
            
        boton_preparar = st.form_submit_button("🔨 Generar Archivo")

    if boton_preparar:
        documento_final = generar_documento(TEMPLATES[opcion], datos_recolectados, logo_subido)
        
        buffer = io.BytesIO()
        documento_final.save(buffer)
        buffer.seek(0)
        
        st.success("✅ ¡Documento generado!")
        st.download_button("📥 Descargar Word con Logo", buffer, f"Final_{opcion}.docx")
