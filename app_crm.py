import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="CRM Pro - Generador Experto", layout="wide")
st.title("🚀 Generador de Documentos CRM")

TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# Configuración detallada para Minuta
CAMPOS_MINUTA = [
    {"id": "Fecha", "label": "Fecha de la reunión", "tipo": "text"},
    {"id": "Objetivo", "label": "Objetivo de la sesión", "tipo": "area"},
    {"id": "Asistentes", "label": "Asistentes (Formato: Nombre, Puesto - Uno por línea)", "tipo": "area", "help": "Ejemplo:\nOscar, Director\nYuliana, Project Manager"},
    {"id": "Puntos discutidos", "label": "Puntos discutidos (Uno por línea)", "tipo": "area", "help": "Cada línea aparecerá en un numeral diferente."},
    {"id": "Pendientes cliente", "label": "Pendientes del Cliente (Formato: Tarea, Responsable, Fecha)", "tipo": "area"},
    {"id": "Pendientes Mycloud", "label": "Pendientes Mycloud (Formato: Tarea, Responsable, Fecha)", "tipo": "area"}
]

def extraer_texto_base(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs] + [c.text for t in doc.tables for r in t.rows for c in r.cells])

def rellenar_tabla_triple(tabla, texto_area):
    """Lógica para tablas de 3 columnas (Pendientes, Responsable, Fecha)"""
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    lineas = [l.strip() for l in texto_area.split('\n') if l.strip()]
    for linea in lineas:
        partes = linea.split(',')
        nueva_fila = tabla.add_row().cells
        nueva_fila[0].text = partes[0].strip() if len(partes) > 0 else ""
        nueva_fila[1].text = partes[1].strip() if len(partes) > 1 else ""
        nueva_fila[2].text = partes[2].strip() if len(partes) > 2 else ""

def generar_documento(template_path, datos, logo=None):
    doc = Document(template_path)
    
    if logo:
        header_para = doc.sections[0].header.paragraphs[0]
        header_para.alignment = 1
        header_para.add_run().add_picture(logo, width=Inches(1.5))

    # Reemplazo de Puntos Discutidos en los numerales 1, 2, 3...
    puntos = [p.strip() for p in datos.get("Puntos discutidos", "").split('\n') if p.strip()]
    idx_punto = 0
    for p in doc.paragraphs:
        # Buscamos los párrafos que tienen los números de la lista
        if re.match(r"^\d+\.", p.text.strip()) and idx_punto < len(puntos):
            p.text = f"{idx_punto + 1}. {puntos[idx_punto]}"
            idx_punto += 1
        # Reemplazo estándar de Fecha y Objetivo
        for campo in ["Fecha", "Objetivo"]:
            if f"{campo}:" in p.text:
                p.text = f"{campo}: {datos.get(campo, '')}" if datos.get(campo) != "[OMITIDO]" else f"{campo}:"

    # Manejo de tablas
    for tabla in doc.tables:
        cabecera = tabla.cell(0, 0).text.lower()
        # 1. Tabla Asistentes
        if "nombre" in cabecera and "puesto" in cabecera:
            if datos.get("Asistentes") != "[OMITIDO]":
                while len(tabla.rows) > 1: tabla._tbl.remove(tabla.rows[-1]._tr)
                for asis in datos.get("Asistentes", "").split('\n'):
                    if ',' in asis:
                        nf = tabla.add_row().cells
                        p = asis.split(',')
                        nf[0].text, nf[1].text = p[0].strip(), p[1].strip()
        
        # 2. Tablas de Pendientes
        elif "pendientes del cliente" in cabecera:
            rellenar_tabla_triple(tabla, datos.get("Pendientes cliente", ""))
        elif "pendientes mycloud" in cabecera:
            rellenar_tabla_triple(tabla, datos.get("Pendientes Mycloud", ""))
            
    return doc

# --- INTERFAZ ---
st.sidebar.header("🎨 Personalización")
logo = st.sidebar.file_uploader("Logo de la empresa", type=["png", "jpg"])

opcion = st.selectbox("Selecciona plantilla:", list(TEMPLATES.keys()))
archivo = st.file_uploader("Sube el Word con la info:", type=["docx"])

if archivo:
    texto_base = extraer_texto_base(archivo)
    datos_finales = {}
    
    with st.form("main_form"):
        st.info("💡 Tip: Para tablas, separa los datos con comas. Ejemplo: Tarea, Juan, 25/Abril")
        
        campos = CAMPOS_MINUTA if opcion == "M100 Minuta" else [] # Aquí podrías añadir los de M102
        
        for c in campos:
            st.markdown(f"**{c['label']}**")
            col_in, col_om = st.columns([4, 1])
            with col_in:
                val = st.text_area("Contenido", key=f"in_{c['id']}", help=c.get("help", ""))
            with col_om:
                omit = st.checkbox("Omitir", key=f"om_{c['id']}")
            datos_finales[c['id']] = "[OMITIDO]" if omit else val
        
        if st.form_submit_button("🔨 Preparar Minuta"):
            doc_out = generar_documento(TEMPLATES[opcion], datos_finales, logo)
            buffer = io.BytesIO()
            doc_out.save(buffer)
            buffer.seek(0)
            st.success("✅ ¡Minuta lista!")
            st.download_button("📥 Descargar Minuta Corregida", buffer, f"Minuta_{datos_finales.get('Fecha','')}.docx")
