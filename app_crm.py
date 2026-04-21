import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

st.set_page_config(page_title="CRM Pro - Generador", layout="wide")
st.title("🚀 Generador de Documentos CRM")

TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

def rellenar_tabla_triple(tabla, texto_area):
    """Lógica para tablas de 3 columnas: Tarea, Responsable, Fecha"""
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
        try:
            header_para = doc.sections[0].header.paragraphs[0]
            header_para.alignment = 1
            header_para.add_run().add_picture(logo, width=Inches(1.5))
        except: pass

    # 1. Reemplazo de Puntos Discutidos (Lista numerada)
    puntos = [p.strip() for p in datos.get("Puntos", "").split('\n') if p.strip()]
    idx = 0
    for p in doc.paragraphs:
        if re.match(r"^\d+\.", p.text.strip()) and idx < len(puntos):
            p.text = f"{idx + 1}. {puntos[idx]}"
            idx += 1
        if "Fecha:" in p.text: p.text = f"Fecha: {datos.get('Fecha', '')}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {datos.get('Objetivo', '')}"

    # 2. Manejo de Tablas
    for tabla in doc.tables:
        cabecera = tabla.cell(0, 0).text.lower()
        if "nombre" in cabecera and "puesto" in cabecera:
            while len(tabla.rows) > 1: tabla._tbl.remove(tabla.rows[-1]._tr)
            for asis in datos.get("Asistentes", "").split('\n'):
                if ',' in asis:
                    nf = tabla.add_row().cells
                    p = asis.split(',')
                    nf[0].text, nf[1].text = p[0].strip(), p[1].strip()
        elif "pendientes del cliente" in cabecera:
            rellenar_tabla_triple(tabla, datos.get("PC", ""))
        elif "pendientes mycloud" in cabecera:
            rellenar_tabla_triple(tabla, datos.get("PM", ""))
    return doc

# --- INTERFAZ ---
st.sidebar.header("🎨 Personalización")
logo_file = st.sidebar.file_uploader("Logo Empresa", type=["png", "jpg"])

opcion = st.selectbox("Selecciona plantilla:", list(TEMPLATES.keys()))
archivo = st.file_uploader("Sube el Word base:", type=["docx"])

if archivo:
    # Usamos session_state para guardar el archivo generado y que no dé error de formulario
    if 'archivo_listo' not in st.session_state:
        st.session_state.archivo_listo = None

    with st.form("main_form"):
        st.info("📝 Instrucciones: En tablas, separa con comas. Ejemplo: Tarea X, Juan Perez, 21/04/2026")
        col1, col2 = st.columns(2)
        with col1:
            f = st.text_input("Fecha")
            obj = st.text_area("Objetivo")
            puntos = st.text_area("Puntos discutidos (Uno por línea)")
        with col2:
            asis = st.text_area("Asistentes (Nombre, Puesto)")
            pc = st.text_area("Pendientes Cliente (Tarea, Responsable, Fecha)")
            pm = st.text_area("Pendientes Mycloud (Tarea, Responsable, Fecha)")
        
        submitted = st.form_submit_button("🔨 Preparar Archivo")
        
        if submitted:
            datos_finales = {"Fecha": f, "Objetivo": obj, "Puntos": puntos, "Asistentes": asis, "PC": pc, "PM": pm}
            doc_gen = generar_documento(TEMPLATES[opcion], datos_finales, logo_file)
            buffer = io.BytesIO()
            doc_gen.save(buffer)
            st.session_state.archivo_listo = buffer.getvalue()

    # EL BOTÓN DE DESCARGA SIEMPRE FUERA DEL FORMULARIO
    if st.session_state.archivo_listo:
        st.success("✅ ¡Documento preparado correctamente!")
        st.download_button(
            label="📥 DESCARGAR ARCHIVO FINAL",
            data=st.session_state.archivo_listo,
            file_name=f"Final_{opcion.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
