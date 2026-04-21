import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="CRM Document Generator", layout="wide")
st.title("📝 Generador de Documentos CRM Profesional")

TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

def replace_text_in_paragraph(paragraph, key, value):
    """Reemplaza texto manteniendo el formato lo mejor posible."""
    if key in paragraph.text:
        # Combinar todos los 'runs' para que la búsqueda no falle
        full_text = paragraph.text.replace(key, f"{key} {value}")
        paragraph.text = "" # Limpiar
        paragraph.add_run(full_text)

def procesar_tabla(tabla, datos_texto, cols=2):
    """Limpia la tabla y añade filas nuevas."""
    # Eliminar filas viejas (excepto cabecera)
    while len(tabla.rows) > 1:
        row = tabla.rows[-1]
        tabla._tbl.remove(row._tr)
    
    # Añadir datos nuevos
    lineas = [l.strip() for l in datos_texto.split('\n') if l.strip()]
    for linea in lineas:
        partes = [p.strip() for p in linea.split(',')]
        nueva_fila = tabla.add_row().cells
        for i in range(min(len(partes), cols)):
            nueva_fila[i].text = partes[i]

def generar_documento(template_path, datos, logo=None):
    doc = Document(template_path)
    
    # 1. Logo
    if logo:
        try:
            header = doc.sections[0].header
            p = header.paragraphs[0]
            p.alignment = 1
            p.add_run().add_picture(logo, width=Inches(1.2))
        except: pass

    # 2. Reemplazo de campos base (Fecha y Objetivo)
    for p in doc.paragraphs:
        if "Fecha:" in p.text: replace_text_in_paragraph(p, "Fecha:", datos['Fecha'])
        if "Objetivo:" in p.text: replace_text_in_paragraph(p, "Objetivo:", datos['Objetivo'])

    # 3. Lógica de Puntos Discutidos (Numerados)
    puntos = [p.strip() for p in datos.get("Puntos", "").split('\n') if p.strip()]
    idx = 0
    for p in doc.paragraphs:
        if re.match(r"^\d+\.", p.text.strip()) and idx < len(puntos):
            p.text = f"{idx+1}. {puntos[idx]}"
            idx += 1

    # 4. Tablas
    for tabla in doc.tables:
        header_text = tabla.cell(0, 0).text.lower()
        
        # Tabla de Asistentes
        if "nombre" in header_text:
            procesar_tabla(tabla, datos.get("Asistentes", ""), cols=2)
        
        # Tablas de Pendientes (Cliente y Mycloud)
        elif "pendientes del cliente" in header_text or "cliente" in header_text and "responsable" in tabla.rows[0].cells[1].text.lower():
            procesar_tabla(tabla, datos.get("PC", ""), cols=3)
        
        elif "pendientes mycloud" in header_text or "mycloud" in header_text and "responsable" in tabla.rows[0].cells[1].text.lower():
            procesar_tabla(tabla, datos.get("PM", ""), cols=3)
            
    return doc

# --- INTERFAZ ---
st.sidebar.header("Opciones")
logo = st.sidebar.file_uploader("Subir Logo", type=["png", "jpg"])

opcion = st.selectbox("Selecciona Plantilla:", list(TEMPLATES.keys()))
archivo_ref = st.file_uploader("Sube el Word con la información (opcional)", type=["docx"])

# Formulario dinámico
with st.form("data_form"):
    st.subheader(f"Datos para: {opcion}")
    col1, col2 = st.columns(2)
    
    with col1:
        f = st.text_input("Fecha")
        obj = st.text_area("Objetivo")
        puntos = st.text_area("Puntos discutidos (Uno por línea)")
        
    with col2:
        st.markdown("**Tablas (Separar por comas)**")
        asis = st.text_area("Asistentes", help="Ejemplo: Juan Perez, Gerente")
        pc = st.text_area("Pendientes Cliente", help="Ejemplo: Entregar accesos, Cliente, 25/Abril")
        pm = st.text_area("Pendientes Mycloud", help="Ejemplo: Configurar portal, Mycloud, 30/Abril")
        
    submit = st.form_submit_button("🔨 Generar Archivo")

if submit:
    if not f or not obj:
        st.error("Por favor llena al menos la Fecha y el Objetivo.")
    else:
        datos_finales = {"Fecha": f, "Objetivo": obj, "Puntos": puntos, "Asistentes": asis, "PC": pc, "PM": pm}
        doc_final = generar_documento(TEMPLATES[opcion], datos_finales, logo)
        
        output = io.BytesIO()
        doc_final.save(output)
        st.success("¡Documento generado con éxito!")
        st.download_button(
            label="📥 Descargar ahora",
            data=output.getvalue(),
            file_name=f"{opcion.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
