import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Generador CRM Inteligente", layout="wide")

# Mapeo de archivos en tu GitHub
TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# Configuración de qué campos mostrar por cada documento
CONFIG_CAMPOS = {
    "M100 Minuta": ["Fecha", "Objetivo", "Asistentes", "Puntos Discutidos", "Pendientes Cliente", "Pendientes Mycloud"],
    "M102 Gap Analysis": ["Fecha", "Objetivo", "Asistentes", "Módulos", "Pendientes General"],
    "M101 Escenarios": ["Fecha", "Objetivo", "Escenarios de Prueba"]
}

def limpiar_texto_word(texto):
    """Limpia caracteres extraños y normaliza espacios."""
    return re.sub(r'\s+', ' ', texto).strip()

def extraer_informacion(archivo_subido):
    """Lee el Word subido e intenta extraer Fecha y Objetivo."""
    datos = {"Fecha": "", "Objetivo": ""}
    if archivo_subido:
        doc = Document(archivo_subido)
        # Unir todo el texto (párrafos y tablas)
        todo_el_texto = ""
        for p in doc.paragraphs: todo_el_texto += p.text + "\n"
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells: todo_el_texto += c.text + " "
        
        # Búsqueda con Regex sensible a mayúsculas/minúsculas
        f_match = re.search(r"Fecha:\s*(.*)", todo_el_texto, re.IGNORECASE)
        o_match = re.search(r"Objetivo:\s*(.*)", todo_el_texto, re.IGNORECASE)
        if f_match: datos["Fecha"] = f_match.group(1).strip()
        if o_match: datos["Objetivo"] = o_match.group(1).strip()
    return datos

def rellenar_tabla(tabla, texto_lineas, columnas):
    """Borra filas de ejemplo y llena con la nueva info."""
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    
    for linea in texto_lineas.split('\n'):
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        for i in range(min(len(partes), columnas)):
            nueva_fila[i].text = partes[i].strip()

def procesar_word(template_name, datos_usuario, logo_img=None):
    doc = Document(TEMPLATES[template_name])
    
    if logo_img:
        try: doc.sections[0].header.paragraphs[0].add_run().add_picture(logo_img, width=Inches(1.2))
        except: pass

    # Reemplazo de texto en párrafos y tablas
    for p in doc.paragraphs:
        if "Fecha:" in p.text: p.text = f"Fecha: {datos_usuario.get('Fecha', '')}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {datos_usuario.get('Objetivo', '')}"

    # Lógica de Puntos Discutidos (Numerados 1, 2, 3...)
    puntos = [p.strip() for p in datos_usuario.get("Puntos Discutidos", "").split('\n') if p.strip()]
    if puntos:
        idx = 0
        for p in doc.paragraphs:
            if re.match(r"^\d+\.", p.text.strip()) and idx < len(puntos):
                p.text = f"{idx+1}. {puntos[idx]}"
                idx += 1

    # Lógica de Tablas
    for tabla in doc.tables:
        cabecera = tabla.cell(0,0).text.lower()
        if "nombre" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Asistentes", ""), 2)
        elif "pendientes del cliente" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Pendientes Cliente", ""), 3)
        elif "pendientes mycloud" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Pendientes Mycloud", ""), 3)
        elif "módulo" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Módulos", ""), 4)
            
    return doc

# --- INTERFAZ STREAMLIT ---
st.title("🚀 Generador CRM Profesional")

# 1. Selección de Plantilla
opcion = st.selectbox("Selecciona la plantilla:", list(TEMPLATES.keys()), key="cambio_plantilla")

# 2. Subida de Archivo (esto debería auto-rellenar los campos)
archivo = st.file_uploader("Sube tu Word de referencia para extraer datos:", type=["docx"])
datos_leidos = extraer_informacion(archivo)

# 3. Formulario Dinámico
with st.form("form_dinamico"):
    st.subheader(f"Editando: {opcion}")
    datos_finales = {}
    
    # Generar campos basados en la configuración de la plantilla
    cols = st.columns(2)
    campos_a_mostrar = CONFIG_CAMPOS[opcion]
    
    for i, campo in enumerate(campos_a_mostrar):
        col_idx = i % 2
        with cols[col_idx]:
            # Si el campo es Fecha u Objetivo, intentamos poner lo que leímos del archivo
            default_val = datos_leidos.get(campo, "") if campo in ["Fecha", "Objetivo"] else ""
            
            if campo in ["Fecha"]:
                datos_finales[campo] = st.text_input(campo, value=default_val)
            else:
                # Instrucciones de ayuda según el campo
                ayuda = "Separa por comas: Tarea, Responsable, Fecha" if "Pendientes" in campo or "Asistentes" in campo else ""
                datos_finales[campo] = st.text_area(campo, value=default_val, help=ayuda, height=100)
    
    submitted = st.form_submit_button("Generar Documento")

if submitted:
    logo = st.sidebar.file_uploader("Opcional: Logo", type=["png", "jpg"])
    doc_final = procesar_word(opcion, datos_finales, logo)
    
    buffer = io.BytesIO()
    doc_final.save(buffer)
    st.success("✅ ¡Documento listo para descargar!")
    st.download_button("📥 Descargar Word", buffer.getvalue(), file_name=f"{opcion}.docx")
