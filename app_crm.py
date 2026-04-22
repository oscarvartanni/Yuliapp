import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

# --- INSERTAR LOGO EN LA WEB ---
# Mostramos el logo de Yuliana Muñoz en la parte superior de la barra lateral
st.sidebar.image("https://r2.community.samsung.com/t5/image/serverpage/image-id/5319983i8E8C0A97E9C9F9A1/image-size/large?v=v2&px=999", use_container_width=True)

TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# Configuración de campos y ejemplos específicos (Placeholders)
CONFIG_DETALLADA = {
    "M100 Minuta": {
        "Fecha": "21 de Abril 2026",
        "Objetivo": "Definir alcances del módulo de ventas",
        "Asistentes": "Juan Perez, Gerente\nMaria Garcia, Consultora",
        "Puntos Discutidos": "Revisión de tiempos\nValidación de campos",
        "Pendientes Cliente": "Tarea, Responsable, Fecha",
        "Pendientes Mycloud": "Tarea, Responsable, Fecha"
    },
    "M102 Gap Analysis": {
        "Fecha": "21 de Abril 2026",
        "Objetivo": "Análisis de brechas técnica vs funcional",
        "Asistentes": "Luis Pascal, Dirección\nAlejandro Chávez, Implementación",
        "Módulos": "Ventas, Prospectos, Nuevo Campo, En proceso",
        "Pendientes General": "Tarea, Responsable, Fecha"
    },
    "M101 Escenarios": {
        "Fecha": "21 de Abril 2026",
        "Objetivo": "Pruebas de aceptación de usuario (UAT)",
        "Escenarios de Prueba": "Descripción del escenario, Módulos, Responsable"
    }
}

def extraer_informacion(archivo_subido):
    datos = {"Fecha": "", "Objetivo": ""}
    if archivo_subido:
        try:
            doc = Document(archivo_subido)
            texto_completo = "\n".join([p.text for p in doc.paragraphs])
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells: texto_completo += "\n" + c.text
            f_match = re.search(r"Fecha[:\s]+([^\n]*)", texto_completo, re.IGNORECASE)
            o_match = re.search(r"Objetivo[:\s]+([^\n]*)", texto_completo, re.IGNORECASE)
            if f_match: datos["Fecha"] = f_match.group(1).strip()
            if o_match: datos["Objetivo"] = o_match.group(1).strip()
        except: pass
    return datos

def rellenar_tabla_estandar(tabla, texto_lineas, columnas):
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    for linea in texto_lineas.split('\n'):
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        for i in range(min(len(partes), columnas)):
            nueva_fila[i].text = partes[i].strip()

def rellenar_tabla_escenarios(tabla, texto_lineas):
    """Especial para M101: 4 columnas (No, Descripción, Módulos, Responsable)"""
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    for idx, linea in enumerate(texto_lineas.split('\n')):
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        nueva_fila[0].text = str(idx + 1)
        for i in range(min(len(partes), 3)):
            nueva_fila[i+1].text = partes[i].strip()

def procesar_word(template_name, datos_usuario, logo_img=None):
    doc = Document(TEMPLATES[template_name])
    
    if logo_img:
        try:
            header = doc.sections[0].header
            p = header.paragraphs[0]
            p.add_run().add_picture(logo_img, width=Inches(1.2))
        except: pass

    for p in doc.paragraphs:
        if "Fecha:" in p.text: p.text = f"Fecha: {datos_usuario.get('Fecha', '')}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {datos_usuario.get('Objetivo', '')}"

    for tabla in doc.tables:
        cabecera = tabla.cell(0,0).text.lower()
        if "no." in cabecera or "escenario" in cabecera:
            rellenar_tabla_escenarios(tabla, datos_usuario.get("Escenarios de Prueba", ""))
        elif "nombre" in cabecera:
            rellenar_tabla_estandar(tabla, datos_usuario.get("Asistentes", ""), 2)
        elif "pendientes del cliente" in cabecera:
            rellenar_tabla_estandar(tabla, datos_usuario.get("Pendientes Cliente", ""), 3)
        elif "pendientes mycloud" in cabecera:
            rellenar_tabla_estandar(tabla, datos_usuario.get("Pendientes Mycloud", ""), 3)
        elif "módulo" in cabecera:
            rellenar_tabla_estandar(tabla, datos_usuario.get("Módulos", ""), 4)

    return doc

# --- INTERFAZ ---
st.title("🚀 Generador CRM Profesional")

with st.sidebar:
    st.header("Configuración")
    opcion = st.selectbox("Plantilla:", list(TEMPLATES.keys()), key="selector_doc")
    logo_doc = st.file_uploader("Logo para el Documento:", type=["png", "jpg"])
    st.divider()
    st.info("💡 Tip: Para las tablas, separa los datos por comas.")

archivo_ref = st.file_uploader("Sube archivo de referencia (opcional):", type=["docx"])
datos_extraidos = extraer_informacion(archivo_ref)

with st.form(key=f"form_v65_{opcion}"):
    st.subheader(f"Editando: {opcion}")
    config = CONFIG_DETALLADA[opcion]
    datos_finales = {}
    
    col1, col2 = st.columns(2)
    for i, (campo, placeholder) in enumerate(config.items()):
        target_col = col1 if i % 2 == 0 else col2
        with target_col:
            val_init = datos_extraidos.get(campo, "") if campo in ["Fecha", "Objetivo"] else ""
            if campo == "Fecha":
                datos_finales[campo] = st.text_input(campo, value=val_init, placeholder=placeholder)
            else:
                datos_finales[campo] = st.text_area(campo, value=val_init, placeholder=f"Ej: {placeholder}", height=150)
    
    btn = st.form_submit_button("🔨 GENERAR")

if btn:
    with st.spinner("Generando..."):
        doc_final = procesar_word(opcion, datos_finales, logo_doc)
        buf = io.BytesIO()
        doc_final.save(buf)
        st.success("✅ ¡Archivo generado!")
        st.download_button("📥 Descargar Word", buf.getvalue(), file_name=f"{opcion}.docx")
