import streamlit as st
from docx import Document
import io
import re
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

# --- 1. DEFINICIÓN DE OPCIONES (Para evitar NameError) ---
TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# --- 2. BARRA LATERAL (LOGO Y SELECTOR) ---
with st.sidebar:
    # Logo automático
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.warning("⚠️ Sube 'logo.png' a GitHub")
    
    st.title("Configuración")
    # Definimos 'opcion' aquí mismo
    opcion = st.selectbox("Selecciona Plantilla:", list(TEMPLATES.keys()))
    st.divider()
    st.info("💡 Tip: En las tablas, separa los datos por comas.")

# --- 3. CONFIGURACIÓN DE CAMPOS ---
CONFIG_DETALLADA = {
    "M100 Minuta": {
        "Fecha": "Ej: 22 de Abril 2026",
        "Objetivo": "Objetivo de la reunión",
        "Asistentes": "Nombres separados por comas o saltos de línea",
        "Puntos Discutidos": "Detalle de los puntos tratados",
        "Pendientes Cliente": "Tarea, Responsable, Fecha",
        "Pendientes Mycloud": "Tarea, Responsable, Fecha"
    },
    "M102 Gap Analysis": {
        "Fecha": "Fecha del análisis",
        "Objetivo": "Objetivo del Gap Analysis",
        "Asistentes": "Participantes",
        "Módulos": "Módulo, Función, Estatus, Observaciones",
        "Pendientes General": "Tarea, Responsable, Fecha"
    },
    "M101 Escenarios": {
        "Fecha": "Fecha de pruebas",
        "Objetivo": "Objetivo del UAT",
        "Escenarios de Prueba": "Descripción, Módulos, Responsable"
    }
}

# --- 4. FUNCIONES DE PROCESAMIENTO ---
def extraer_informacion(archivo_subido):
    datos = {k: "" for k in ["Fecha", "Objetivo", "Asistentes", "Puntos Discutidos", "Pendientes Cliente", "Pendientes Mycloud"]}
    if not archivo_subido: return datos
    
    try:
        doc = Document(archivo_subido)
        parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        
        current_key = None
        for p in parrafos:
            p_lower = p.lower()
            if "fecha:" in p_lower:
                datos["Fecha"] = p.split(":", 1)[1].strip()
                current_key = "Fecha"
            elif "asistentes:" in p_lower:
                datos["Asistentes"] = p.split(":", 1)[1].strip()
                current_key = "Asistentes"
            elif "objetivo:" in p_lower:
                datos["Objetivo"] = p.split(":", 1)[1].strip()
                current_key = "Objetivo"
            elif "puntos discutidos:" in p_lower:
                datos["Puntos Discutidos"] = p.split(":", 1)[1].strip()
                current_key = "Puntos Discutidos"
            elif "pendientes del cliente" in p_lower:
                current_key = "Pendientes Cliente"
            elif "pendientes mycloud" in p_lower:
                current_key = "Pendientes Mycloud"
            elif current_key and not ":" in p_lower:
                datos[current_key] += "\n" + p
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

def procesar_word(template_name, datos_usuario):
    doc = Document(TEMPLATES[template_name])
    for p in doc.paragraphs:
        p_lower = p.text.lower()
        if "fecha:" in p_lower: p.text = f"Fecha: {datos_usuario.get('Fecha', '')}"
        elif "objetivo:" in p_lower: p.text = f"Objetivo: {datos_usuario.get('Objetivo', '')}"
        elif "puntos discutidos:" in p_lower:
            p.text = f"Puntos discutidos:\n{datos_usuario.get('Puntos Discutidos', '')}"
    
    for tabla in doc.tables:
        cabecera = tabla.cell(0,0).text.lower()
        if "nombre" in cabecera:
            rellenar_tabla_estandar(tabla, datos_usuario.get("Asistentes", ""), 2)
        elif "pendientes del cliente" in cabecera:
            rellenar_tabla_estandar(tabla, datos_usuario.get("Pendientes Cliente", ""), 3)
        elif "pendientes mycloud" in cabecera:
            rellenar_tabla_estandar(tabla, datos_usuario.get("Pendientes Mycloud", ""), 3)
    return doc

# --- 5. INTERFAZ PRINCIPAL ---
st.title("🚀 Generador CRM Profesional")

archivo_ref = st.file_uploader("📂 Sube archivo de referencia:", type=["docx"])
datos_auto = extraer_informacion(archivo_ref)

# Formulario
with st.form(key=f"form_{opcion}"):
    st.subheader(f"Editando: {opcion}")
    campos = CONFIG_DETALLADA[opcion]
    datos_finales = {}
    
    c1, c2 = st.columns(2)
    for i, (campo, placeholder) in enumerate(campos.items()):
        col = c1 if i % 2 == 0 else c2
        with col:
            val = datos_auto.get(campo, "")
            if campo == "Fecha":
                datos_finales[campo] = st.text_input(campo, value=val)
            else:
                datos_finales[campo] = st.text_area(campo, value=val, height=150)
    
    enviado = st.form_submit_button("🔨 GENERAR")

# Botón de descarga
if enviado:
    doc_res = procesar_word(opcion, datos_finales)
    buf = io.BytesIO()
    doc_res.save(buf)
    st.success("✅ ¡Listo!")
    st.download_button("📥 Descargar Word", buf.getvalue(), file_name=f"{opcion}.docx")
