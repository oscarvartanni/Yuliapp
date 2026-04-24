import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
import io
import os

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

# --- 2. DEFINICIÓN DE PLANTILLAS ---
TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# --- 3. FUNCIONES AUXILIARES ---
def aplicar_poppins(run, size=11):
    run.font.name = 'Poppins'
    run.font.size = Pt(size)

def rellenar_tabla(tabla, texto_lineas, columnas):
    """Limpia y rellena tablas de forma genérica."""
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    for linea in texto_lineas.split('\n'):
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        for i in range(min(len(partes), columnas)):
            nueva_fila[i].text = partes[i].strip()
            for p in nueva_fila[i].paragraphs:
                for run in p.runs: aplicar_poppins(run)

def procesar_word(template_path, datos, es_gap=False):
    doc = Document(template_path)
    
    # Rellenar Párrafos (Fecha y Objetivo)
    for p in doc.paragraphs:
        if "Fecha:" in p.text:
            p.text = "Fecha: "
            aplicar_poppins(p.add_run(datos.get('Fecha', '')))
        elif "Objetivo:" in p.text:
            p.text = "Objetivo: "
            aplicar_poppins(p.add_run(datos.get('Objetivo', '')))
        elif "Puntos discutidos:" in p.text and not es_gap:
            p.text = "Puntos discutidos:"
            lineas = datos.get('Puntos Discutidos', '').split('\n')
            for i, linea in enumerate(lineas, 1):
                if linea.strip():
                    nuevo_p = p.insert_paragraph_before(f"{i}. {linea.strip()}")
                    aplicar_poppins(nuevo_p.runs[0] if nuevo_p.runs else nuevo_p.add_run())

    # Rellenar Tablas según el tipo de documento
    for tabla in doc.tables:
        header = tabla.cell(0,0).text.lower()
        if "nombre" in header: # Asistentes
            rellenar_tabla(tabla, datos.get("Asistentes", ""), 2)
        
        # Lógica para Minuta (M100)
        elif "pendientes del cliente" in header:
            rellenar_tabla(tabla, datos.get("Pendientes Cliente", ""), 3)
        elif "pendientes mycloud" in header:
            rellenar_tabla(tabla, datos.get("Pendientes Mycloud", ""), 3)
        
        # Lógica para Gap Analysis (M102)
        elif "módulos" in header or "item" in header:
            if "módulo" in header: rellenar_tabla(tabla, datos.get("Modulos", ""), 4)
            elif "descripción" in header: rellenar_tabla(tabla, datos.get("Custom", ""), 2)
        elif "pendientes" in header and es_gap:
            rellenar_tabla(tabla, datos.get("Pendientes_Gap", ""), 3)
        elif "web services" in header:
            rellenar_tabla(tabla, datos.get("WebServices", ""), 4)
        elif "workflows" in header:
            rellenar_tabla(tabla, datos.get("Workflows", ""), 5)
            
    return doc

# --- 4. INTERFAZ PRINCIPAL ---
st.title("🚀 Generador CRM Profesional")

with st.sidebar:
    st.header("Configuración")
    logo_web = st.file_uploader("Sube el logo:", type=["png", "jpg", "jpeg"])
    if logo_web: st.image(logo_web, use_container_width=True)
    st.divider()
    opcion = st.selectbox("Plantilla:", list(TEMPLATES.keys()))

# --- 5. FORMULARIO DINÁMICO ---
with st.form(key="form_dinamico"):
    st.subheader(f"Campos para: {opcion}")
    col1, col2 = st.columns(2)
    
    with col1:
        fecha = st.text_input("Fecha")
        asistentes = st.text_area("Asistentes (Nombre, Cargo)", height=100)
        objetivo = st.text_area("Objetivo / Alcance", height=100)

    with col2:
        if opcion == "M102 Gap Analysis":
            modulos = st.text_area("Módulos (Item, Nombre, Descripción, Estatus)", placeholder="1, Ventas, Ajuste de campos, Pendiente")
            pendientes_gap = st.text_area("Entrega / Pendientes (Tarea, Responsable, Fecha)")
            custom = st.text_area("Custom Functions (Item, Descripción)")
            # Campos extra abajo para Gap
            ws = st.text_area("Web Services (Item, Nombre, Tipo, Parámetros)")
            wf = st.text_area("Workflows (Item, Módulo, Cuándo, Qué, Acciones)")
        else:
            puntos = st.text_area("Puntos Discutidos (Uno por línea)", height=150)
            p_cliente = st.text_area("Pendientes Cliente (Tarea, Responsable, Fecha)", height=100)
            p_mycloud = st.text_area("Pendientes Mycloud (Tarea, Responsable, Fecha)", height=100)

    submit = st.form_submit_button("🔨 GENERAR DOCUMENTO")

if submit:
    # Recopilar datos según la plantilla
    es_gap = (opcion == "M102 Gap Analysis")
    if es_gap:
        datos_finales = {
            "Fecha": fecha, "Asistentes": asistentes, "Objetivo": objetivo,
            "Modulos": modulos, "Pendientes_Gap": pendientes_gap, 
            "Custom": custom, "WebServices": ws, "Workflows": wf
        }
    else:
        datos_finales = {
            "Fecha": fecha, "Asistentes": asistentes, "Objetivo": objetivo,
            "Puntos Discutidos": puntos, "Pendientes Cliente": p_cliente, "Pendientes Mycloud": p_mycloud
        }

    try:
        doc = procesar_word(TEMPLATES[opcion], datos_finales, es_gap=es_gap)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.success("✅ ¡Documento generado con éxito!")
        st.download_button("📥 Descargar Word", buf, f"{opcion}.docx")
    except Exception as e:
        st.error(f"Error: {e}")
