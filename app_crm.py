import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
import io
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

# --- 1. DEFINICIÓN DE PLANTILLAS ---
TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# --- 2. BARRA LATERAL (LOGO Y SELECTOR) ---
with st.sidebar:
    st.header("Identidad Visual")
    # Restaurado: Cargador de logo como en tu versión original 
    logo_web = st.file_uploader("1. Sube el logo para la WEB:", type=["png", "jpg", "jpeg"], key="logo_web")
    
    if logo_web:
        st.image(logo_web, use_container_width=True)
    elif os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    
    st.divider()
    st.header("2. Configuración")
    opcion = st.selectbox("Selecciona la Plantilla:", list(TEMPLATES.keys()))
    st.info("💡 Tip: Para las tablas, separa los datos por comas.")

# --- 3. FUNCIONES DE EXTRACCIÓN (ORDEN SECUENCIAL) ---
def iterar_bloques(parent):
    from docx.document import Document as _Document
    parent_elm = parent.element.body if isinstance(parent, _Document) else parent._tc
    for child in parent_elm.iterchildren():
        if child.tag.endswith('p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('tbl'):
            yield Table(child, parent)

def extraer_informacion(archivo_subido):
    datos = {k: "" for k in ["Fecha", "Objetivo", "Asistentes", "Puntos Discutidos", "Pendientes Cliente", "Pendientes Mycloud"]}
    if not archivo_subido: return datos
    
    try:
        doc = Document(archivo_subido)
        contexto_actual = None
        
        for bloque in iterar_bloques(doc):
            if isinstance(bloque, Paragraph):
                texto = bloque.text.strip()
                texto_l = texto.lower()
                
                if "fecha:" in texto_l:
                    datos["Fecha"] = texto.split(":", 1)[1].strip()
                elif "asistentes:" in texto_l: contexto_actual = "Asistentes"
                elif "objetivo:" in texto_l:
                    datos["Objetivo"] = texto.split(":", 1)[1].strip()
                    contexto_actual = "Objetivo"
                elif "puntos discutidos:" in texto_l: contexto_actual = "Puntos Discutidos"
                elif "pendientes del cliente" in texto_l or "pendientes cliente" in texto_l: contexto_actual = "Pendientes Cliente"
                elif "pendientes mycloud" in texto_l: contexto_actual = "Pendientes Mycloud"
                elif texto and contexto_actual in ["Objetivo", "Puntos Discutidos"]:
                    datos[contexto_actual] = (datos[contexto_actual] + "\n" + texto).strip()

            elif isinstance(bloque, Table):
                filas_texto = []
                for row in bloque.rows[1:]:
                    contenido = [c.text.strip() for c in row.cells if c.text.strip()]
                    if contenido: filas_texto.append(", ".join(contenido))
                
                if contexto_actual in ["Asistentes", "Pendientes Cliente", "Pendientes Mycloud"]:
                    datos[contexto_actual] = "\n".join(filas_texto)
    except: pass
    return datos

# --- 4. FUNCIONES DE GENERACIÓN ---
def rellenar_tabla(tabla, texto_lineas, columnas):
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    for linea in texto_lineas.split('\n'):
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        for i in range(min(len(partes), columnas)):
            nueva_fila[i].text = partes[i].strip()

def procesar_word(template_path, datos_usuario):
    doc = Document(template_path)
    # Reemplazo de texto simple
    for p in doc.paragraphs:
        if "Fecha:" in p.text: p.text = f"Fecha: {datos_usuario['Fecha']}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {datos_usuario['Objetivo']}"
        if "Puntos discutidos:" in p.text: p.text = f"Puntos discutidos:\n{datos_usuario['Puntos Discutidos']}"
    
    # Reemplazo en tablas
    for tabla in doc.tables:
        header = tabla.cell(0,0).text.lower()
        if "nombre" in header: rellenar_tabla(tabla, datos_usuario["Asistentes"], 2)
        elif "pendientes del cliente" in header: rellenar_tabla(tabla, datos_usuario["Pendientes Cliente"], 3)
        elif "pendientes mycloud" in header: rellenar_tabla(tabla, datos_usuario["Pendientes Mycloud"], 3)
    
    return doc

# --- 5. INTERFAZ PRINCIPAL ---
st.title("🚀 Generador CRM Profesional")

archivo_ref = st.file_uploader("📂 Sube archivo de referencia (opcional):", type=["docx"])
datos_auto = extraer_informacion(archivo_ref)

with st.form(key=f"form_{opcion}"):
    st.subheader(f"Editando: {opcion}")
    c1, c2 = st.columns(2)
    
    with c1:
        fecha = st.text_input("Fecha", value=datos_auto["Fecha"])
        asistentes = st.text_area("Asistentes", value=datos_auto["Asistentes"], height=150)
        p_cliente = st.text_area("Pendientes Cliente", value=datos_auto["Pendientes Cliente"], height=150)
    with c2:
        objetivo = st.text_area("Objetivo", value=datos_auto["Objetivo"], height=68)
        puntos = st.text_area("Puntos Discutidos", value=datos_auto["Puntos Discutidos"], height=150)
        p_mycloud = st.text_area("Pendientes Mycloud", value=datos_auto["Pendientes Mycloud"], height=150)
    
    submit = st.form_submit_button("🔨 GENERAR VISTA PREVIA")

# El botón de descarga debe estar FUERA del formulario para funcionar correctamente
if submit:
    datos_finales = {
        "Fecha": fecha, "Objetivo": objetivo, "Asistentes": asistentes,
        "Puntos Discutidos": puntos, "Pendientes Cliente": p_cliente, "Pendientes Mycloud": p_mycloud
    }
    
    try:
        doc_generado = procesar_word(TEMPLATES[opcion], datos_finales)
        buffer = io.BytesIO()
        doc_generado.save(buffer)
        buffer.seek(0)
        
        st.success("✅ Documento procesado correctamente.")
        st.download_button(
            label="📥 DESCARGAR ARCHIVO WORD",
            data=buffer,
            file_name=f"{opcion.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Error al generar el archivo: {e}. Asegúrate de que '{TEMPLATES[opcion]}' esté en tu repositorio.")
