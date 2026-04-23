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

# --- 3. FUNCIONES DE EXTRACCIÓN Y FORMATO ---
def aplicar_poppins(run, size=11):
    """Aplica la fuente Poppins a un fragmento de texto."""
    run.font.name = 'Poppins'
    run.font.size = Pt(size)

def iterar_bloques(parent):
    from docx.document import Document as _Document
    parent_elm = parent.element.body if isinstance(parent, _Document) else parent._tc
    for child in parent_elm.iterchildren():
        if child.tag.endswith('p'): yield Paragraph(child, parent)
        elif child.tag.endswith('tbl'): yield Table(child, parent)

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
                if "fecha:" in texto_l: datos["Fecha"] = texto.split(":", 1)[1].strip()
                elif "asistentes:" in texto_l: contexto_actual = "Asistentes"
                elif "objetivo:" in texto_l:
                    datos["Objetivo"] = texto.split(":", 1)[1].strip()
                    contexto_actual = "Objetivo"
                elif "puntos discutidos:" in texto_l: contexto_actual = "Puntos Discutidos"
                elif "pendientes" in texto_l: 
                    contexto_actual = "Pendientes Cliente" if "cliente" in texto_l else "Pendientes Mycloud"
                elif texto and contexto_actual in ["Objetivo", "Puntos Discutidos"]:
                    datos[contexto_actual] = (datos[contexto_actual] + "\n" + texto).strip()
            elif isinstance(bloque, Table) and contexto_actual:
                filas = [", ".join(c.text.strip() for c in r.cells if c.text.strip()) for r in bloque.rows[1:]]
                datos[contexto_actual] = "\n".join(f for f in filas if f)
    except: pass
    return datos

# --- 4. GENERACIÓN DEL DOCUMENTO ---
def procesar_word(template_path, datos):
    doc = Document(template_path)
    
    for p in doc.paragraphs:
        # Reemplazo con fuente Poppins
        if "Fecha:" in p.text:
            p.text = "Fecha: "
            aplicar_poppins(p.add_run(datos['Fecha']))
        elif "Objetivo:" in p.text:
            p.text = "Objetivo: "
            aplicar_poppins(p.add_run(datos['Objetivo']))
        elif "Puntos discutidos:" in p.text:
            p.text = "Puntos discutidos:"
            # Numeración consecutiva [cite: 24]
            puntos = datos['Puntos Discutidos'].split('\n')
            for i, punto in enumerate(puntos, 1):
                if punto.strip():
                    nuevo_p = p.insert_paragraph_before(f"{i}. {punto.strip()}")
                    nuevo_p.style = doc.styles['List Number'] # Estilo nativo [cite: 26]
                    aplicar_poppins(nuevo_p.runs[0])

    for tabla in doc.tables:
        # Lógica de llenado de tablas adaptada
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for run in parrafo.runs: aplicar_poppins(run)
    return doc

# --- 5. INTERFAZ ---
st.title("🚀 Yuliaapp: Generador CRM")
archivo_ref = st.file_uploader("📂 Archivo de referencia:", type=["docx"])
datos_auto = extraer_informacion(archivo_ref)

with st.form("crm_form"):
    opcion = st.selectbox("Plantilla:", list(TEMPLATES.keys()))
    col1, col2 = st.columns(2)
    with col1:
        fecha = st.text_input("Fecha", value=datos_auto["Fecha"])
        asistentes = st.text_area("Asistentes (Nombre, Cargo)", value=datos_auto["Asistentes"])
    with col2:
        objetivo = st.text_input("Objetivo", value=datos_auto["Objetivo"])
        puntos = st.text_area("Puntos Discutidos (un punto por línea)", value=datos_auto["Puntos Discutidos"])
    
    if st.form_submit_button("🔨 GENERAR"):
        doc = procesar_word(TEMPLATES[opcion], {"Fecha": fecha, "Objetivo": objetivo, "Puntos Discutidos": puntos})
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button("📥 Descargar", buf.getvalue(), f"{opcion}.docx")
