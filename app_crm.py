import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
import io
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

# --- 1. PLANTILLAS ---
TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# --- 2. BARRA LATERAL (LOGO Y SELECTOR) ---
with st.sidebar:
    st.header("Identidad Visual")
    # Restaurado el cargador de logo
    logo_web = st.file_uploader("1. Sube el logo para la WEB:", type=["png", "jpg", "jpeg"], key="logo_web")
    
    if logo_web:
        st.image(logo_web, use_container_width=True)
    
    st.divider()
    st.header("2. Configuración")
    opcion = st.selectbox("Selecciona la Plantilla:", list(TEMPLATES.keys()))

# --- 3. FUNCIONES DE APOYO ---
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
                if "fecha:" in texto_l: datos["Fecha"] = texto.split(":", 1)[1].strip()
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
                filas = [", ".join([c.text.strip() for c in r.cells if c.text.strip()]) for r in bloque.rows[1:]]
                if contexto_actual in ["Asistentes", "Pendientes Cliente", "Pendientes Mycloud"]:
                    datos[contexto_actual] = "\n".join([f for f in filas if f])
    except: pass
    return datos

# --- 4. LÓGICA DE GENERACIÓN CON NUMERACIÓN ---
def rellenar_tabla(tabla, texto_lineas, columnas):
    while len(tabla.rows) > 1: tabla._tbl.remove(tabla.rows[-1]._tr)
    for linea in texto_lineas.split('\n'):
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        for i in range(min(len(partes), columnas)):
            nueva_fila[i].text = partes[i].strip()

def procesar_word(template_path, datos):
    doc = Document(template_path)
    
    # 1. Numerar los puntos discutidos
    lineas_puntos = [l.strip() for l in datos['Puntos Discutidos'].split('\n') if l.strip()]
    puntos_numerados = "\n".join([f"{i+1}. {linea}" for i, linea in enumerate(lineas_puntos)])

    # 2. Reemplazo en párrafos
    for p in doc.paragraphs:
        if "Fecha:" in p.text: p.text = f"Fecha: {datos['Fecha']}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {datos['Objetivo']}"
        if "Puntos discutidos:" in p.text:
            p.text = "Puntos discutidos:"
            # Insertamos el texto numerado justo después
            doc.add_paragraph(puntos_numerados)
    
    # 3. Reemplazo en tablas
    for tabla in doc.tables:
        header = tabla.cell(0,0).text.lower()
        if "nombre" in header: rellenar_tabla(tabla, datos["Asistentes"], 2)
        elif "pendientes del cliente" in header: rellenar_tabla(tabla, datos["Pendientes Cliente"], 3)
        elif "pendientes mycloud" in header: rellenar_tabla(tabla, datos["Pendientes Mycloud"], 3)
    
    return doc

# --- 5. INTERFAZ ---
st.title("🚀 Generador CRM Profesional")

archivo_ref = st.file_uploader("📂 Sube archivo de referencia (opcional):", type=["docx"])
datos_auto = extraer_informacion(archivo_ref)

with st.form(key="crm_form"):
    c1, c2 = st.columns(2)
    with c1:
        fecha = st.text_input("Fecha", value=datos_auto["Fecha"])
        asistentes = st.text_area("Asistentes", value=datos_auto["Asistentes"], height=150)
        p_cliente = st.text_area("Pendientes Cliente", value=datos_auto["Pendientes Cliente"], height=150)
    with c2:
        objetivo = st.text_area("Objetivo", value=datos_auto["Objetivo"], height=68)
        puntos = st.text_area("Puntos Discutidos", value=datos_auto["Puntos Discutidos"], height=150, help="Escribe cada punto en una línea nueva.")
        p_mycloud = st.text_area("Pendientes Mycloud", value=datos_auto["Pendientes Mycloud"], height=150)
    
    submit = st.form_submit_button("🔨 GENERAR DOCUMENTO")

if submit:
    datos_f = {"Fecha": fecha, "Objetivo": objetivo, "Asistentes": asistentes, "Puntos Discutidos": puntos, "Pendientes Cliente": p_cliente, "Pendientes Mycloud": p_mycloud}
    doc_final = procesar_word(TEMPLATES[opcion], datos_f)
    
    buffer = io.BytesIO()
    doc_final.save(buffer)
    buffer.seek(0)
    
    st.success("✅ ¡Documento listo!")
    st.download_button(
        label="📥 DESCARGAR ARCHIVO WORD",
        data=buffer,
        file_name=f"{opcion.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
