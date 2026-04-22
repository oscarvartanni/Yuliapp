import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
import io
import os

# --- FUNCIÓN CRÍTICA: ITERAR EN ORDEN (Párrafos y Tablas) ---
def iterar_bloques(parent):
    """Itera por párrafos y tablas en el orden exacto en que aparecen."""
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
                
                # Detectar cambios de sección por encabezado
                if "fecha:" in texto_l:
                    datos["Fecha"] = texto.split(":", 1)[1].strip()
                    contexto_actual = "Fecha"
                elif "asistentes:" in texto_l:
                    contexto_actual = "Asistentes"
                elif "objetivo:" in texto_l:
                    datos["Objetivo"] = texto.split(":", 1)[1].strip()
                    contexto_actual = "Objetivo"
                elif "puntos discutidos:" in texto_l:
                    contexto_actual = "Puntos Discutidos"
                elif "pendientes del cliente" in texto_l or "pendientes cliente" in texto_l:
                    contexto_actual = "Pendientes Cliente"
                elif "pendientes mycloud" in texto_l:
                    contexto_actual = "Pendientes Mycloud"
                elif texto and contexto_actual in ["Objetivo", "Puntos Discutidos"]:
                    # Acumular texto multilínea (como los puntos discutidos)
                    if datos[contexto_actual]: datos[contexto_actual] += "\n" + texto
                    else: datos[contexto_actual] = texto

            elif isinstance(bloque, Table):
                # Si encontramos una tabla, procesarla según el contexto del párrafo anterior
                filas_texto = []
                for row in bloque.rows[1:]: # Saltar encabezado de la tabla
                    contenido = [c.text.strip() for c in row.cells if c.text.strip()]
                    if contenido: filas_texto.append(", ".join(contenido))
                
                if contexto_actual in ["Asistentes", "Pendientes Cliente", "Pendientes Mycloud"]:
                    datos[contexto_actual] = "\n".join(filas_texto)
                    contexto_actual = None # Resetear contexto tras procesar su tabla
                    
    except Exception as e:
        st.error(f"Error leyendo el archivo: {e}")
    return datos

# --- RESTO DEL CÓDIGO (INTERFAZ Y GENERACIÓN) ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

# (Mantener definición de TEMPLATES y CONFIG_DETALLADA igual que antes)
TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

with st.sidebar:
    st.title("Configuración")
    opcion = st.selectbox("Selecciona Plantilla:", list(TEMPLATES.keys()))

st.title("🚀 Generador CRM Profesional")
archivo_ref = st.file_uploader("📂 Sube archivo de referencia (Minuta 22-04):", type=["docx"])
datos_auto = extraer_informacion(archivo_ref)

with st.form(key=f"form_{opcion}"):
    st.subheader(f"Editando: {opcion}")
    col1, col2 = st.columns(2)
    datos_finales = {}
    
    # Mapeo de campos para el formulario
    campos = {
        "Fecha": "Fecha", "Objetivo": "Objetivo", 
        "Asistentes": "Asistentes", "Puntos Discutidos": "Puntos Discutidos",
        "Pendientes Cliente": "Pendientes Cliente", "Pendientes Mycloud": "Pendientes Mycloud"
    }

    for i, (campo, label) in enumerate(campos.items()):
        with (col1 if i % 2 == 0 else col2):
            val_defecto = datos_auto.get(campo, "")
            if campo == "Fecha":
                datos_finales[campo] = st.text_input(label, value=val_defecto)
            else:
                datos_finales[campo] = st.text_area(label, value=val_defecto, height=150)
    
    boton_generar = st.form_submit_button("🔨 GENERAR DOCUMENTO")

# (La lógica de procesar_word y download_button se mantiene igual)
