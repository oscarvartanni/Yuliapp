import streamlit as st
from docx import Document
import io
import re
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

# --- LOGO EN LA WEB (Automático) ---
if os.path.exists("logo.png"):
    st.sidebar.image("logo.png", use_container_width=True)
else:
    st.sidebar.warning("⚠️ Sube 'logo.png' a GitHub para verlo aquí.")

st.sidebar.title("Configuración")

TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

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

# --- FUNCIÓN DE LECTURA MEJORADA (Para Minuta 22-04.docx) ---
def extraer_informacion(archivo_subido):
    datos = {k: "" for k in ["Fecha", "Objetivo", "Asistentes", "Puntos Discutidos", "Pendientes Cliente", "Pendientes Mycloud"]}
    if not archivo_subido: return datos
    
    try:
        doc = Document(archivo_subido)
        # Unimos párrafos para búsqueda por secciones
        parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        
        current_key = None
        for p in parrafos:
            p_lower = p.lower()
            
            # Detectar encabezados y empezar a capturar
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
            elif ":" in p_lower and any(x in p_lower for x in ["acuerdos", "acciones", "conclusiones", "nota"]):
                current_key = None # Dejar de capturar si llega a otra sección
            elif current_key:
                # Si ya estamos en una sección, añadir el texto acumulado
                if datos[current_key]:
                    datos[current_key] += "\n" + p
                else:
                    datos[current_key] = p
        
        # Leer tablas (para los pendientes que vienen en tabla)
        for t in doc.tables:
            header = t.cell(0,0).text.lower()
            key = None
            if "cliente" in header: key = "Pendientes Cliente"
            elif "mycloud" in header: key = "Pendientes Mycloud"
            
            if key:
                rows = []
                for r in t.rows[1:]: # Saltar encabezado de la tabla
                    content = [c.text.strip() for c in r.cells if c.text.strip()]
                    if content: rows.append(", ".join(content))
                datos[key] = "\n".join(rows)

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
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    for idx, linea in enumerate(texto_lineas.split('\n')):
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        nueva_fila[0].text = str(idx + 1)
        for i in range(min(len(partes), 3)):
            nueva_fila[i+1].text = partes[i].strip()

# --- FUNCIÓN DE GENERACIÓN MEJORADA ---
def procesar_word(template_name, datos_usuario):
    doc = Document(TEMPLATES[template_name])
    
    for p in doc.paragraphs:
        p_text_lower = p.text.lower()
        
        # Reemplazo dinámico de etiquetas de texto
        if "fecha:" in p_text_lower:
            p.text = f"Fecha: {datos_usuario.get('Fecha', '')}"
        elif "objetivo:" in p_text_lower:
            p.text = f"Objetivo: {datos_usuario.get('Objetivo', '')}"
        elif "asistentes:" in p_text_lower:
            p.text = f"Asistentes: {datos_usuario.get('Asistentes', '').replace('\\n', ', ')}"
        elif "puntos discutidos:" in p_text_lower:
            # Aquí colocamos el título y el contenido debajo con saltos de línea
            p.text = f"Puntos discutidos:\n{datos_usuario.get('Puntos Discutidos', '')}"

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

# --- INTERFAZ STREAMLIT ---
st.title("🚀 Generador CRM Profesional")

archivo_ref = st.file_uploader("📂 Sube archivo de referencia (Minuta 22-04):", type=["docx"])
datos_auto = extraer_informacion(archivo_ref)

with st.form(key=f"f_{opcion}"):
    st.subheader(f"Campos para {opcion}")
    config = CONFIG_DETALLADA[opcion]
    datos_finales = {}
    
    col1, col2 = st.columns(2)
    for i, (campo, placeholder) in enumerate(config.items()):
        target_col = col1 if i % 2 == 0 else col2
        with target_col:
            # Usar los datos extraídos si existen, si no, vacío
            val_defecto = datos_auto.get(campo, "")
            
            if campo == "Fecha":
                datos_finales[campo] = st.text_input(campo, value=val_defecto, placeholder=placeholder)
            else:
                datos_finales[campo] = st.text_area(campo, value=val_defecto, placeholder=placeholder, height=150)
    
    enviado = st.form_submit_button("🔨 GENERAR DOCUMENTO")

# El botón de descarga debe estar FUERA del formulario
if enviado:
    with st.spinner("Procesando..."):
        try:
            doc_final = procesar_word(opcion, datos_finales)
            buf = io.BytesIO()
            doc_final.save(buf)
            
            st.success("✅ ¡Documento listo!")
            st.download_button(
                label="📥 Descargar Minuta Word",
                data=buf.getvalue(),
                file_name=f"{opcion.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error al procesar: {e}")
