import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Document Generator v3", layout="wide")
st.title("🚀 Generador CRM: Edición Profesional")

TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

def limpiar_y_llenar_tabla(tabla, datos_lineas, columnas_esperadas=2):
    """Borra filas de ejemplo y llena con datos nuevos."""
    # Eliminar todas las filas excepto la cabecera (fila 0)
    while len(tabla.rows) > 1:
        row = tabla.rows[-1]
        tabla._tbl.remove(row._tr)
    
    for linea in datos_lineas:
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        # Dividir por coma
        partes = linea.split(',')
        for i in range(min(len(partes), columnas_esperadas)):
            nueva_fila[i].text = partes[i].strip()

def generar_word_final(template_path, data_dict, logo=None):
    doc = Document(template_path)
    
    # 1. LOGO
    if logo:
        try:
            header = doc.sections[0].header
            p = header.paragraphs[0]
            p.alignment = 1
            p.add_run().add_picture(logo, width=Inches(1.2))
        except: pass

    # 2. REEMPLAZO DE TEXTO (EN TODO EL DOCUMENTO)
    def reemplazar_en_texto(texto_original, clave, valor):
        if valor == "[OMITIDO]": valor = ""
        return texto_original.replace(f"{clave}:", f"{clave}: {valor}")

    # En párrafos normales
    for p in doc.paragraphs:
        for clave in ["Fecha", "Objetivo"]:
            if f"{clave}:" in p.text:
                p.text = reemplazar_en_texto(p.text, clave, data_dict.get(clave, ""))

    # En tablas (algunas plantillas tienen Fecha/Objetivo dentro de celdas)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for clave in ["Fecha", "Objetivo"]:
                    if f"{clave}:" in celda.text:
                        celda.text = reemplazar_en_texto(celda.text, clave, data_dict.get(clave, ""))

    # 3. LLENADO DE TABLAS ESPECÍFICAS
    for tabla in doc.tables:
        # Identificar por la primera celda
        primera_celda = tabla.cell(0, 0).text.lower()
        
        if "nombre" in primera_celda and "puesto" in primera_celda:
            if data_dict.get("Asistentes") != "[OMITIDO]":
                limpiar_y_llenar_tabla(tabla, data_dict.get("Asistentes", "").split('\n'), 2)
        
        elif "pendientes del cliente" in primera_celda:
            if data_dict.get("PC") != "[OMITIDO]":
                limpiar_y_llenar_tabla(tabla, data_dict.get("PC", "").split('\n'), 3)
                
        elif "pendientes mycloud" in primera_celda:
            if data_dict.get("PM") != "[OMITIDO]":
                limpiar_y_llenar_tabla(tabla, data_dict.get("PM", "").split('\n'), 3)

    # 4. PUNTOS DISCUTIDOS (Listas numeradas)
    puntos = [p.strip() for p in data_dict.get("Puntos", "").split('\n') if p.strip()]
    if puntos and data_dict.get("Puntos") != "[OMITIDO]":
        idx = 0
        for p in doc.paragraphs:
            # Busca párrafos que empiecen con "1.", "2.", etc.
            if re.match(r"^\d+\.", p.text.strip()) and idx < len(puntos):
                p.text = f"{idx + 1}. {puntos[idx]}"
                idx += 1

    return doc

# --- INTERFAZ ---
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/281/281760.png", width=50)
logo_subido = st.sidebar.file_uploader("Opcional: Cargar Logo", type=["png", "jpg"])

opcion = st.selectbox("1. Selecciona Plantilla", list(TEMPLATES.keys()))
archivo_base = st.file_uploader("2. Sube el archivo de referencia", type=["docx"])

if archivo_base:
    # Contenedor para persistir el archivo generado
    if 'file_bytes' not in st.session_state:
        st.session_state.file_bytes = None

    with st.form("main_form"):
        st.markdown("### 📝 Rellenar Información")
        col1, col2 = st.columns(2)
        
        with col1:
            fecha = st.text_input("Fecha (ej: 21 de Abril 2026)")
            objetivo = st.text_area("Objetivo de la sesión")
            puntos = st.text_area("Puntos discutidos (Uno por línea)")
        
        with col2:
            asis = st.text_area("Asistentes (Nombre, Puesto)", help="Ej: Juan Perez, Gerente")
            pc = st.text_area("Pendientes Cliente (Tarea, Responsable, Fecha)", help="Ej: Entregar accesos, Maria, 25/04")
            pm = st.text_area("Pendientes Mycloud (Tarea, Responsable, Fecha)", help="Ej: Configurar CRM, Oscar, 30/04")

        submitted = st.form_submit_button("🔨 GENERAR DOCUMENTO")
        
        if submitted:
            datos = {"Fecha": fecha, "Objetivo": objetivo, "Puntos": puntos, "Asistentes": asis, "PC": pc, "PM": pm}
            doc_final = generar_word_final(TEMPLATES[opcion], datos, logo_subido)
            
            output = io.BytesIO()
            doc_final.save(output)
            st.session_state.file_bytes = output.getvalue()

    # BOTÓN DE DESCARGA (Fuera del formulario para evitar errores)
    if st.session_state.file_bytes:
        st.divider()
        st.balloons()
        st.success("¡El archivo se procesó con éxito!")
        st.download_button(
            label="📥 DESCARGAR ARCHIVO AHORA",
            data=st.session_state.file_bytes,
            file_name=f"CRM_{opcion.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
