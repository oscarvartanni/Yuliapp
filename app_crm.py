import streamlit as st
from docx import Document
import io
import re

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="CRM Pro - Generador", layout="wide")
st.title("🚀 Generador de Documentos CRM")

# Archivos base en tu repositorio
TEMPLATES = {
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# Campos que la app preguntará según el template
CAMPOS_CONFIG = {
    "M102 Gap Analysis": ["Fecha", "Objetivo", "Asistentes", "Módulos", "Pendientes"],
    "M100 Minuta": ["Fecha", "Objetivo", "Asistentes", "Puntos discutidos", "Acuerdos"],
    "M101 Escenarios": ["Fecha", "Objetivo", "Escenarios de prueba"]
}

def extraer_texto_base(file):
    doc = Document(file)
    texto = "\n".join([p.text for p in doc.paragraphs])
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto += "\n" + celda.text
    return texto

def generar_documento(template_path, datos_finales):
    doc = Document(template_path)
    
    # 1. Reemplazo en párrafos
    for p in doc.paragraphs:
        for campo, valor in datos_finales.items():
            etiqueta = f"{campo}:"
            if etiqueta in p.text:
                texto_sustituto = "" if valor == "[OMITIDO]" else valor
                p.text = f"{etiqueta} {texto_sustituto}"

    # 2. Lógica especial para tabla de Asistentes
    if "Asistentes" in datos_finales and datos_finales["Asistentes"] != "[OMITIDO]":
        asistentes = [a.strip() for a in datos_finales["Asistentes"].split('\n') if a.strip()]
        for tabla in doc.tables:
            if len(tabla.rows) > 0 and "Nombre" in tabla.cell(0, 0).text:
                # Limpiar filas de ejemplo
                while len(tabla.rows) > 1:
                    tr = tabla.rows[-1]._tr
                    tabla._tbl.remove(tr)
                # Insertar nuevos
                for asis in asistentes:
                    nueva_fila = tabla.add_row().cells
                    partes = asis.split(',')
                    nueva_fila[0].text = partes[0].strip()
                    if len(partes) > 1:
                        nueva_fila[1].text = partes[1].strip()
                break
    return doc

# --- INTERFAZ DE USUARIO ---
opcion = st.selectbox("1. Selecciona tu plantilla:", list(TEMPLATES.keys()))
archivo_subido = st.file_uploader("2. Sube el archivo con la información nueva:", type=["docx"])

if archivo_subido:
    texto_fuente = extraer_texto_base(archivo_subido)
    datos_recolectados = {}
    
    st.info("Valida la información de cada campo. Marca 'Omitir' si no deseas incluirlo.")
    
    with st.form("formulario_campos"):
        for campo in CAMPOS_CONFIG[opcion]:
            st.markdown(f"### Sección: {campo}")
            col_txt, col_chk = st.columns([4, 1])
            
            # Intento de encontrar el dato automáticamente
            match = re.search(rf"{campo}:\s*(.*)", texto_fuente, re.I)
            sugerencia = match.group(1).strip() if match else ""
            
            with col_txt:
                val_usuario = st.text_area(f"Editar {campo}", value=sugerencia, key=f"input_{campo}")
            with col_chk:
                marcado_omitir = st.checkbox("Omitir", key=f"omitir_{campo}")
            
            datos_recolectados[campo] = "[OMITIDO]" if marcado_omitir else val_usuario
            st.write("---")
            
        boton_preparar = st.form_submit_button("🔨 Generar Archivo")

    # Botón de descarga (debe ir fuera del formulario)
    if boton_preparar:
        documento_final = generar_documento(TEMPLATES[opcion], datos_recolectados)
        
        buffer = io.BytesIO()
        documento_final.save(buffer)
        buffer.seek(0)
        
        st.success("✅ ¡El documento ha sido generado con éxito!")
        st.download_button(
            label="📥 Descargar Word Final",
            data=buffer,
            file_name=f"Resultado_{opcion.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
