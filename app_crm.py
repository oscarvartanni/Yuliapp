import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# Definimos los campos y sus ejemplos
# Formato: "Nombre del Campo": "Ejemplo de cómo escribirlo"
CONFIG_DETALLADA = {
    "M100 Minuta": {
        "Fecha": "21 de Abril 2026",
        "Objetivo": "Definir alcances del módulo de ventas",
        "Asistentes": "Juan Perez, Gerente\nMaria Garcia, Consultora",
        "Puntos Discutidos": "Revisión de tiempos\nValidación de campos\nAcuerdos de entrega",
        "Pendientes Cliente": "Enviar accesos, Juan Perez, 25/04/2026",
        "Pendientes Mycloud": "Configurar portal, Oscar V., 30/04/2026"
    },
    "M102 Gap Analysis": {
        "Fecha": "21 de Abril 2026",
        "Objetivo": "Análisis de brechas técnica vs funcional",
        "Asistentes": "Luis Pascal, Dirección\nAlejandro Chávez, Implementación",
        "Módulos": "Ventas, Prospectos, Nuevo Campo, En proceso",
        "Pendientes General": "Revisión de API, Dev Team, 05/05/2026"
    },
    "M101 Escenarios": {
        "Fecha": "21 de Abril 2026",
        "Objetivo": "Pruebas de aceptación de usuario (UAT)",
        "Escenarios de Prueba": "Login exitoso\nFallo de contraseña\nRecuperación de cuenta"
    }
}

def extraer_informacion(archivo_subido):
    datos = {"Fecha": "", "Objetivo": ""}
    if archivo_subido:
        try:
            doc = Document(archivo_subido)
            texto_completo = ""
            for p in doc.paragraphs: texto_completo += p.text + "\n"
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells: texto_completo += c.text + " "
            
            f_match = re.search(r"Fecha[:\s]+([^\n]*)", texto_completo, re.IGNORECASE)
            o_match = re.search(r"Objetivo[:\s]+([^\n]*)", texto_completo, re.IGNORECASE)
            if f_match: datos["Fecha"] = f_match.group(1).strip()
            if o_match: datos["Objetivo"] = o_match.group(1).strip()
        except: pass
    return datos

def rellenar_tabla(tabla, texto_lineas, columnas):
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    for linea in texto_lineas.split('\n'):
        if not linea.strip(): continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        for i in range(min(len(partes), columnas)):
            nueva_fila[i].text = partes[i].strip()

def procesar_word(template_name, datos_usuario, logo_img=None):
    doc = Document(TEMPLATES[template_name])
    if logo_img:
        try: doc.sections[0].header.paragraphs[0].add_run().add_picture(logo_img, width=Inches(1.2))
        except: pass

    # Reemplazo en texto (Párrafos y Tablas)
    for p in doc.paragraphs:
        if "Fecha:" in p.text: p.text = f"Fecha: {datos_usuario.get('Fecha', '')}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {datos_usuario.get('Objetivo', '')}"

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if "Fecha:" in celda.text: celda.text = f"Fecha: {datos_usuario.get('Fecha', '')}"
                if "Objetivo:" in celda.text: celda.text = f"Objetivo: {datos_usuario.get('Objetivo', '')}"

        # Lógica de llenado de tablas dinámicas
        cabecera = tabla.cell(0,0).text.lower()
        if "nombre" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Asistentes", ""), 2)
        elif "pendientes del cliente" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Pendientes Cliente", ""), 3)
        elif "pendientes mycloud" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Pendientes Mycloud", ""), 3)
        elif "módulo" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Módulos", ""), 4)

    # Puntos numerados (solo para Minuta)
    puntos = [p.strip() for p in datos_usuario.get("Puntos Discutidos", "").split('\n') if p.strip()]
    if puntos:
        idx = 0
        for p in doc.paragraphs:
            if re.match(r"^\d+\.", p.text.strip()) and idx < len(puntos):
                p.text = f"{idx+1}. {puntos[idx]}"
                idx += 1
    return doc

# --- INTERFAZ ---
st.title("🚀 Generador de Documentos CRM")

with st.sidebar:
    st.header("Configuración")
    opcion = st.selectbox("Selecciona Plantilla:", list(TEMPLATES.keys()))
    logo = st.file_uploader("Subir Logo (Opcional):", type=["png", "jpg"])
    st.divider()
    st.info("💡 Tip: En las tablas, separa los datos por comas.")

archivo_ref = st.file_uploader("Subir archivo de referencia (opcional):", type=["docx"])
datos_extraidos = extraer_informacion(archivo_ref)

# Formulario con Key Dinámica para refrescar campos
with st.form(key=f"form_dinamico_{opcion}"):
    st.subheader(f"Campos para {opcion}")
    datos_finales = {}
    config_actual = CONFIG_DETALLADA[opcion]
    
    col1, col2 = st.columns(2)
    for i, (campo, ejemplo) in enumerate(config_actual.items()):
        target_col = col1 if i % 2 == 0 else col2
        with target_col:
            # Si el campo es Fecha u Objetivo, usamos lo extraído del archivo si existe
            valor_default = datos_extraidos.get(campo, "") if campo in ["Fecha", "Objetivo"] else ""
            
            if campo == "Fecha":
                datos_finales[campo] = st.text_input(
                    label=campo, 
                    value=valor_default, 
                    placeholder=f"Ejemplo: {ejemplo}"
                )
            else:
                datos_finales[campo] = st.text_area(
                    label=campo, 
                    value=valor_default, 
                    placeholder=f"Escribe así: {ejemplo}",
                    height=120
                )
    
    submit = st.form_submit_button("🔨 Generar Documento")

if submit:
    with st.spinner("Procesando..."):
        doc_final = procesar_word(opcion, datos_finales, logo)
        buffer = io.BytesIO()
        doc_final.save(buffer)
        
        st.success("✅ ¡Documento generado!")
        st.download_button(
            label="📥 Descargar Word",
            data=buffer.getvalue(),
            file_name=f"{opcion.replace(' ','_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
