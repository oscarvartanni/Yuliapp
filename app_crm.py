import streamlit as st
from docx import Document
from docx.shared import Inches
import io
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Generator Pro", layout="wide")

# Mapeo exacto de plantillas
TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# Configuración de campos y sus ejemplos (Placeholders)
CONFIG_DETALLADA = {
    "M100 Minuta": {
        "Fecha": "21 de Abril 2026",
        "Objetivo": "Definir alcances del módulo de ventas",
        "Asistentes": "Juan Perez, Gerente\nMaria Garcia, Consultora",
        "Puntos Discutidos": "Revisión de tiempos\nValidación de campos",
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
        "Escenarios de Prueba": "1. Login exitoso con credenciales válidas\n2. Intento de acceso con contraseña errónea\n3. Recuperación de cuenta vía email"
    }
}

def extraer_informacion(archivo_subido):
    datos = {"Fecha": "", "Objetivo": ""}
    if archivo_subido:
        try:
            doc = Document(archivo_subido)
            texto_completo = "\n".join([p.text for p in doc.paragraphs])
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells: texto_completo += "\n" + c.text
            
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
    
    # 1. Logo en encabezado
    if logo_img:
        try:
            header = doc.sections[0].header
            p = header.paragraphs[0]
            p.alignment = 1
            p.add_run().add_picture(logo_img, width=Inches(1.2))
        except: pass

    # 2. Reemplazo de Fecha y Objetivo (en párrafos y tablas)
    for p in doc.paragraphs:
        if "Fecha:" in p.text: p.text = f"Fecha: {datos_usuario.get('Fecha', '')}"
        if "Objetivo:" in p.text: p.text = f"Objetivo: {datos_usuario.get('Objetivo', '')}"

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                if "Fecha:" in celda.text: celda.text = f"Fecha: {datos_usuario.get('Fecha', '')}"
                if "Objetivo:" in celda.text: celda.text = f"Objetivo: {datos_usuario.get('Objetivo', '')}"

        # 3. Lógica de tablas dinámicas por contenido
        cabecera = tabla.cell(0,0).text.lower()
        if "nombre" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Asistentes", ""), 2)
        elif "pendientes del cliente" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Pendientes Cliente", ""), 3)
        elif "pendientes mycloud" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Pendientes Mycloud", ""), 3)
        elif "módulo" in cabecera:
            rellenar_tabla(tabla, datos_usuario.get("Módulos", ""), 4)

    # 4. Escenarios de Prueba o Puntos Discutidos (Listas)
    # Buscamos campos de área de texto grande para ponerlos como lista
    lista_texto = datos_usuario.get("Escenarios de Prueba") or datos_usuario.get("Puntos Discutidos")
    if lista_texto:
        items = [i.strip() for i in lista_texto.split('\n') if i.strip()]
        idx = 0
        for p in doc.paragraphs:
            if re.match(r"^\d+\.", p.text.strip()) and idx < len(items):
                p.text = f"{idx+1}. {items[idx]}"
                idx += 1
                
    return doc

# --- INTERFAZ PRINCIPAL ---
st.title("🚀 CRM Document Builder")

# Sidebar para selección principal
with st.sidebar:
    st.header("1. Configuración")
    # Al cambiar esta opción, el formulario central cambiará automáticamente
    opcion = st.selectbox("Selecciona la Plantilla:", list(TEMPLATES.keys()), key="main_selector")
    logo = st.file_uploader("Subir Logo:", type=["png", "jpg"])
    st.divider()
    st.info("💡 RECUERDA: En tablas, usa una coma (,) para separar columnas.")

# Área central
archivo_ref = st.file_uploader("2. Sube archivo de referencia (opcional):", type=["docx"])
datos_extraidos = extraer_informacion(archivo_ref)

# FORMULARIO DINÁMICO
# La clave 'form_dinamico_' + opcion obliga a Streamlit a refrescar los campos
with st.form(key=f"form_dinamico_{opcion}"):
    st.subheader(f"Campos para: {opcion}")
    
    config_actual = CONFIG_DETALLADA[opcion]
    datos_finales = {}
    
    col1, col2 = st.columns(2)
    
    for i, (campo, ejemplo) in enumerate(config_actual.items()):
        # Dividir campos en dos columnas
        target_col = col1 if i % 2 == 0 else col2
        with target_col:
            # Pre-llenar solo Fecha y Objetivo si vienen del archivo
            valor_inicial = datos_extraidos.get(campo, "") if campo in ["Fecha", "Objetivo"] else ""
            
            if campo == "Fecha":
                datos_finales[campo] = st.text_input(campo, value=valor_inicial, placeholder=ejemplo)
            else:
                datos_finales[campo] = st.text_area(campo, value=valor_inicial, placeholder=f"Ej: {ejemplo}", height=150)
                
    btn_generar = st.form_submit_button("🔨 GENERAR ARCHIVO")

# Acción después de enviar el formulario
if btn_generar:
    with st.spinner("Creando documento..."):
        try:
            doc_final = procesar_word(opcion, datos_finales, logo)
            buffer = io.BytesIO()
            doc_final.save(buffer)
            
            st.success(f"✅ ¡{opcion} generado correctamente!")
            st.download_button(
                label="📥 DESCARGAR AHORA",
                data=buffer.getvalue(),
                file_name=f"{opcion.replace(' ','_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error al generar: {e}")
