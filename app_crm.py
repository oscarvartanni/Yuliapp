import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
import io
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="CRM Generator Pro 2.0", layout="wide")

# --- 1. DEFINICIÓN DE PLANTILLAS ---
TEMPLATES = {
    "M100 Minuta": "M100_CRM_Minuta v2 (2).docx",
    "M102 Gap Analysis": "M102_CRM_Gap_Analysis V2 (3).docx",
    "M101 Escenarios": "M101_CRM_Lista_de_escenarios_para_CRPUAT V2 (1).docx"
}

# --- 2. FUNCIONES DE APOYO ---
def aplicar_poppins(run, size=11):
    run.font.name = 'Poppins'
    run.font.size = Pt(size)

def iterar_bloques(parent):
    from docx.document import Document as _Document
    parent_elm = parent.element.body if isinstance(parent, _Document) else parent._tc
    for child in parent_elm.iterchildren():
        if child.tag.endswith('p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('tbl'):
            yield Table(child, parent)

# FIX 1: Helper que lee TODA la fila de encabezado (no solo cell(0,0))
def get_table_header(tabla):
    """Devuelve el texto completo de la primera fila, en minúsculas, separado por |"""
    if not tabla.rows:
        return ""
    return " | ".join(c.text.strip().lower() for c in tabla.rows[0].cells)

# --- 3. EXTRACCIÓN INTELIGENTE 2.0 ---
# FIX 2: extraer_informacion ahora también extrae campos específicos del Gap Analysis
def extraer_informacion(archivo_subido):
    datos = {k: "" for k in [
        "Fecha", "Objetivo", "Asistentes",
        "Puntos Discutidos", "Pendientes Cliente", "Pendientes Mycloud",
        # Campos Gap Analysis
        "Modulos", "Pendientes_Gap", "Custom", "WebServices", "Workflows"
    ]}
    if not archivo_subido:
        return datos

    try:
        doc = Document(archivo_subido)
        contexto = None
        prev_was_objetivo = False

        for bloque in iterar_bloques(doc):
            if isinstance(bloque, Paragraph):
                txt = bloque.text.strip()
                txt_l = txt.lower()

                # Extracción de Fecha
                if "fecha:" in txt_l and not txt_l.startswith("entrega"):
                    datos["Fecha"] = txt.split(":", 1)[1].strip()

                # Detección de Secciones (Cambio de contexto)
                elif any(x in txt_l for x in ["objetivo:", "alcance:"]):
                    contexto = "Objetivo"
                    res = txt.split(":", 1)
                    if len(res) > 1 and res[1].strip():
                        datos["Objetivo"] = res[1].strip()
                        prev_was_objetivo = False
                    else:
                        prev_was_objetivo = True
                elif "asistentes:" in txt_l:
                    contexto = "Asistentes"
                    prev_was_objetivo = False
                elif "puntos discutidos:" in txt_l:
                    contexto = "Puntos Discutidos"
                    prev_was_objetivo = False
                elif "pendientes del cliente" in txt_l or "pendientes cliente" in txt_l:
                    contexto = "Pendientes Cliente"
                    prev_was_objetivo = False
                elif "pendientes mycloud" in txt_l:
                    contexto = "Pendientes Mycloud"
                    prev_was_objetivo = False
                # Secciones Gap Analysis
                elif "ajustes en módulos" in txt_l or "módulos y funcionalidades" in txt_l:
                    contexto = "Modulos"
                    prev_was_objetivo = False
                elif "entrega" in txt_l and "pendientes" in txt_l:
                    contexto = "Pendientes_Gap"
                    prev_was_objetivo = False
                elif "custom functions" in txt_l:
                    contexto = "Custom"
                    prev_was_objetivo = False
                elif "web services" in txt_l:
                    contexto = "WebServices"
                    prev_was_objetivo = False
                elif txt_l.startswith("workflows"):
                    contexto = "Workflows"
                    prev_was_objetivo = False

                # Captura de párrafo suelto de Objetivo (cuando viene en el siguiente párrafo)
                elif prev_was_objetivo and txt:
                    datos["Objetivo"] = txt
                    prev_was_objetivo = False

                # Captura de contenido de párrafo (solo para secciones de texto)
                elif txt and contexto in ("Puntos Discutidos",):
                    datos[contexto] = (datos[contexto] + "\n" + txt).strip()

            elif isinstance(bloque, Table) and contexto:
                # Para secciones que se leen de tabla
                h = get_table_header(bloque)

                if contexto == "Asistentes" and "nombre" in h and "puesto" in h and "firma" not in h:
                    filas = [
                        ", ".join(c.text.strip() for c in r.cells if c.text.strip())
                        for r in bloque.rows[1:]
                    ]
                    datos["Asistentes"] = "\n".join(f for f in filas if f)

                elif contexto in ("Pendientes Cliente", "Pendientes Mycloud"):
                    filas = [
                        ", ".join(c.text.strip() for c in r.cells)
                        for r in bloque.rows[1:]
                        if any(c.text.strip() for c in r.cells)
                    ]
                    datos[contexto] = "\n".join(f for f in filas if f)

                elif contexto == "Modulos" and ("nombre del módulo" in h or ("módulo" in h and "estatus" in h)):
                    filas = [
                        ", ".join(c.text.strip() for c in r.cells)
                        for r in bloque.rows[1:]
                        if r.cells[0].text.strip()
                    ]
                    datos["Modulos"] = "\n".join(filas)

                elif contexto == "Pendientes_Gap" and "responsable" in h:
                    filas = [
                        ", ".join(c.text.strip() for c in r.cells)
                        for r in bloque.rows[1:]
                        if r.cells[0].text.strip()
                    ]
                    datos["Pendientes_Gap"] = "\n".join(filas)

                elif contexto == "Custom" and "descripción" in h and "tipo" not in h:
                    filas = [
                        ", ".join(c.text.strip() for c in r.cells)
                        for r in bloque.rows[1:]
                        if r.cells[0].text.strip()
                    ]
                    datos["Custom"] = "\n".join(filas)

                elif contexto == "WebServices" and "tipo" in h and "parámetros" in h:
                    filas = [
                        ", ".join(c.text.strip() for c in r.cells)
                        for r in bloque.rows[1:]
                        if r.cells[0].text.strip()
                    ]
                    datos["WebServices"] = "\n".join(filas)

                elif contexto == "Workflows" and ("cuándo" in h or "qué registros" in h):
                    filas = [
                        ", ".join(c.text.strip() for c in r.cells)
                        for r in bloque.rows[1:]
                        if r.cells[0].text.strip()
                    ]
                    datos["Workflows"] = "\n".join(filas)

    except Exception as e:
        st.warning(f"Error al leer el documento: {e}")
    return datos

# --- 4. GENERACIÓN DE DOCUMENTOS ---
def rellenar_tabla(tabla, texto_lineas, columnas):
    while len(tabla.rows) > 1:
        tabla._tbl.remove(tabla.rows[-1]._tr)
    for linea in texto_lineas.split('\n'):
        if not linea.strip():
            continue
        nueva_fila = tabla.add_row().cells
        partes = linea.split(',')
        for i in range(min(len(partes), columnas)):
            nueva_fila[i].text = partes[i].strip()
            for p in nueva_fila[i].paragraphs:
                for run in p.runs:
                    aplicar_poppins(run)

def procesar_word(template_path, datos, es_gap=False):
    doc = Document(template_path)
    for p in doc.paragraphs:
        if "Fecha:" in p.text:
            p.text = "Fecha: "
            aplicar_poppins(p.add_run(datos.get('Fecha', '')))
        elif any(x in p.text for x in ["Objetivo:", "Alcance:"]):
            p.text = "Objetivo: " if not es_gap else "Objetivo : "
            aplicar_poppins(p.add_run(datos.get('Objetivo', '')))
        elif "Puntos discutidos:" in p.text and not es_gap:
            p.text = "Puntos discutidos:"
            for i, linea in enumerate(datos.get('Puntos Discutidos', '').split('\n'), 1):
                if linea.strip():
                    np = p.insert_paragraph_before(f"{i}. {linea.strip()}")
                    aplicar_poppins(np.runs[0] if np.runs else np.add_run())

    # FIX 3: Usar get_table_header() en lugar de solo cell(0,0)
    for tabla in doc.tables:
        h = get_table_header(tabla)

        if "nombre" in h and "puesto" in h and "firma" not in h:
            rellenar_tabla(tabla, datos.get("Asistentes", ""), 2)
        elif "pendientes del cliente" in h:
            rellenar_tabla(tabla, datos.get("Pendientes Cliente", ""), 3)
        elif "pendientes mycloud" in h:
            rellenar_tabla(tabla, datos.get("Pendientes Mycloud", ""), 3)
        elif "nombre del módulo" in h or ("módulo" in h and "estatus" in h):
            rellenar_tabla(tabla, datos.get("Modulos", ""), 4)
        elif "pendientes" in h and "responsable" in h:
            rellenar_tabla(tabla, datos.get("Pendientes_Gap", ""), 3)
        elif h.startswith("ítem") and "descripción" in h and "tipo" not in h and "módulo" not in h:
            rellenar_tabla(tabla, datos.get("Custom", ""), 2)
        elif "tipo" in h and "parámetros" in h:
            rellenar_tabla(tabla, datos.get("WebServices", ""), 4)
        elif "cuándo" in h or "qué registros" in h:
            rellenar_tabla(tabla, datos.get("Workflows", ""), 5)
        # Tablas de Revisión/Aprobación (firma) → se dejan intactas

    return doc

# --- 5. INTERFAZ DE USUARIO ---
with st.sidebar:
    st.header("🎨 Identidad Visual")
    logo_web = st.file_uploader("Cambiar logo:", type=["png", "jpg", "jpeg"])
    if logo_web:
        st.image(logo_web, use_container_width=True)
    elif os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    st.divider()
    opcion = st.selectbox("Selecciona Plantilla:", list(TEMPLATES.keys()))

st.title("🚀 Generador CRM Profesional v2.0")

archivo_ref = st.file_uploader("📂 Sube la minuta anterior para auto-rellenar:", type=["docx"])
datos_auto = extraer_informacion(archivo_ref)

with st.form(key="main_form"):
    c1, c2 = st.columns(2)
    with c1:
        fecha = st.text_input("Fecha", value=datos_auto["Fecha"])
        asistentes = st.text_area("Asistentes (Nombre, Cargo)", value=datos_auto["Asistentes"], height=150)
        objetivo = st.text_area("Objetivo / Alcance", value=datos_auto["Objetivo"], height=100)
    with c2:
        if opcion == "M102 Gap Analysis":
            # FIX 4: Pasar datos_auto como value para que el auto-relleno funcione en Gap también
            modulos = st.text_area("Módulos (Item, Nombre, Desc, Estatus)", value=datos_auto["Modulos"])
            pend_gap = st.text_area("Pendientes/Entrega (Tarea, Resp, Fecha)", value=datos_auto["Pendientes_Gap"])
            custom = st.text_area("Custom Functions (Item, Desc)", value=datos_auto["Custom"])
            ws = st.text_area("Web Services (Item, Nombre, Tipo, Param)", value=datos_auto["WebServices"])
            wf = st.text_area("Workflows (Item, Módulo, Cuándo, Qué, Acciones)", value=datos_auto["Workflows"])
        else:
            puntos = st.text_area("Puntos Discutidos", value=datos_auto["Puntos Discutidos"], height=150)
            p_cli = st.text_area("Pendientes Cliente", value=datos_auto["Pendientes Cliente"], height=100)
            p_my = st.text_area("Pendientes Mycloud", value=datos_auto["Pendientes Mycloud"], height=100)

    btn = st.form_submit_button("🔨 GENERAR DOCUMENTO")

if btn:
    es_gap = (opcion == "M102 Gap Analysis")
    final_dict = {
        "Fecha": fecha,
        "Asistentes": asistentes,
        "Objetivo": objetivo,
        "Puntos Discutidos": puntos if not es_gap else "",
        "Pendientes Cliente": p_cli if not es_gap else "",
        "Pendientes Mycloud": p_my if not es_gap else "",
        "Modulos": modulos if es_gap else "",
        "Pendientes_Gap": pend_gap if es_gap else "",
        "Custom": custom if es_gap else "",
        "WebServices": ws if es_gap else "",
        "Workflows": wf if es_gap else ""
    }
    try:
        resultado = procesar_word(TEMPLATES[opcion], final_dict, es_gap=es_gap)
        buf = io.BytesIO()
        resultado.save(buf)
        buf.seek(0)
        st.success("✅ ¡Documento generado!")
        st.download_button("📥 Descargar Archivo Word", buf, f"{opcion}.docx")
    except Exception as e:
        st.error(f"Error al generar: {e}")
