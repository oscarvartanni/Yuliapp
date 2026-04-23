from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def procesar_word(template_path, datos_usuario):
    doc = Document(template_path)
    
    # Definir estilo de fuente Poppins si está disponible, o aplicarlo directamente
    def aplicar_formato_poppins(run):
        run.font.name = 'Poppins'
        run.font.size = Pt(11)

    for p in doc.paragraphs:
        # Reemplazo de Fecha y Objetivo con Poppins
        if "Fecha:" in p.text:
            p.text = "" # Limpiamos para aplicar formato desde cero
            run = p.add_run(f"Fecha: {datos_usuario['Fecha']}")
            aplicar_formato_poppins(run)
            
        if "Objetivo:" in p.text:
            p.text = ""
            run = p.add_run(f"Objetivo: {datos_usuario['Objetivo']}")
            aplicar_formato_poppins(run)

        # Lógica para PUNTOS DISCUTIDOS NUMERADOS
        if "Puntos discutidos:" in p.text:
            p.text = "Puntos discutidos:"
            # Iterar sobre las líneas de puntos discutidos
            lineas = datos_usuario['Puntos Discutidos'].split('\n')
            for i, linea in enumerate(lineas, 1):
                if linea.strip():
                    # Crear un nuevo párrafo justo después del encabezado
                    nuevo_p = p.insert_paragraph_before(f"{i}. {linea.strip()}")
                    nuevo_p.style = doc.styles['List Number'] # Usa el estilo nativo de Word
                    run = nuevo_p.runs[0]
                    aplicar_formato_poppins(run)
            # Eliminar el párrafo original de marcador si es necesario o dejarlo como título
    
    # Aplicar Poppins también a las TABLAS
    for tabla in doc.tables:
        header = tabla.cell(0,0).text.lower()
        if "nombre" in header: 
            rellenar_tabla(tabla, datos_usuario["Asistentes"], 2)
        elif "pendientes del cliente" in header: 
            rellenar_tabla(tabla, datos_usuario["Pendientes Cliente"], 3)
        elif "pendientes mycloud" in header: 
            rellenar_tabla(tabla, datos_usuario["Pendientes Mycloud"], 3)
        
        # Aplicar fuente a todas las celdas de la tabla
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for run in parrafo.runs:
                        aplicar_formato_poppins(run)
    
    return doc
