import io
import streamlit as st
from pptx import Presentation
from pptx.util import Inches

# Configuración de página
st.set_page_config(page_title="Editor PPTX Avanzado", page_icon="📊", layout="wide")

# Inicializar la "memoria" para guardar las plantillas
if 'plantillas_guardadas' not in st.session_state:
    st.session_state['plantillas_guardadas'] = {}

# ==========================================
# FUNCIONES NÚCLEO (PROCESAMIENTO PPTX)
# ==========================================

def extraer_secciones_pptx(file_stream, max_slides=5):
    """Extrae el texto identificando si es Título o Contenido para el Preview"""
    file_stream.seek(0)
    prs = Presentation(file_stream)
    slides_data = []
    
    for i, slide in enumerate(prs.slides):
        if i >= max_slides: break
        slide_info = {"titulo": "", "contenido": [], "layout": slide.slide_layout.name}
        
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.text.strip():
                continue
            
            # Si el shape es el título de la diapositiva
            if shape == slide.shapes.title:
                slide_info["titulo"] = shape.text
            else:
                # Todo lo demás se considera contenido/cuerpo
                slide_info["contenido"].append(shape.text)
                
        slides_data.append(slide_info)
    file_stream.seek(0)
    return slides_data

def procesar_pptx(stream_origen, stream_plantilla_bytes):
    """Aplica la plantilla transfiriendo el texto a los placeholders nativos"""
    stream_origen.seek(0)
    prs_origen = Presentation(stream_origen)
    prs_plantilla = Presentation(io.BytesIO(stream_plantilla_bytes))

    # Truco avanzado: Eliminar las diapositivas de ejemplo que traiga la plantilla
    # para dejar solo los formatos maestros limpios.
    xml_slides = prs_plantilla.slides._sldIdLst  
    slides = list(xml_slides)
    for s in slides:
        xml_slides.remove(s)

    # Procesar cada diapositiva de la presentación original
    for slide_orig in prs_origen.slides:
        
        # 1. Intentar usar el mismo índice de Layout (Diseño) que el original
        try:
            layout_idx = prs_origen.slide_layouts.index(slide_orig.slide_layout)
            # Asegurarse que la plantilla tenga ese índice, si no, usa el layout 1 (Título y Objetos)
            if layout_idx >= len(prs_plantilla.slide_layouts):
                layout_idx = 1 
        except:
            layout_idx = 1
            
        layout_target = prs_plantilla.slide_layouts[layout_idx]
        new_slide = prs_plantilla.slides.add_slide(layout_target)
        
        # 2. Mapear y transferir el contenido respetando el Master Format
        for shape_orig in slide_orig.shapes:
            if not shape_orig.has_text_frame or not shape_orig.text.strip():
                continue
                
            target_shape = None
            
            # Si es el Título original, pasarlo al Título de la nueva
            if shape_orig == slide_orig.shapes.title and new_slide.shapes.title:
                target_shape = new_slide.shapes.title
            else:
                # Buscar el siguiente marcador de posición (Placeholder) vacío para el cuerpo
                for placeholder in new_slide.placeholders:
                    if placeholder.placeholder_format.idx != 0 and placeholder.text == "":
                        target_shape = placeholder
                        break
            
            # Si se encontró un lugar donde poner el texto en la plantilla
            if target_shape and target_shape.has_text_frame:
                target_tf = target_shape.text_frame
                
                # Copiar párrafo por párrafo para mantener las viñetas (bullets)
                # IMPORTANTE: NO tocamos las fuentes/colores para que hereden de la plantilla
                for i, p_orig in enumerate(shape_orig.text_frame.paragraphs):
                    if i == 0:
                        p_new = target_tf.paragraphs[0]
                    else:
                        p_new = target_tf.add_paragraph()
                        
                    p_new.text = p_orig.text
                    try:
                        p_new.level = p_orig.level # Mantiene la indentación/viñetas
                    except:
                        pass
            else:
                # Fallback: Si la plantilla no tiene espacios suficientes, crea un cuadro de texto libre
                tb = new_slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
                tb.text = shape_orig.text

    output_stream = io.BytesIO()
    prs_plantilla.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# INTERFAZ DE USUARIO (UI) Y NAVEGACIÓN
# ==========================================

st.sidebar.header("📁 Gestor de Plantillas")
st.sidebar.write(f"Almacenadas: {len(st.session_state['plantillas_guardadas'])} / 5")

# Listar y eliminar plantillas
if st.session_state['plantillas_guardadas']:
    st.sidebar.markdown("**Tus Plantillas:**")
    for nombre_plantilla in list(st.session_state['plantillas_guardadas'].keys()):
        col1, col2 = st.sidebar.columns([0.8, 0.2])
        col1.write(f"🎨 {nombre_plantilla}")
        if col2.button("❌", key=f"del_{nombre_plantilla}"):
            del st.session_state['plantillas_guardadas'][nombre_plantilla]
            st.rerun()

st.sidebar.markdown("---")

# Subir plantillas nuevas
if len(st.session_state['plantillas_guardadas']) < 5:
    nuevas_plantillas = st.sidebar.file_uploader(
        "Añadir plantilla base (.pptx)", 
        type=["pptx"], 
        accept_multiple_files=True
    )
    if nuevas_plantillas:
        for p in nuevas_plantillas:
            if len(st.session_state['plantillas_guardadas']) < 5:
                st.session_state['plantillas_guardadas'][p.name] = p.getvalue()
        st.rerun()
else:
    st.sidebar.warning("⚠️ Máximo de 5 plantillas alcanzado.")

st.sidebar.info("💡 **Nota:** Ya no hay selector de colores. El sistema aplicará los colores, tamaños y tipografías directamente desde la plantilla que elijas.")

# ==========================================
# ÁREA PRINCIPAL Y PREVISUALIZACIÓN
# ==========================================
st.title("📊 Procesador de Diapositivas Inteligente")

archivo_original = st.file_uploader("1. Sube la presentación con el contenido a procesar (.pptx)", type=["pptx"])

if archivo_original and st.session_state['plantillas_guardadas']:
    st.markdown("---")
    
    nombres_plantillas = list(st.session_state['plantillas_guardadas'].keys())
    plantilla_seleccionada = st.selectbox("2. Elige la Plantilla Maestra a aplicar:", nombres_plantillas)
    bytes_plantilla = st.session_state['plantillas_guardadas'][plantilla_seleccionada]
    
    st.markdown(f"### 👀 Análisis de Secciones y Distribución")
    st.caption("Visualiza cómo la Inteligencia del script desglosa tu documento original y lo prepara para inyectarlo en los marcadores de la plantilla seleccionada.")
    
    datos_diapositivas = extraer_secciones_pptx(archivo_original)
    
    # Mostrar la previsualización seccionada
    for idx, slide in enumerate(datos_diapositivas):
        with st.container():
            st.markdown(f"**Diapositiva {idx+1}** (Diseño original detectado: `{slide['layout']}`)")
            col1, col2 = st.columns(2)
            
            # Caja de Título
            with col1:
                html_titulo = f"""
                <div style='border-left: 5px solid #0056b3; background-color: #f0f7ff; padding: 15px; border-radius: 5px; margin-bottom: 10px; height: 100%;'>
                    <small style='color: #0056b3; font-weight: bold;'>SECCIÓN: TÍTULO PRINCIPAL</small>
                    <h3 style='color: #333; margin-top: 5px;'>{slide['titulo'] if slide['titulo'] else '<i>Sin título detectado</i>'}</h3>
                    <small style='color: #666;'>→ Se aplicará el color y fuente del Título de '{plantilla_seleccionada}'</small>
                </div>
                """
                st.markdown(html_titulo, unsafe_allow_html=True)
            
            # Caja de Contenido
            with col2:
                texto_cuerpo = "<br>".join(slide['contenido']) if slide['contenido'] else "<i>Sin cuerpo de texto detectado</i>"
                html_cuerpo = f"""
                <div style='border-left: 5px solid #28a745; background-color: #f2fff5; padding: 15px; border-radius: 5px; margin-bottom: 10px; height: 100%;'>
                    <small style='color: #28a745; font-weight: bold;'>SECCIÓN: CUERPO Y VIÑETAS</small>
                    <p style='color: #444; margin-top: 5px; font-size: 14px;'>{texto_cuerpo}</p>
                    <small style='color: #666;'>→ Heredará distribución, colores base y tipografía de cuerpo de '{plantilla_seleccionada}'</small>
                </div>
                """
                st.markdown(html_cuerpo, unsafe_allow_html=True)
        st.write("") # Espaciador

    st.markdown("---")
    
    # Procesar y Descargar
    if st.button("🚀 Aplicar Plantilla y Generar PowerPoint", type="primary"):
        with st.spinner(f"Inyectando contenido en '{plantilla_seleccionada}'..."):
            ppt_final = procesar_pptx(archivo_original, bytes_plantilla)
            
        st.success("¡Documento transformado exitosamente bajo los estándares de la plantilla!")
        st.download_button(
            label="📥 Descargar Documento Final",
            data=ppt_final,
            file_name=f"PlantillaAplicada_{archivo_original.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

elif not st.session_state['plantillas_guardadas']:
    st.info("👈 Sube tus plantillas maestras en el menú de la izquierda para comenzar.")
elif not archivo_original:
    st.info("Sube un archivo original arriba para ver la distribución.")
