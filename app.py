import io
import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from lxml import etree

st.set_page_config(page_title="Editor PPTX Avanzado", page_icon="📊", layout="wide")

# ==========================================
# INICIALIZACIÓN DE MEMORIA
# ==========================================
if 'plantillas_guardadas' not in st.session_state:
    st.session_state['plantillas_guardadas'] = {}

# ==========================================
# 1. EXTRACTORES DE DATOS Y TEMAS
# ==========================================
def extraer_tema_plantilla(bytes_plantilla):
    """Lee el XML interno de la plantilla para el Preview Web (Fuentes y Colores)"""
    theme_data = {
        'font_title': 'sans-serif', 'font_body': 'sans-serif',
        'color_title': '#003366', 'color_body': '#333333', 'bg_color': '#ffffff'
    }
    try:
        prs = Presentation(io.BytesIO(bytes_plantilla))
        for part in prs.part.package.parts:
            if part.partname.startswith('/ppt/theme/'):
                root = etree.fromstring(part.blob)
                ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                
                maj_font = root.xpath('//a:majorFont/a:latin/@typeface', namespaces=ns)
                if maj_font: theme_data['font_title'] = maj_font[0]
                min_font = root.xpath('//a:minorFont/a:latin/@typeface', namespaces=ns)
                if min_font: theme_data['font_body'] = min_font[0]
                
                dk1 = root.xpath('//a:clrScheme/a:dk1/a:srgbClr/@val', namespaces=ns)
                if dk1: 
                    theme_data['color_title'] = '#' + dk1[0]
                    theme_data['color_body'] = '#' + dk1[0]
                
                lt1 = root.xpath('//a:clrScheme/a:lt1/a:srgbClr/@val', namespaces=ns)
                if lt1: theme_data['bg_color'] = '#' + lt1[0]
                break
    except Exception:
        pass
    return theme_data

def extraer_secciones_pptx(file_stream, max_slides=5):
    """Extrae el contenido original dividiendo Título y Cuerpo"""
    file_stream.seek(0)
    prs = Presentation(file_stream)
    slides_data = []
    
    for i, slide in enumerate(prs.slides):
        if i >= max_slides: break
        slide_info = {"titulo": "", "cuerpo": []}
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.text.strip(): continue
            if shape == slide.shapes.title:
                slide_info["titulo"] = shape.text
            else:
                for p in shape.text_frame.paragraphs:
                    if p.text.strip():
                        slide_info["cuerpo"].append({'text': p.text, 'level': p.level})
        slides_data.append(slide_info)
    file_stream.seek(0)
    return slides_data

# ==========================================
# 2. PROCESAMIENTO Y APLICACIÓN DE PLANTILLA
# ==========================================
def procesar_pptx(stream_origen, stream_plantilla_bytes):
    """Aplica la plantilla transfiriendo el texto a los placeholders nativos"""
    stream_origen.seek(0)
    prs_origen = Presentation(stream_origen)
    prs_plantilla = Presentation(io.BytesIO(stream_plantilla_bytes))

    xml_slides = prs_plantilla.slides._sldIdLst  
    for s in list(xml_slides): xml_slides.remove(s)

    for slide_orig in prs_origen.slides:
        layout_idx = 0 if slide_orig.slide_layout.name == "Title Slide" else 1 
        if layout_idx >= len(prs_plantilla.slide_layouts): layout_idx = 1
        
        layout_target = prs_plantilla.slide_layouts[layout_idx]
        new_slide = prs_plantilla.slides.add_slide(layout_target)
        
        titulo = ""
        cuerpo = []
        for shape in slide_orig.shapes:
            if not shape.has_text_frame or not shape.text.strip(): continue
            if shape == slide_orig.shapes.title:
                titulo = shape.text
            else:
                for p in shape.text_frame.paragraphs:
                    if p.text.strip(): cuerpo.append({'text': p.text, 'level': p.level})

        if new_slide.shapes.title and titulo:
            new_slide.shapes.title.text = titulo
            
        if cuerpo:
            body_ph = None
            for ph in new_slide.placeholders:
                if ph.placeholder_format.idx != 0:
                    body_ph = ph
                    break
            
            if body_ph:
                tf = body_ph.text_frame
                tf.clear() 
                for i, p_data in enumerate(cuerpo):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = p_data['text']
                    try: p.level = p_data['level']
                    except: pass
            else:
                tb = new_slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
                tf = tb.text_frame
                for i, p_data in enumerate(cuerpo):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = p_data['text']

    output_stream = io.BytesIO()
    prs_plantilla.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# INTERFAZ DE USUARIO (UI)
# ==========================================

# --- PANEL LATERAL: Gestor de Plantillas ---
st.sidebar.header("📁 Gestor de Plantillas")
if st.session_state['plantillas_guardadas']:
    for nombre_plantilla in list(st.session_state['plantillas_guardadas'].keys()):
        col1, col2 = st.sidebar.columns([0.8, 0.2])
        col1.write(f"🎨 {nombre_plantilla}")
        if col2.button("❌", key=f"del_{nombre_plantilla}"):
            del st.session_state['plantillas_guardadas'][nombre_plantilla]
            st.rerun()

st.sidebar.markdown("---")
if len(st.session_state['plantillas_guardadas']) < 5:
    nuevas_plantillas = st.sidebar.file_uploader("Añadir plantilla (.pptx)", type=["pptx"], accept_multiple_files=True)
    if nuevas_plantillas:
        for p in nuevas_plantillas:
            if len(st.session_state['plantillas_guardadas']) < 5:
                st.session_state['plantillas_guardadas'][p.name] = p.getvalue()
        st.rerun()

# --- ÁREA PRINCIPAL ---
st.title("📊 Editor y Previsualizador de Diapositivas")

archivo_original = st.file_uploader("1. Sube tu archivo a modificar (.pptx)", type=["pptx"])

if archivo_original and st.session_state['plantillas_guardadas']:
    st.markdown("---")
    
    plantilla_seleccionada = st.selectbox("2. Elige la Plantilla a aplicar:", list(st.session_state['plantillas_guardadas'].keys()))
    bytes_plantilla = st.session_state['plantillas_guardadas'][plantilla_seleccionada]
    
    tema = extraer_tema_plantilla(bytes_plantilla)
    datos_diapositivas = extraer_secciones_pptx(archivo_original)
    
    # ==========================================
    # SECCIÓN DE PREVIEW (LADO A LADO)
    # ==========================================
    st.markdown("### 👁️ Sección de Preview (Antes y Después)")
    st.caption("Compara cómo se reorganizan y formatean tus textos. El panel derecho cambiará según la plantilla seleccionada.")
    
    bg_color = tema['bg_color'] if tema['bg_color'] != '#000000' else '#ffffff'
    
    for idx, slide in enumerate(datos_diapositivas):
        st.markdown(f"**Diapositiva {idx+1}**")
        col_orig, col_mod = st.columns(2)
        
        # COLUMNA IZQUIERDA: ORIGINAL
        with col_orig:
            html_orig = f"""
            <div style="background-color: #f8f9fa; border: 1px dashed #ccc; padding: 20px; border-radius: 8px; height: 100%;">
                <span style="color: #666; font-size: 12px; font-weight: bold;">DOCUMENTO ORIGINAL</span>
                <h3 style="color: #333; margin-top: 10px; font-family: sans-serif;">{slide['titulo'] if slide['titulo'] else '[Sin Título]'}</h3>
                <div style="color: #555; font-family: sans-serif; font-size: 14px;">
            """
            if slide['cuerpo']:
                for item in slide['cuerpo']:
                    html_orig += f"<div style='margin-left: {item['level']*15 if item['level'] else 0}px;'>• {item['text']}</div>"
            else:
                html_orig += "<i>[Sin contenido]</i>"
            html_orig += "</div></div>"
            st.markdown(html_orig, unsafe_allow_html=True)
            
        # COLUMNA DERECHA: PREVIEW CON PLANTILLA
        with col_mod:
            html_mod = f"""
            <div style="background-color: {bg_color}; border: 2px solid {tema['color_title']}55; padding: 20px; border-radius: 8px; height: 100%; box-shadow: 3px 3px 10px rgba(0,0,0,0.05);">
                <span style="background-color: {tema['color_title']}; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px;">PREVIEW: {plantilla_seleccionada}</span>
                <h2 style="font-family: '{tema['font_title']}', sans-serif; color: {tema['color_title']}; margin-top: 10px; margin-bottom: 15px; border-bottom: 1px solid {tema['color_title']}33; padding-bottom: 5px;">
                    {slide['titulo'] if slide['titulo'] else '[Sin Título]'}
                </h2>
                <div style="font-family: '{tema['font_body']}', sans-serif; color: {tema['color_body']}; font-size: 16px;">
            """
            if slide['cuerpo']:
                for item in slide['cuerpo']:
                    html_mod += f"<div style='margin-left: {item['level']*20 if item['level'] else 0}px; margin-bottom: 5px;'>• {item['text']}</div>"
            else:
                html_mod += "<i>[Sin contenido]</i>"
            html_mod += "</div></div>"
            st.markdown(html_mod, unsafe_allow_html=True)
            
        st.write("") # Espaciador entre diapositivas

    st.markdown("---")
    
    # ==========================================
    # PROCESAMIENTO FINAL Y DESCARGA
    # ==========================================
    if st.button("🚀 Aplicar Formato y Descargar PPTX", type="primary"):
        with st.spinner(f"Modificando archivo usando la plantilla '{plantilla_seleccionada}'..."):
            ppt_final = procesar_pptx(archivo_original, bytes_plantilla)
            
        st.success("¡Archivo generado! Todo el formato (letra, color, viñetas, distribución) ha sido aplicado correctamente.")
        st.download_button(
            label="📥 Descargar Presentación Final",
            data=ppt_final,
            file_name=f"Final_{plantilla_seleccionada}_{archivo_original.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
