import io
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt

st.set_page_config(page_title="Editor Automático PPTX v4", page_icon="📊", layout="wide")

if 'plantillas_guardadas' not in st.session_state:
    st.session_state['plantillas_guardadas'] = {}

def hex_a_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

def extraer_texto_estructurado(file_stream, max_slides=3):
    """Extrae separando inteligentemente el Título del Cuerpo."""
    file_stream.seek(0)
    prs = Presentation(file_stream)
    slides_data = []
    
    for i, slide in enumerate(prs.slides):
        if i >= max_slides: break
        titulo = ""
        cuerpo = []
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            
            # Identificar si la forma es el título de la diapositiva
            if shape == slide.shapes.title:
                titulo = shape.text.strip()
            else:
                # Extraer viñetas/párrafos
                for p in shape.text_frame.paragraphs:
                    if p.text.strip():
                        cuerpo.append(p.text.strip())
                        
        slides_data.append({"titulo": titulo, "cuerpo": cuerpo})
        
    file_stream.seek(0)
    return slides_data

def procesar_pptx(stream_origen, stream_plantilla_bytes, color_tit, color_cuer, fuente, size_tit, size_cuer):
    stream_origen.seek(0)
    prs_origen = Presentation(stream_origen)
    prs_plantilla = Presentation(io.BytesIO(stream_plantilla_bytes))

    rgb_titulo = hex_a_rgb(color_tit)
    rgb_cuerpo = hex_a_rgb(color_cuer)

    for i, slide_origen in enumerate(prs_origen.slides):
        # DISTRIBUCIÓN MEJORADA: Diapositiva 1 = Portada, Resto = Título y Contenido
        if i == 0:
            layout = prs_plantilla.slide_layouts[0] # Slide Master: Portada
        else:
            layout = prs_plantilla.slide_layouts[1] if len(prs_plantilla.slide_layouts) > 1 else prs_plantilla.slide_layouts[0]
            
        nueva_slide = prs_plantilla.slides.add_slide(layout)

        # 1. Extraer textos
        origen_titulo = ""
        origen_cuerpo = []
        for shape in slide_origen.shapes:
            if not shape.has_text_frame: continue
            if shape == slide_origen.shapes.title:
                origen_titulo = shape.text.strip()
            else:
                for p in shape.text_frame.paragraphs:
                    if p.text.strip(): origen_cuerpo.append(p.text.strip())

        # 2. Aplicar Título y Formato
        if nueva_slide.shapes.title:
            tf = nueva_slide.shapes.title.text_frame
            tf.text = origen_titulo
            for p in tf.paragraphs:
                p.font.name = fuente
                p.font.size = Pt(size_tit)
                p.font.color.rgb = rgb_titulo
                p.font.bold = True

        # 3. Aplicar Cuerpo y Formato
        if len(nueva_slide.placeholders) > 1:
            # Encontrar el cuadro de texto principal (índice 1)
            body_shape = None
            for shape in nueva_slide.placeholders:
                if shape != nueva_slide.shapes.title and shape.has_text_frame:
                    body_shape = shape
                    break
            
            if body_shape and origen_cuerpo:
                tf = body_shape.text_frame
                tf.clear() # Limpiar texto de relleno de la plantilla
                for idx, texto in enumerate(origen_cuerpo):
                    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                    p.text = texto
                    p.font.name = fuente
                    p.font.size = Pt(size_cuer)
                    p.font.color.rgb = rgb_cuerpo

    output = io.BytesIO()
    prs_plantilla.save(output)
    output.seek(0)
    return output

# ==========================================
# PANEL LATERAL: PLANTILLAS Y ESTILOS
# ==========================================
st.sidebar.header("📁 1. Gestor de Plantillas")
if st.session_state['plantillas_guardadas']:
    for nombre in list(st.session_state['plantillas_guardadas'].keys()):
        col1, col2 = st.sidebar.columns([0.8, 0.2])
        col1.write(f"📄 {nombre}")
        if col2.button("❌", key=f"del_{nombre}"):
            del st.session_state['plantillas_guardadas'][nombre]
            st.rerun()

st.sidebar.markdown("---")
if len(st.session_state['plantillas_guardadas']) < 5:
    nuevas = st.sidebar.file_uploader("Subir plantillas (.pptx)", type=["pptx"], accept_multiple_files=True)
    if nuevas:
        for p in nuevas:
            if len(st.session_state['plantillas_guardadas']) < 5:
                st.session_state['plantillas_guardadas'][p.name] = p.getvalue()
        st.rerun()
else:
    st.sidebar.warning("⚠️ Límite de 5 plantillas. Elimina una para subir otra.")

st.sidebar.header("🎨 2. Formato y Paleta Exacta")
fuente = st.sidebar.selectbox("Tipografía", ["Arial", "Calibri", "Helvetica", "Georgia", "Times New Roman"])

col_c1, col_c2 = st.sidebar.columns(2)
color_tit = col_c1.color_picker("Color Título", "#0F2043")
color_cuer = col_c2.color_picker("Color Cuerpo", "#323232")

col_s1, col_s2 = st.sidebar.columns(2)
size_tit = col_s1.number_input("Tamaño Título (Pt)", min_value=12, max_value=72, value=32)
size_cuer = col_s2.number_input("Tamaño Cuerpo (Pt)", min_value=10, max_value=40, value=18)

# ==========================================
# ÁREA PRINCIPAL
# ==========================================
st.title("📊 Agente Procesador de Presentaciones v4")

archivo_origen = st.file_uploader("Sube el PowerPoint original a modificar (.pptx)", type=["pptx"])

if archivo_origen and st.session_state['plantillas_guardadas']:
    
    st.markdown("---")
    plantilla_seleccionada = st.selectbox("Elige la plantilla base:", list(st.session_state['plantillas_guardadas'].keys()))
    bytes_plantilla = st.session_state['plantillas_guardadas'][plantilla_seleccionada]
    
    st.markdown("### 👀 Previsualización de Distribución y Formato")
    
    col_orig, col_mod = st.columns(2)
    datos = extraer_texto_estructurado(archivo_origen)
    
    # Renderizar Original
    with col_orig:
        st.markdown("**📄 Documento Original (Sin Formato)**")
        for idx, slide in enumerate(datos):
            st.markdown(f"""
            <div style='border:1px solid #ddd; padding:20px; margin-bottom:15px; background:#f8f9fa;'>
                <h4 style='color:black;'>{slide['titulo'] if slide['titulo'] else '[Sin Título]'}</h4>
                <ul>{''.join([f"<li style='color:#555;'>{item}</li>" for item in slide['cuerpo']])}</ul>
            </div>
            """, unsafe_allow_html=True)

    # Renderizar Preview con Formato Aplicado
    with col_mod:
        st.markdown(f"**✨ Resultado con Plantilla y Estilos Aplicados**")
        for idx, slide in enumerate(datos):
            st.markdown(f"""
            <div style='border:1px solid #ccc; padding:20px; margin-bottom:15px; background:#fff; box-shadow: 2px 2px 10px rgba(0,0,0,0.1);'>
                <div style='color:{color_tit}; font-family:{fuente}; font-size:{size_tit}px; font-weight:bold; margin-bottom:10px;'>
                    {slide['titulo'] if slide['titulo'] else '[Sin Título]'}
                </div>
                <ul style='color:{color_cuer}; font-family:{fuente}; font-size:{size_cuer}px;'>
                    {''.join([f"<li>{item}</li>" for item in slide['cuerpo']])}
                </ul>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")
    if st.button("🚀 Aplicar Formato y Descargar", type="primary"):
        with st.spinner("Modificando archivo original..."):
            ppt_final = procesar_pptx(
                archivo_origen, bytes_plantilla, color_tit, color_cuer, fuente, size_tit, size_cuer
            )
            
        st.success("¡Documento generado perfectamente!")
        st.download_button(
            label="📥 Descargar PowerPoint Final",
            data=ppt_final,
            file_name=f"Plantilla_Aplicada_{archivo_origen.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
