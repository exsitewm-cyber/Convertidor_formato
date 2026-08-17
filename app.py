import io
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

st.set_page_config(page_title="Editor Automático de PPTX v3.1", page_icon="📊", layout="wide")

if 'plantillas_guardadas' not in st.session_state:
    st.session_state['plantillas_guardadas'] = {}

def hex_a_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

def extraer_texto_pptx(file_stream, max_slides=3):
    file_stream.seek(0)
    prs = Presentation(file_stream)
    slides_data = []
    for i, slide in enumerate(prs.slides):
        if i >= max_slides: break
        slide_text = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip() != "":
                slide_text.append(shape.text.strip())
        slides_data.append(slide_text)
    file_stream.seek(0)
    return slides_data

# ==========================================
# NUEVA FUNCIÓN CORREGIDA
# ==========================================
def procesar_pptx(stream_origen, stream_plantilla_bytes, color_titulo, color_cuerpo, fuente):
    stream_origen.seek(0)
    stream_plantilla = io.BytesIO(stream_plantilla_bytes)
    
    prs_origen = Presentation(stream_origen)
    prs_plantilla = Presentation(stream_plantilla)

    # 1. Eliminar las diapositivas de ejemplo que vienen por defecto en la plantilla
    xml_slides = prs_plantilla.slides._sldIdLst  
    for sld in list(xml_slides):
        xml_slides.remove(sld)

    rgb_titulo = hex_a_rgb(color_titulo)
    rgb_cuerpo = hex_a_rgb(color_cuerpo)

    # 2. Procesar cada diapositiva
    for slide_origen in prs_origen.slides:
        # Usar el diseño "Título y Objetos" de la plantilla (usualmente índice 1)
        layout = prs_plantilla.slide_layouts[1] if len(prs_plantilla.slide_layouts) > 1 else prs_plantilla.slide_layouts[0]
        nueva_slide = prs_plantilla.slides.add_slide(layout)

        for shape_origen in slide_origen.shapes:
            if not shape_origen.has_text_frame:
                continue
            
            texto = shape_origen.text_frame.text.strip()
            if not texto:
                continue

            # A) Si el texto pertenece al TÍTULO de la diapositiva original
            if shape_origen == slide_origen.shapes.title:
                if nueva_slide.shapes.title:
                    tf_titulo = nueva_slide.shapes.title.text_frame
                    tf_titulo.text = texto
                    for p in tf_titulo.paragraphs:
                        p.font.name = fuente
                        p.font.color.rgb = rgb_titulo
                        p.font.bold = True
            
            # B) Si el texto pertenece al CUERPO de la diapositiva original
            else:
                body_placeholder = None
                # Buscar el marcador de posición del cuerpo (ignorando el título que es 0)
                for ph in nueva_slide.placeholders:
                    if ph.placeholder_format.idx != 0:
                        body_placeholder = ph
                        break
                
                if body_placeholder:
                    # Si ya hay texto, agregarlo abajo (por si hay múltiples cuadros en el origen)
                    if body_placeholder.text_frame.text:
                        body_placeholder.text_frame.text += "\n" + texto
                    else:
                        body_placeholder.text_frame.text = texto
                        
                    for p in body_placeholder.text_frame.paragraphs:
                        p.font.name = fuente
                        p.font.size = Pt(16)
                        p.font.color.rgb = rgb_cuerpo
                else:
                    # Respaldo de seguridad si la plantilla tiene un diseño inusual
                    txBox = nueva_slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
                    tf_cuerpo = txBox.text_frame
                    tf_cuerpo.text = texto
                    for p in tf_cuerpo.paragraphs:
                        p.font.name = fuente
                        p.font.size = Pt(16)
                        p.font.color.rgb = rgb_cuerpo

    output_stream = io.BytesIO()
    prs_plantilla.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# INTERFAZ (Sin cambios)
# ==========================================
st.sidebar.header("📁 Gestor de Plantillas")
st.sidebar.write(f"Almacenadas: {len(st.session_state['plantillas_guardadas'])} / 5")

if st.session_state['plantillas_guardadas']:
    st.sidebar.markdown("**Plantillas disponibles:**")
    for nombre_plantilla in list(st.session_state['plantillas_guardadas'].keys()):
        col1, col2 = st.sidebar.columns([0.8, 0.2])
        col1.write(f"📄 {nombre_plantilla}")
        if col2.button("❌", key=f"del_{nombre_plantilla}", help="Eliminar plantilla"):
            del st.session_state['plantillas_guardadas'][nombre_plantilla]
            st.rerun()

st.sidebar.markdown("---")

if len(st.session_state['plantillas_guardadas']) < 5:
    nuevas_plantillas = st.sidebar.file_uploader(
        "Subir nuevas plantillas (.pptx)", 
        type=["pptx"], 
        accept_multiple_files=True
    )
    if nuevas_plantillas:
        for p in nuevas_plantillas:
            if len(st.session_state['plantillas_guardadas']) < 5:
                st.session_state['plantillas_guardadas'][p.name] = p.getvalue()
        st.rerun()
else:
    st.sidebar.warning("⚠️ Límite de 5 plantillas alcanzado. Elimina una (❌) para poder subir otra.")

st.sidebar.header("🎨 Diseño y Marca")
color_titulo_hex = st.sidebar.color_picker("Color de Títulos", "#0F2043")
color_texto_hex = st.sidebar.color_picker("Color de Texto", "#323232")
nombre_fuente = st.sidebar.selectbox("Tipografía", ["Arial", "Calibri", "Helvetica", "Georgia", "Verdana"])

st.title("📊 Agente Procesador de Presentaciones v3.1")

archivo_original = st.file_uploader("1. Sube la presentación original (.pptx)", type=["pptx"])

if archivo_original and st.session_state['plantillas_guardadas']:
    st.markdown("---")
    nombres_plantillas = list(st.session_state['plantillas_guardadas'].keys())
    plantilla_seleccionada = st.selectbox("2. Elige la plantilla a aplicar:", nombres_plantillas)
    
    bytes_plantilla_elegida = st.session_state['plantillas_guardadas'][plantilla_seleccionada]
    
    st.markdown("### 👀 Previsualización Estructural (Primeras 3 diapositivas)")
    
    col_orig, col_mod = st.columns(2)
    datos_diapositivas = extraer_texto_pptx(archivo_original)
    
    with col_orig:
        st.markdown("**📄 Contenido Original**")
        for idx, textos in enumerate(datos_diapositivas):
            if not textos: continue
            html_orig = "<div style='border:1px solid #ddd; padding:15px; margin-bottom:10px; border-radius:8px; background-color:#f8f9fa;'>"
            for j, texto in enumerate(textos):
                if j == 0: html_orig += f"<h4 style='color:black; margin-bottom:5px;'>{texto}</h4>"
                else: html_orig += f"<p style='color:#555; font-size:14px;'>{texto}</p>"
            html_orig += "</div>"
            st.markdown(html_orig, unsafe_allow_html=True)

    with col_mod:
        st.markdown(f"**✨ Vista Previa con: {plantilla_seleccionada}**")
        for idx, textos in enumerate(datos_diapositivas):
            if not textos: continue
            html_mod = "<div style='border:1px solid #ccc; padding:15px; margin-bottom:10px; border-radius:8px; background-color:#ffffff; box-shadow: 2px 2px 8px rgba(0,0,0,0.1);'>"
            for j, texto in enumerate(textos):
                if j == 0: html_mod += f"<h4 style='color:{color_titulo_hex}; font-family:{nombre_fuente}; margin-bottom:5px;'>{texto}</h4>"
                else: html_mod += f"<p style='color:{color_texto_hex}; font-family:{nombre_fuente}; font-size:14px; margin-top:0px;'>{texto}</p>"
            html_mod += "</div>"
            st.markdown(html_mod, unsafe_allow_html=True)

    st.markdown("---")
    if st.button("🚀 Procesar y Generar PowerPoint Final", type="primary"):
        with st.spinner(f"Aplicando formato a tu presentación..."):
            ppt_final = procesar_pptx(
                archivo_original, 
                bytes_plantilla_elegida, 
                color_titulo_hex, 
                color_texto_hex, 
                nombre_fuente
            )
            
        st.success("¡Presentación procesada con éxito!")
        st.download_button(
            label="📥 Descargar PowerPoint Modificado",
            data=ppt_final,
            file_name=f"Formateado_{archivo_original.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

elif not st.session_state['plantillas_guardadas']:
    st.info("👈 Comienza subiendo al menos una plantilla en el panel lateral.")
elif not archivo_original:
    st.info("Sube la presentación original que deseas modificar en el recuadro superior.")
