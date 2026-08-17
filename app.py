import io
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# Configuración de página
st.set_page_config(page_title="Editor Automático de PPTX v3", page_icon="📊", layout="wide")

# Inicializar la "memoria" para guardar las plantillas
if 'plantillas_guardadas' not in st.session_state:
    st.session_state['plantillas_guardadas'] = {}

# Funciones de ayuda
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
                slide_text.append(shape.text)
        slides_data.append(slide_text)
    file_stream.seek(0)
    return slides_data

def procesar_pptx(stream_origen, stream_plantilla_bytes, color_titulo, color_cuerpo, fuente):
    stream_origen.seek(0)
    # Cargar la plantilla desde los bytes guardados en memoria
    stream_plantilla = io.BytesIO(stream_plantilla_bytes)
    
    prs_origen = Presentation(stream_origen)
    prs_plantilla = Presentation(stream_plantilla)

    rgb_titulo = hex_a_rgb(color_titulo)
    rgb_cuerpo = hex_a_rgb(color_cuerpo)

    for slide_origen in prs_origen.slides:
        layout = prs_plantilla.slide_layouts[1] if len(prs_plantilla.slide_layouts) > 1 else prs_plantilla.slide_layouts[0]
        nueva_slide = prs_plantilla.slides.add_slide(layout)

        for shape in slide_origen.shapes:
            if shape.has_text_frame:
                if len(nueva_slide.placeholders) > 1:
                    target_tf = nueva_slide.placeholders[1].text_frame
                else:
                    target_tf = nueva_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4.5)).text_frame

                target_tf.text = shape.text_frame.text

                for i, p in enumerate(target_tf.paragraphs):
                    p.font.name = fuente
                    if i == 0:
                        p.font.size = Pt(22)
                        p.font.bold = True
                        p.font.color.rgb = rgb_titulo
                    else:
                        p.font.size = Pt(16)
                        p.font.color.rgb = rgb_cuerpo

    output_stream = io.BytesIO()
    prs_plantilla.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# GESTOR DE PLANTILLAS (PANEL LATERAL)
# ==========================================
st.sidebar.header("📁 Gestor de Plantillas")
st.sidebar.write(f"Almacenadas: {len(st.session_state['plantillas_guardadas'])} / 5")

# 1. Mostrar las plantillas guardadas actualmente con opción a eliminar
if st.session_state['plantillas_guardadas']:
    st.sidebar.markdown("**Plantillas disponibles:**")
    for nombre_plantilla in list(st.session_state['plantillas_guardadas'].keys()):
        col1, col2 = st.sidebar.columns([0.8, 0.2])
        col1.write(f"📄 {nombre_plantilla}")
        # Botón para eliminar la plantilla
        if col2.button("❌", key=f"del_{nombre_plantilla}", help="Eliminar plantilla"):
            del st.session_state['plantillas_guardadas'][nombre_plantilla]
            st.rerun() # Recargar la página para actualizar la lista

st.sidebar.markdown("---")

# 2. Lógica para subir nuevas plantillas (Solo si hay menos de 5)
if len(st.session_state['plantillas_guardadas']) < 5:
    nuevas_plantillas = st.sidebar.file_uploader(
        "Subir nuevas plantillas (.pptx)", 
        type=["pptx"], 
        accept_multiple_files=True
    )
    
    if nuevas_plantillas:
        for p in nuevas_plantillas:
            # Verificar de nuevo el límite por si suben varios archivos de golpe
            if len(st.session_state['plantillas_guardadas']) < 5:
                # Guardar el contenido del archivo en la memoria de la sesión
                st.session_state['plantillas_guardadas'][p.name] = p.getvalue()
        st.rerun() # Recargar para que aparezcan arriba
else:
    st.sidebar.warning("⚠️ Límite de 5 plantillas alcanzado. Elimina una (❌) para poder subir otra.")

# Configuración de colores
st.sidebar.header("🎨 Diseño y Marca")
color_titulo_hex = st.sidebar.color_picker("Color de Títulos", "#0F2043")
color_texto_hex = st.sidebar.color_picker("Color de Texto", "#323232")
nombre_fuente = st.sidebar.selectbox("Tipografía", ["Arial", "Calibri", "Helvetica", "Georgia", "Verdana"])

# ==========================================
# ÁREA PRINCIPAL
# ==========================================
st.title("📊 Agente Procesador de Presentaciones v3")

archivo_original = st.file_uploader("1. Sube la presentación original que deseas modificar (.pptx)", type=["pptx"])

# Verificar si hay archivo original y si existe al menos una plantilla guardada
if archivo_original and st.session_state['plantillas_guardadas']:
    
    st.markdown("---")
    # Crear lista de nombres desde las plantillas guardadas
    nombres_plantillas = list(st.session_state['plantillas_guardadas'].keys())
    plantilla_seleccionada = st.selectbox("2. Elige la plantilla a aplicar:", nombres_plantillas)
    
    # Obtener los bytes de la plantilla seleccionada desde la memoria
    bytes_plantilla_elegida = st.session_state['plantillas_guardadas'][plantilla_seleccionada]
    
    st.markdown("### 👀 Previsualización Estructural (Primeras 3 diapositivas)")
    
    col_orig, col_mod = st.columns(2)
    datos_diapositivas = extraer_texto_pptx(archivo_original)
    
    with col_orig:
        st.markdown("**📄 Contenido Original**")
        for idx, textos in enumerate(datos_diapositivas):
            html_orig = "<div style='border:1px solid #ddd; padding:15px; margin-bottom:10px; border-radius:8px; background-color:#f8f9fa;'>"
            for j, texto in enumerate(textos):
                if j == 0: html_orig += f"<h4 style='color:black;'>{texto}</h4>"
                else: html_orig += f"<p style='color:#555;'>{texto}</p>"
            html_orig += "</div>"
            st.markdown(html_orig, unsafe_allow_html=True)

    with col_mod:
        st.markdown(f"**✨ Vista Previa con: {plantilla_seleccionada}**")
        for idx, textos in enumerate(datos_diapositivas):
            html_mod = "<div style='border:1px solid #ccc; padding:15px; margin-bottom:10px; border-radius:8px; background-color:#ffffff; box-shadow: 2px 2px 8px rgba(0,0,0,0.1);'>"
            for j, texto in enumerate(textos):
                if j == 0: html_mod += f"<h4 style='color:{color_titulo_hex}; font-family:{nombre_fuente}; margin-bottom:5px;'>{texto}</h4>"
                else: html_mod += f"<p style='color:{color_texto_hex}; font-family:{nombre_fuente}; margin-top:0px;'>{texto}</p>"
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
