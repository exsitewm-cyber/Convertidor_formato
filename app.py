import io
import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# Configuración de la página
st.set_page_config(page_title="Editor Automático de PPTX", page_icon="📊", layout="wide")

st.title("📊 Agente Procesador de Presentaciones PowerPoint")
st.write("Sube tu archivo base y tu plantilla para aplicar estilos y colores automáticos.")

# Convertir Hexadecimal a RGBColor de pptx
def hex_a_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

# Panel lateral para la configuración de marca
st.sidebar.header("🎨 Paleta de Colores y Fuente")
color_titulo_hex = st.sidebar.color_picker("Color de Títulos", "#0F2043")
color_texto_hex = st.sidebar.color_picker("Color de Texto", "#323232")
nombre_fuente = st.sidebar.selectbox("Tipografía", ["Arial", "Calibri", "Helvetica", "Georgia", "Verdana"])

# Área principal de carga de archivos
col1, col2 = st.columns(2)

with col1:
    archivo_original = st.file_uploader("1. Sube tu archivo PowerPoint (.pptx)", type=["pptx"])

with col2:
    archivo_plantilla = st.file_uploader("2. Sube tu Plantilla Base (.pptx)", type=["pptx"])

# Función para procesar la presentación en memoria
def procesar_pptx(stream_origen, stream_plantilla, color_titulo, color_cuerpo, fuente):
    prs_origen = Presentation(stream_origen)
    prs_plantilla = Presentation(stream_plantilla)

    rgb_titulo = hex_a_rgb(color_titulo)
    rgb_cuerpo = hex_a_rgb(color_cuerpo)

    for slide_origen in prs_origen.slides:
        # Usar el diseño de título y contenido (Layout 1) de la plantilla
        layout = prs_plantilla.slide_layouts[1] if len(prs_plantilla.slide_layouts) > 1 else prs_plantilla.slide_layouts[0]
        nueva_slide = prs_plantilla.slides.add_slide(layout)

        # Transferir texto e imágenes
        for shape in slide_origen.shapes:
            if shape.has_text_frame:
                # Si la plantilla tiene marcadores de posición, los aprovecha
                if len(nueva_slide.placeholders) > 1:
                    target_tf = nueva_slide.placeholders[1].text_frame
                else:
                    target_tf = nueva_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4.5)).text_frame

                target_tf.text = shape.text_frame.text

                # Aplicar formato de texto
                for i, p in enumerate(target_tf.paragraphs):
                    p.font.name = fuente
                    if i == 0:  # Primer párrafo (considerado título/encabezado)
                        p.font.size = Pt(22)
                        p.font.bold = True
                        p.font.color.rgb = rgb_titulo
                    else:  # Párrafos de cuerpo
                        p.font.size = Pt(16)
                        p.font.color.rgb = rgb_cuerpo

    # Guardar resultado en memoria buffer
    output_stream = io.BytesIO()
    prs_plantilla.save(output_stream)
    output_stream.seek(0)
    return output_stream

# Botón de ejecución y descarga
if archivo_original and archivo_plantilla:
    if st.button("🚀 Procesar Presentación", type="primary"):
        with st.spinner("Aplicando plantilla y formatos..."):
            ppt_modificado = procesar_pptx(
                archivo_original, 
                archivo_plantilla, 
                color_titulo_hex, 
                color_texto_hex, 
                nombre_fuente
            )

        st.success("¡Presentación procesada con éxito!")
        
        st.download_button(
            label="📥 Descargar PowerPoint Modificado",
            data=ppt_modificado,
            file_name="presentacion_formateada.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
elif archivo_original or archivo_plantilla:
    st.info("Por favor sube ambos archivos para continuar.")