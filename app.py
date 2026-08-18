import io
import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

# Configuración de página
st.set_page_config(page_title="Editor PPTX Avanzado v4", page_icon="📊", layout="wide")

if 'plantillas_guardadas' not in st.session_state:
    st.session_state['plantillas_guardadas'] = {}

# ==========================================
# MOTOR DE PROCESAMIENTO Y MAPEO PPTX
# ==========================================

def analizar_textos_originales(file_stream, max_slides=5):
    """Extrae el texto puro del archivo a modificar, sin ningún formato."""
    file_stream.seek(0)
    prs = Presentation(file_stream)
    datos = []
    
    for i, slide in enumerate(prs.slides):
        if i >= max_slides: break
        diapositiva_datos = {"titulo": "", "cuerpo": []}
        
        # Extraer Título
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            diapositiva_datos["titulo"] = slide.shapes.title.text.strip()
            
        # Extraer Cuerpo (párrafos y niveles de viñetas)
        for shape in slide.shapes:
            if shape == slide.shapes.title: continue
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    texto = p.text.strip()
                    if texto:
                        diapositiva_datos["cuerpo"].append({"texto": texto, "nivel": p.level})
                        
        datos.append(diapositiva_datos)
    file_stream.seek(0)
    return datos

def analizar_layout_plantilla(stream_plantilla_bytes):
    """Analiza los espacios (Placeholders) disponibles en el diseño de la plantilla seleccionada."""
    prs_plantilla = Presentation(io.BytesIO(stream_plantilla_bytes))
    layout = prs_plantilla.slide_layouts[1] if len(prs_plantilla.slide_layouts) > 1 else prs_plantilla.slide_layouts[0]
    
    placeholders_info = []
    for ph in layout.placeholders:
        tipo = ph.placeholder_format.type
        nombre_tipo = "TÍTULO" if tipo == PP_PLACEHOLDER.TITLE else ("CUERPO/OBJETO" if tipo in [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT] else f"OTRO ({tipo})")
        placeholders_info.append({"id": ph.placeholder_format.idx, "tipo": nombre_tipo, "nombre": ph.name})
        
    return placeholders_info

def inyectar_formato_puro(stream_origen, stream_plantilla_bytes):
    """Genera el PPTX forzando el formato maestro de la plantilla seleccionada."""
    stream_origen.seek(0)
    prs_origen = Presentation(stream_origen)
    prs_plantilla = Presentation(io.BytesIO(stream_plantilla_bytes))

    # Limpiar diapositivas de ejemplo que traiga la plantilla
    xml_slides = prs_plantilla.slides._sldIdLst  
    for s in list(xml_slides):
        xml_slides.remove(s)

    for slide_orig in prs_origen.slides:
        # Extraer texto crudo de la original
        titulo_texto = slide_orig.shapes.title.text.strip() if slide_orig.shapes.title else ""
        cuerpo_datos = []
        for shape in slide_orig.shapes:
            if shape == slide_orig.shapes.title: continue
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    if p.text.strip():
                        cuerpo_datos.append((p.text, p.level))

        # Crear nueva diapositiva basada en la plantilla elegida
        layout = prs_plantilla.slide_layouts[1] if len(prs_plantilla.slide_layouts) > 1 else prs_plantilla.slide_layouts[0]
        new_slide = prs_plantilla.slides.add_slide(layout)

        # Inyectar Título (al asignar .text directamente, hereda 100% el formato de la plantilla)
        if new_slide.shapes.title and titulo_texto:
            new_slide.shapes.title.text = titulo_texto

        # Inyectar Cuerpo
        cuerpo_inyectado = False
        for ph in new_slide.placeholders:
            if ph.placeholder_format.type in [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT]:
                tf = ph.text_frame
                tf.clear() # Limpia texto predeterminado del marcador
                
                for texto, nivel in cuerpo_datos:
                    p = tf.add_paragraph()
                    p.text = texto # Asignación cruda: fuerza la fuente y color del Patrón
                    try:
                        p.level = nivel # Mantiene la jerarquía de viñetas
                    except:
                        pass
                cuerpo_inyectado = True
                break # Solo llenamos el primer marcador de cuerpo principal

    output_stream = io.BytesIO()
    prs_plantilla.save(output_stream)
    output_stream.seek(0)
    return output_stream

# ==========================================
# INTERFAZ (UI) Y NAVEGADOR
# ==========================================

st.sidebar.header("📁 Tus Plantillas Maestras")
st.sidebar.write(f"Almacenadas: {len(st.session_state['plantillas_guardadas'])} / 5")

if st.session_state['plantillas_guardadas']:
    st.sidebar.markdown("---")
    for nombre_plantilla in list(st.session_state['plantillas_guardadas'].keys()):
        col1, col2 = st.sidebar.columns([0.8, 0.2])
        col1.write(f"🎨 {nombre_plantilla}")
        if col2.button("❌", key=f"del_{nombre_plantilla}"):
            del st.session_state['plantillas_guardadas'][nombre_plantilla]
            st.rerun()

st.sidebar.markdown("---")
if len(st.session_state['plantillas_guardadas']) < 5:
    nuevas_plantillas = st.sidebar.file_uploader(
        "Añadir plantilla (.pptx)", type=["pptx"], accept_multiple_files=True
    )
    if nuevas_plantillas:
        for p in nuevas_plantillas:
            if len(st.session_state['plantillas_guardadas']) < 5:
                st.session_state['plantillas_guardadas'][p.name] = p.getvalue()
        st.rerun()

st.title("📊 Procesador de Plantillas PPTX Exacto")

archivo_original = st.file_uploader("1. Sube tu presentación original (.pptx)", type=["pptx"])

if archivo_original and st.session_state['plantillas_guardadas']:
    st.markdown("---")
    nombres_plantillas = list(st.session_state['plantillas_guardadas'].keys())
    
    # 2. Selector de Plantillas (Al cambiar esto, cambia la vista previa dinámicamente)
    plantilla_seleccionada = st.selectbox("2. Selecciona la Plantilla a Aplicar:", nombres_plantillas)
    bytes_plantilla = st.session_state['plantillas_guardadas'][plantilla_seleccionada]
    
    # Analizar cómo está construida la plantilla seleccionada
    estructura_plantilla = analizar_layout_plantilla(bytes_plantilla)
    datos_originales = analizar_textos_originales(archivo_original)
    
    st.markdown(f"### 👀 Mapeo Dinámico: {plantilla_seleccionada}")
    st.caption("Esta vista te muestra exactamente en qué secciones de la plantilla seleccionada se distribuirá tu texto original. El archivo descargado tendrá los colores y fuentes definidos por el diseñador de esta plantilla.")
    
    # Mostrar la estructura detectada en la plantilla
    st.info(f"**Estructura detectada en '{plantilla_seleccionada}':** " + 
            " | ".join([f"📦 {ph['tipo']} ({ph['nombre']})" for ph in estructura_plantilla]))
    
    # Mostrar el preview dinámico (Primeras 3 diapositivas)
    for idx, slide in enumerate(datos_originales[:3]):
        st.markdown(f"#### Diapositiva {idx+1}")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Contenido Original Extraído:**")
            st.markdown(f"> **TÍTULO:** {slide['titulo'] if slide['titulo'] else '*(Vacío)*'}")
            cuerpo_str = "\n".join([f"- {item['texto']}" for item in slide['cuerpo']])
            st.markdown(f"> **CUERPO:**\n{cuerpo_str if cuerpo_str else '*(Vacío)*'}")
            
        with col2:
            st.markdown(f"**Cómo se inyectará en '{plantilla_seleccionada}':**")
            html_preview = "<div style='border: 2px dashed #4CAF50; padding: 15px; border-radius: 10px; background-color: #fafafa;'>"
            
            # Simular Título en Plantilla
            if slide['titulo']:
                html_preview += f"<div style='background-color:#e8f5e9; padding:10px; margin-bottom:10px; border-radius:5px;'><small style='color:#2E7D32;'>Enviado al marcador: TÍTULO</small><br><b>{slide['titulo']}</b></div>"
            
            # Simular Cuerpo en Plantilla
            if slide['cuerpo']:
                html_preview += f"<div style='background-color:#e3f2fd; padding:10px; border-radius:5px;'><small style='color:#1565C0;'>Enviado al marcador: CUERPO/OBJETO (Heredando fuente, color y viñetas de la plantilla)</small><br>"
                for item in slide['cuerpo']:
                    indent = "&nbsp;" * (item['nivel'] * 4)
                    html_preview += f"<div>{indent}• {item['texto']}</div>"
                html_preview += "</div>"
                
            html_preview += "</div>"
            st.markdown(html_preview, unsafe_allow_html=True)
            
        st.write("---")

    # Procesar y Descargar
    if st.button("🚀 Aplicar Plantilla y Descargar Archivo", type="primary"):
        with st.spinner(f"Aplicando el formato exacto de '{plantilla_seleccionada}'..."):
            ppt_final = inyectar_formato_puro(archivo_original, bytes_plantilla)
            
        st.success("¡Documento procesado! El formato, tipografía y colores de la plantilla han sido forzados.")
        st.download_button(
            label="📥 Descargar Presentación Final",
            data=ppt_final,
            file_name=f"Final_{plantilla_seleccionada}_{archivo_original.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
