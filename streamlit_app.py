import streamlit as st
import datetime
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import io
# import gspread
# from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Rejilla OPT con Dibujo", layout="wide")

st.title("📋 Rejilla de Observación de Puesto (OPT)")
st.write("Complete los datos de la observación de terreno. Al final podrá adjuntar un croquis o dibujo.")

# --- FORMULARIO PRINCIPAL ---
with st.form("rejilla_opt_form"):
    
    # SECCIÓN 0: ENCABEZADO
    st.header("Información General")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        fecha = st.date_input("Fecha", datetime.date.today())
    with col2:
        observador = st.text_input("Observador*")
    with col3:
        ute_equipo = st.text_input("UTE / Equipo")
    with col4:
        puesto_operario = st.text_input("Puesto / Operario*")

    st.divider()

    # SECCIÓN 1: PREPARACIÓN
    st.header("1. Preparación de la Observación")
    col1_1, col1_2 = st.columns(2)
    with col1_1:
        q1_1 = st.radio("1.1. ¿Los estándares están al día y completos?", ["Sí", "No"])
        q1_2_skill = st.selectbox("1.2. Nivel Skill del operario", ["Básico", "Intermedio", "Avanzado", "Experto"])
        
    with col1_2:
        q1_4 = st.radio("1.4. ¿Algún problema de Ergonomía o Seguridad?", ["No", "Sí"])
        if q1_4 == "Sí":
            q1_4_detalle = st.text_input("Detalle el problema de seguridad:")
        else:
            q1_4_detalle = "N/A"

    st.divider()

    # SECCIÓN DE DIBUJO / CROQUIS
    st.header("Anexo Visual: Croquis de la Estación")
    st.write("Dibuja un esquema rápido del puesto o señala dónde está el problema detectado.")
    
    # Controles del lienzo
    col_tools1, col_tools2 = st.columns(2)
    with col_tools1:
        stroke_width = st.slider("Grosor del lápiz:", 1, 10, 2)
    with col_tools2:
        stroke_color = st.color_picker("Color del lápiz:", "#FF0000") # Rojo por defecto para marcar errores

    # Lienzo incrustado
    canvas_result = st_canvas(
        fill_color="rgba(255, 165, 0, 0.3)",
        stroke_width=stroke_width,
        stroke_color=stroke_color,
        background_color="#f0f2f6",
        height=300,
        width=800,
        drawing_mode="freedraw",
        key="canvas",
    )

    st.divider()

    # BOTÓN DE ENVÍO
    enviado = st.form_submit_button("Guardar Observación y Enviar", type="primary")

# --- LÓGICA DE PROCESAMIENTO POST-ENVÍO ---
if enviado:
    if observador and puesto_operario:
        st.success("Procesando los datos...")
        
        # 1. Empaquetar los datos de texto para Google Sheets
        fila_datos = [
            str(fecha), observador, ute_equipo, puesto_operario,
            q1_1, q1_2_skill, q1_4, q1_4_detalle
        ]
        
        # Aquí enviarías la fila a Sheets: 
        # sheet.append_row(fila_datos)
        st.write("**Datos de texto listos para Sheets:**", fila_datos)
        
        # 2. Procesar el dibujo si el usuario hizo trazos
        if canvas_result.image_data is not None:
            # Convertir el array a una imagen PNG real
            img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
            
            # Comprobar si el lienzo está en blanco o tiene dibujo
            # (Opcional: lógica avanzada para evitar subir lienzos vacíos)
            
            st.write("**Croquis capturado:**")
            st.image(img, caption="Croquis adjunto a la OPT")
            
            # Aquí iría el código para subir 'img' a Google Drive o AWS
            # url_imagen = subir_a_drive(img)
            # sheet.update_cell(fila_actual, columna_imagen, f'=IMAGE("{url_imagen}")')
            
        st.balloons()
    else:
        st.error("⚠️ Por favor, completa los campos obligatorios (marcados con *) antes de enviar.")
