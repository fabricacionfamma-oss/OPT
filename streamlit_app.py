import streamlit as st
import datetime
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import io

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Rejilla OPT Completa", layout="wide")

st.title("📋 Rejilla de Observación de Puesto (OPT)")
st.info("Formulario completo basado en el formato de Rejilla_OPT_ES_2.xlsx")

# --- FUNCIÓN AUXILIAR PARA PREGUNTAS Y OBSERVACIONES ---
# Esta función crea la pregunta a la izquierda y el cuadro de observaciones a la derecha
def hacer_pregunta(id_pregunta, texto_pregunta, tipo="radio", opciones=["Sí", "No"]):
    col_preg, col_obs = st.columns([2, 1]) # Proporción: 2 tercios para pregunta, 1 tercio para observación
    
    with col_preg:
        if tipo == "radio":
            respuesta = st.radio(texto_pregunta, opciones, key=f"q_{id_pregunta}")
        elif tipo == "texto":
            respuesta = st.text_input(texto_pregunta, key=f"q_{id_pregunta}")
        elif tipo == "area":
            respuesta = st.text_area(texto_pregunta, key=f"q_{id_pregunta}")
        elif tipo == "select":
            respuesta = st.selectbox(texto_pregunta, opciones, key=f"q_{id_pregunta}")
            
    with col_obs:
        # st.text_area permite múltiples líneas. Ajustamos la altura para que encaje bien.
        observacion = st.text_area("Observaciones", key=f"obs_{id_pregunta}", height=68, label_visibility="collapsed", placeholder="Escribe tus observaciones aquí...")
        
    st.write("---") # Línea divisoria suave entre preguntas
    return respuesta, observacion

# --- INICIO DEL FORMULARIO ---
with st.form("rejilla_opt_form"):
    
    # --- ENCABEZADO ---
    st.header("Información General")
    c1, c2, c3, c4 = st.columns(4)
    with c1: fecha = st.date_input("Fecha", datetime.date.today())
    with c2: observador = st.text_input("Observador*")
    with c3: ute_equipo = st.text_input("UTE / Equipo")
    with c4: puesto_operario = st.text_input("Puesto / Operario*")
    st.divider()

    # --- 1. PREPARACIÓN DE LA OBSERVACIÓN ---
    st.header("1. Preparación de la Observación")
    q1_1, o1_1 = hacer_pregunta("1_1", "1.1. ¿Los estándares están al día y completos (FOS, Estado de Referencia 5S, TEO)?")
    q1_2, o1_2 = hacer_pregunta("1_2", "1.2. ¿Cuál es el nivel Skill del operario? ¿Ha seguido las formaciones en el TEO?", tipo="texto")
    q1_3, o1_3 = hacer_pregunta("1_3", "1.3. ¿La FSSE está al día (verificar vs última modificación)?")
    q1_4, o1_4 = hacer_pregunta("1_4", "1.4. ¿Ha sido identificado algún problema de Ergonomía o Seguridad? Si sí, ¿Cuál?", tipo="texto")
    q1_5, o1_5 = hacer_pregunta("1_5", "1.5. ¿Se tiene/tenía algún problema de calidad en el puesto recientemente? Si sí, ¿Cuál?", tipo="texto")
    q1_6, o1_6 = hacer_pregunta("1_6", "1.6. ¿Cuál es el indicador prioritario a mejorar en la UTE?", tipo="texto")
    q1_7, o1_7 = hacer_pregunta("1_7", "1.7. Filtro escogido (seguridad, calidad, costos, plazo…)", tipo="select", opciones=["Seguridad", "Calidad", "Costos", "Plazo", "Ergonomía", "Medio Ambiente"])

    # --- 2. OBSERVACIÓN DE RESPETO DE LOS ESTÁNDARES (DE LEJOS) ---
    st.header("2. Observación de respeto de los estándares - Observación de lejos")
    q2_1, o2_1 = hacer_pregunta("2_1", "2.1. ¿El operario tiene los EPP mencionados en las FOS A/P - Ficha de Seguridad?")
    q2_2, o2_2 = hacer_pregunta("2_2", "2.2. ¿El puesto está conforme con el estado de referencia 5S y comprende señalización CSR?")
    q2_3, o2_3 = hacer_pregunta("2_3", "2.3. ¿El operario respeta la FOS (Orden de las etapas principales)?")
    q2_4, o2_4 = hacer_pregunta("2_4", "2.4. ¿Las actividades no cíclicas (cambios, mantenimiento autónomo) se hacen conforme al estándar?")
    q2_5, o2_5 = hacer_pregunta("2_5", "2.5. ¿Las actividades de calidad frecuenciales (vigilancia, Poka Yoke) se hacen conforme al estándar?")

# --- 3. DIVERSIDAD (GRILLA DE 5 PIEZAS) ---
    st.header("3. Diversidad")
    st.write("Complete los valores para cada una de las piezas observadas (1 al 5) y añada comentarios.")

    # 1. Encabezados de la tabla (solo se imprimen una vez arriba)
    col_tit, c1, c2, c3, c4, c5, col_obs = st.columns([3, 1, 1, 1, 1, 1, 3])
    with col_tit: st.write("**Concepto**")
    with c1: st.write("**Pz 1**")
    with c2: st.write("**Pz 2**")
    with c3: st.write("**Pz 3**")
    with c4: st.write("**Pz 4**")
    with c5: st.write("**Pz 5**")
    with col_obs: st.write("**Respuestas y comentarios**")

    st.write("---") # Línea debajo de los encabezados

    # 2. Función para crear cada fila de la grilla
    def fila_grilla(id_item, nombre_item):
        # Mismas proporciones [3, 1, 1, 1, 1, 1, 3] para que todo quede alineado
        col_tit, c1, c2, c3, c4, c5, col_obs = st.columns([3, 1, 1, 1, 1, 1, 3])
        
        with col_tit:
            # Usamos HTML básico para que el texto baje un poco y se alinee con las cajas
            st.markdown(f"<div style='font-size: 14px; padding-top: 8px;'>{nombre_item}</div>", unsafe_allow_html=True)
            
        with c1: p1 = st.text_input("1", key=f"{id_item}_1", label_visibility="collapsed")
        with c2: p2 = st.text_input("2", key=f"{id_item}_2", label_visibility="collapsed")
        with c3: p3 = st.text_input("3", key=f"{id_item}_3", label_visibility="collapsed")
        with c4: p4 = st.text_input("4", key=f"{id_item}_4", label_visibility="collapsed")
        with c5: p5 = st.text_input("5", key=f"{id_item}_5", label_visibility="collapsed")
        
        with col_obs: 
            obs = st.text_input("Obs", key=f"{id_item}_obs", label_visibility="collapsed", placeholder="Comentarios...")
            
        # Retornamos los datos por si necesitas enviarlos a Google Sheets después
        return [p1, p2, p3, p4, p5, obs]

    # 3. Construimos las filas llamando a la función
    # Tiempos
    datos_3_1 = fila_grilla("3_1", "3.1. Tiempo Operatorio estándar (FOS)")
    datos_3_2 = fila_grilla("3_2", "3.2. Tiempo operatorio medido")
    datos_3_3 = fila_grilla("3_3", "3.3. Tiempo de actividades ok o no ok (+/-5%)")

    # Separador visual
    st.markdown("<br>", unsafe_allow_html=True) 

    # No Valor Agregado
    st.markdown("**3.4. No Valor Agregado**")
    datos_3_4_pasos = fila_grilla("3_4_pasos", "• Número de pasos")
    datos_3_4_tomar = fila_grilla("3_4_tomar", "• Tomar o depositar intermedio")
    datos_3_4_esperas = fila_grilla("3_4_esperas", "• Esperas")

    st.write("---")

    # --- 4. OBSERVACIÓN DEL RESPETO DEL ESTÁNDAR (DE CERCA) ---
    st.header("4. Observación del respeto del estándar - Observación de cerca")
    q4_1, o4_1 = hacer_pregunta("4_1", "4.1. Si una FOS A ha sido definida, ¿el operario respeta la que está ligada a un problema/defecto?")
    q4_2, o4_2 = hacer_pregunta("4_2", "4.2. ¿Los puntos clave son respetados y apropiados a los problemas del puesto?")
    q4_3, o4_3 = hacer_pregunta("4_3", "4.3. ¿El producto está conforme a lo requerido (Materias Primas & Producto terminado)?")
    q4_4, o4_4 = hacer_pregunta("4_4", "4.4. ¿Los embalajes, útiles, y ayudas escogidas son las del estado de referencia?")
    q4_5, o4_5 = hacer_pregunta("4_5", "4.5. ¿Las piezas son correctamente identificadas y los registros de trazabilidad son efectuados?")
    q4_6, o4_6 = hacer_pregunta("4_6", "4.6. ¿Los procedimientos de gestión de desechos y reglas de seguridad (químicos) son respetados?")

    # --- 5. MEJORA DEL ESTÁNDAR ---
    st.header("5. Observación para la mejora del estándar")
    q5_1, o5_1 = hacer_pregunta("5_1", "Identificar las mejoras y acciones a mediano plazo:", tipo="area")

    # --- 6. SÍNTESIS DE LA OBSERVACIÓN ---
    st.header("6. Síntesis de la Observación")
    q6_1, o6_1 = hacer_pregunta("6_1", "6.1. Intercambio con el operario sobre el respeto del estándar", tipo="area")
    q6_2, o6_2 = hacer_pregunta("6_2", "6.2. Verificar que el operario es capaz de nombrar etapas, prohibiciones y puntos CSR", tipo="radio", opciones=["Es capaz", "No es capaz / Dudoso"])
    q6_3, o6_3 = hacer_pregunta("6_3", "6.3. ¿Hay algún elemento que debe ser adjuntado al Cuadro de Control?")
    q6_4, o6_4 = hacer_pregunta("6_4", "6.4. Compartir las mejoras: (Del operario / Del observador)", tipo="area")
    q6_5, o6_5 = hacer_pregunta("6_5", "6.5. ¿Las mejoras pueden ser transversalizadas?")

    st.divider()

    # --- ANEXO: LIENZO DE DIBUJO ---
    st.header("Anexo Visual: Croquis de la Estación")
    c_dibujo1, c_dibujo2 = st.columns(2)
    with c_dibujo1: stroke_width = st.slider("Grosor del lápiz:", 1, 10, 2)
    with c_dibujo2: stroke_color = st.color_picker("Color del lápiz:", "#FF0000")
    
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

    # BOTÓN DE ENVÍO
    enviado = st.form_submit_button("Guardar Observación Completa", type="primary")

# --- LÓGICA DE PROCESAMIENTO ---
if enviado:
    if observador and puesto_operario:
        st.success("Formulario completado correctamente. Los datos y observaciones están listos para enviarse a Google Sheets.")
        
        # Ejemplo de cómo se vería la extracción de la Pregunta 1.1 y su observación
        st.write("**Ejemplo de captura de la primera pregunta:**")
        st.code(f"Pregunta 1.1: {q1_1} \nObservación 1.1: {o1_1}")
        
    else:
        st.error("⚠️ Por favor, completa los campos obligatorios (*).")
