import streamlit as st
import datetime
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from streamlit_drawable_canvas import st_canvas

# Librerías de Google
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Rejilla OPT Automatizada", layout="wide")

st.title("📋 Sistema de Registro: Rejilla de Observación de Puesto (OPT)")
st.write("Complete el formulario. Al enviar, se guardará en Drive y se actualizará la base de datos central.")

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================

def hacer_pregunta_estandar(id_pregunta, texto_pregunta, tipo="radio", opciones=["Sí", "No"]):
    col_preg, col_obs = st.columns([2, 1])
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
        observacion = st.text_area("Observaciones", key=f"obs_{id_pregunta}", height=68, label_visibility="collapsed", placeholder="Comentarios y desvíos...")
    st.write("---")
    return respuesta, observacion

def fila_grilla_interactiva(id_item, texto_item):
    g_tit, g1, g2, g3, g4, g5, g_obs = st.columns([3, 1, 1, 1, 1, 1, 3])
    with g_tit: st.markdown(f"<div style='padding-top: 8px;'>{texto_item}</div>", unsafe_allow_html=True)
    with g1: p1 = st.text_input("1", key=f"g_{id_item}_1", label_visibility="collapsed")
    with g2: p2 = st.text_input("2", key=f"g_{id_item}_2", label_visibility="collapsed")
    with g3: p3 = st.text_input("3", key=f"g_{id_item}_3", label_visibility="collapsed")
    with g4: p4 = st.text_input("4", key=f"g_{id_item}_4", label_visibility="collapsed")
    with g5: p5 = st.text_input("5", key=f"g_{id_item}_5", label_visibility="collapsed")
    with g_obs: obs = st.text_input("Obs", key=f"g_{id_item}_obs", label_visibility="collapsed", placeholder="Notas de tiempo...")
    st.write("---")
    return [p1, p2, p3, p4, p5], obs

def generar_excel_profesional(datos_form):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultado OPT"
    ws.views.sheetView[0].showGridLines = True
    
    HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    SECTION_FILL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    GRAY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    ws["A1"] = "REJILLA DE OBSERVACIÓN DE PUESTO (OPT) - RESULTADOS"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F497D")
    
    ws["A3"] = "Fecha:"; ws["B3"] = datos_form['Encabezado']['Fecha']
    ws["A4"] = "Observador:"; ws["B4"] = datos_form['Encabezado']['Observador']
    ws["D3"] = "UTE / Equipo:"; ws["E3"] = datos_form['Encabezado']['UTE']
    ws["D4"] = "Puesto / Operario:"; ws["E4"] = datos_form['Encabezado']['Operario']
    
    headers = ["ID", "Criterio / Pregunta de la Observación", "Pz 1 / Resp", "Pz 2", "Pz 3", "Pz 4", "Pz 5", "Observaciones"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=col_idx, value=h)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        
    r = 7
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    
    for seccion, preguntas in datos_form['Contenido'].items():
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        s_cell = ws.cell(row=r, column=1, value=seccion)
        s_cell.font = Font(name="Arial", size=11, bold=True, color="1F497D")
        s_cell.fill = SECTION_FILL
        r += 1
        
        for q_id, q_data in preguntas.items():
            ws.cell(row=r, column=1, value=q_id).font = Font(bold=True)
            ws.cell(row=r, column=2, value=q_data['texto'])
            
            if q_data['es_grilla']:
                for p_idx in range(5):
                    ws.cell(row=r, column=3 + p_idx, value=q_data['valores'][p_idx]).alignment = Alignment(horizontal="center")
            else:
                ws.cell(row=r, column=3, value=q_data['valores']).alignment = Alignment(horizontal="center")
                ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
                for c in range(4, 8):
                    ws.cell(row=r, column=c).fill = GRAY_FILL
                    
            ws.cell(row=r, column=8, value=q_data['observacion'])
            for col in range(1, 9):
                ws.cell(row=r, column=col).border = thin_border
            r += 1
            
    ws.column_dimensions["A"].width = 8; ws.column_dimensions["B"].width = 45
    for c in ["C", "D", "E", "F", "G"]: ws.column_dimensions[c].width = 10
    ws.column_dimensions["H"].width = 35
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def subir_excel_a_drive(excel_bytes, nombre_archivo, folder_id):
    credenciales_dict = json.loads(st.secrets["google_credentials"])
    SCOPES = ['https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_info(credenciales_dict, scopes=SCOPES)
    
    servicio_drive = build('drive', 'v3', credentials=creds)
    archivo_stream = io.BytesIO(excel_bytes)
    media = MediaIoBaseUpload(archivo_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
    
    metadatos_archivo = {'name': nombre_archivo, 'parents': [folder_id]}
    
    # supportsAllDrives=True previene el error 403 en Unidades Compartidas
    archivo_subido = servicio_drive.files().create(
        body=metadatos_archivo, 
        media_body=media, 
        fields='id, webViewLink',
        supportsAllDrives=True
    ).execute()
    
    return archivo_subido.get('webViewLink')

# ==========================================
# INTERFAZ WEB DEL FORMULARIO STREAMLIT
# ==========================================
with st.form("formulario_opt_avanzado"):
    st.header("1. Información del Relevamiento")
    c1, c2, c3, c4 = st.columns(4)
    with c1: fecha = st.date_input("Fecha", datetime.date.today())
    with c2: observador = st.text_input("Observador*")
    with c3: ute_equipo = st.text_input("UTE / Equipo")
    with c4: puesto_operario = st.text_input("Puesto / Operario*")
    st.divider()

    st.header("1. Preparación de la Observación")
    r1_1, o1_1 = hacer_pregunta_estandar("1_1", "1.1. ¿Los estándares están al día y completos (FOS, 5S, TEO)?")
    r1_2, o1_2 = hacer_pregunta_estandar("1_2", "1.2. Nivel Skill del operario / Formaciones TEO", tipo="texto")
    r1_4, o1_4 = hacer_pregunta_estandar("1_4", "1.4. ¿Problemas de Ergonomía o Seguridad detectados?", tipo="texto")
    r1_7, o1_7 = hacer_pregunta_estandar("1_7", "1.7. Filtro escogido para esta OPT", tipo="select", opciones=["Seguridad", "Calidad", "Costos", "Plazo", "Ergonomía"])

    st.header("2. Observación de respeto de los estándares")
    r2_1, o2_1 = hacer_pregunta_estandar("2_1", "2.1. ¿El operario cuenta y utiliza los EPP obligatorios?")
    r2_2, o2_2 = hacer_pregunta_estandar("2_2", "2.2. ¿El puesto cumple con el estándar 5S y señalización CSR?")

    st.header("3. Diversidad (Medición por Pieza)")
    g_tit, g1, g2, g3, g4, g5, g_obs = st.columns([3, 1, 1, 1, 1, 1, 3])
    with g_tit: st.write("**Criterio de Tiempos**")
    with g1: st.write("**Pz 1**")
    with g2: st.write("**Pz 2**")
    with g3: st.write("**Pz 3**")
    with g4: st.write("**Pz 4**")
    with g5: st.write("**Pz 5**")
    with g_obs: st.write("**Comentarios**")
    st.write("---")

    v3_1, obs3_1 = fila_grilla_interactiva("3_1", "3.1. Tiempo Operatorio estándar (FOS)")
    v3_2, obs3_2 = fila_grilla_interactiva("3_2", "3.2. Tiempo operatorio medido")
    v3_3, obs3_3 = fila_grilla_interactiva("3_3", "3.3. Tiempo de actividades ok o no ok (+/-5%)")
    
    st.write("**3.4. No Valor Agregado**")
    v3_4a, obs3_4a = fila_grilla_interactiva("3_4a", "• Número de pasos")
    v3_4b, obs3_4b = fila_grilla_interactiva("3_4b", "• Tomar o depositar intermedio")
    v3_4c, obs3_4c = fila_grilla_interactiva("3_4c", "• Esperas")

    st.header("4 a 6. Síntesis y Evaluación de Cerca")
    r4_1, o4_1 = hacer_pregunta_estandar("4_1", "4.1. Si aplica FOS A (Calidad/Defectos), ¿se respeta rigurosamente?")
    r4_2, o4_2 = hacer_pregunta_estandar("4_2", "4.2. ¿Los puntos clave de seguridad y calidad son comprendidos?")
    r5_1, o5_1 = hacer_pregunta_estandar("5_1", "5.1. Propuestas de mejora para el estándar", tipo="area")
    r6_4, o6_4 = hacer_pregunta_estandar("6_4", "6.4. Ideas de mejora (co-creadas Operario + Observador)", tipo="area")

    st.header("🎨 Croquis Opcional del Puesto")
    canvas_result = st_canvas(fill_color="rgba(255, 165, 0, 0.3)", stroke_width=2, stroke_color="#FF0000", background_color="#f0f2f6", height=200, width=700, drawing_mode="freedraw", key="canvas_opt")

    btn_enviar = st.form_submit_button("🚀 Procesar Respuestas", type="primary")

# ==========================================
# PROCESAMIENTO AL ENVIAR
# ==========================================
if btn_enviar:
    if not observador or not puesto_operario:
        st.error("❌ Por favor complete los campos obligatorios ('Observador' y 'Puesto / Operario').")
    else:
        # 1. Empaquetar Datos para el Excel
        datos_totales = {
            'Encabezado': {'Fecha': str(fecha), 'Observador': observador, 'UTE': ute_equipo, 'Operario': puesto_operario},
            'Contenido': {
                '1. Preparación': {
                    '1.1': {'texto': 'Estándares al día y completos', 'es_grilla': False, 'valores': r1_1, 'observacion': o1_1},
                    '1.2': {'texto': 'Nivel Skill / Formaciones', 'es_grilla': False, 'valores': r1_2, 'observacion': o1_2},
                    '1.4': {'texto': 'Ergonomía/Seguridad', 'es_grilla': False, 'valores': r1_4, 'observacion': o1_4},
                    '1.7': {'texto': 'Filtro seleccionado', 'es_grilla': False, 'valores': r1_7, 'observacion': o1_7},
                },
                '2. Observación de Lejos': {
                    '2.1': {'texto': 'Uso de EPP obligatorios', 'es_grilla': False, 'valores': r2_1, 'observacion': o2_1},
                    '2.2': {'texto': 'Cumplimiento 5S / CSR', 'es_grilla': False, 'valores': r2_2, 'observacion': o2_2},
                },
                '3. Diversidad (Grilla por Piezas)': {
                    '3.1': {'texto': 'Tiempo Estándar FOS', 'es_grilla': True, 'valores': v3_1, 'observacion': obs3_1},
                    '3.2': {'texto': 'Tiempo Medido', 'es_grilla': True, 'valores': v3_2, 'observacion': obs3_2},
                    '3.3': {'texto': 'Tiempos +/-5%', 'es_grilla': True, 'valores': v3_3, 'observacion': obs3_3},
                    '3.4a': {'texto': '• Número de pasos', 'es_grilla': True, 'valores': v3_4a, 'observacion': obs3_4a},
                    '3.4b': {'texto': '• Tomar/Depositar', 'es_grilla': True, 'valores': v3_4b, 'observacion': obs3_4b},
                    '3.4c': {'texto': '• Esperas', 'es_grilla': True, 'valores': v3_4c, 'observacion': obs3_4c},
                },
                '4 a 6. Síntesis y Evaluación de Cerca': {
                    '4.1': {'texto': 'Respeto de FOS A (Calidad)', 'es_grilla': False, 'valores': r4_1, 'observacion': o4_1},
                    '4.2': {'texto': 'Puntos clave comprendidos', 'es_grilla': False, 'valores': r4_2, 'observacion': o4_2},
                    '5.1': {'texto': 'Propuestas para Plan UTE/LUP', 'es_grilla': False, 'valores': r5_1, 'observacion': o5_1},
                    '6.4': {'texto': 'Ideas de mejora cocreadas', 'es_grilla': False, 'valores': r6_4, 'observacion': o6_4},
                }
            }
        }
        
        # 2. Compilar Excel y subir a Google Drive
        excel_bytes = generar_excel_profesional(datos_totales)
        nombre_excel = f"OPT_{puesto_operario.replace(' ', '_')}_{fecha}.xlsx"
        ID_CARPETA_DRIVE = "1Io4SlOlxISQIA0Jrriiz5LF5fWEYM5Ys" 
        
        link_del_excel = "No se pudo subir" # Valor por defecto si Drive falla
        
        try:
            link_del_excel = subir_excel_a_drive(excel_bytes, nombre_excel, ID_CARPETA_DRIVE)
            st.success("📁 ¡El Excel se guardó automáticamente en Google Drive!")
            st.markdown(f"[🔗 Haz clic aquí para ver el Excel en la nube]({link_del_excel})")
        except Exception as e:
            st.error(f"Error al subir a Drive: {e}")
            st.download_button("📥 Descargar Excel Manualmente", data=excel_bytes, file_name=nombre_excel)

        # 3. Preparar la fila completa estilo "Google Forms"
        def formato_pz(lista_pz):
            return " | ".join([p if p else "-" for p in lista_pz])

        fila_completa_sheets = [
            str(fecha), observador, ute_equipo, puesto_operario,
            r1_1, o1_1, r1_2, o1_2, r1_4, o1_4, r1_7, o1_7,
            r2_1, o2_1, r2_2, o2_2,
            formato_pz(v3_1), obs3_1, 
            formato_pz(v3_2), obs3_2, 
            formato_pz(v3_3), obs3_3,
            formato_pz(v3_4a), obs3_4a, 
            formato_pz(v3_4b), obs3_4b, 
            formato_pz(v3_4c), obs3_4c,
            r4_1, o4_1, r4_2, o4_2,
            r5_1, o5_1, r6_4, o6_4,
            link_del_excel  # El link siempre al final
        ]
        
        # 4. Enviar datos a Google Sheets
        try:
            credenciales_dict = json.loads(st.secrets["google_credentials"])
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_dict(credenciales_dict, scope)
            client = gspread.authorize(creds)
            
            sheet = client.open_by_key("1S-N8Ok_Q0NmYqQSzKUBpXgW3OvcjnNkrNBsrL5Xd4dY").sheet1 
            sheet.append_row(fila_completa_sheets)
            st.success("☁️ ¡Base de datos central actualizada con todas las respuestas estilo Forms!")
            st.balloons()
        except Exception as e:
            st.error(f"Error en Sheets: {e}")
