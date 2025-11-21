import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
import os
from datetime import datetime
import zipfile

def insert_text_in_pdf(pdf_bytes, aconex_text, sistema_text, subsistema_text, position_settings):
    """Inserta texto en un PDF en las posiciones específicas con ajustes precisos"""
    try:
        # Abrir el PDF
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Obtener ajustes de posición
        x_offset = position_settings['x_offset']
        y_offset = position_settings['y_offset']
        font_size = position_settings['font_size']
        
        # Procesar cada página
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # Función para buscar texto case insensitive
            def find_text_position(page, target_text):
                areas = page.search_for(target_text)
                if not areas:
                    areas = page.search_for(target_text.upper())
                if not areas:
                    areas = page.search_for(target_text.lower())
                return areas
            
            # Buscar todas las instancias de las palabras clave
            aconex_instances = find_text_position(page, "ACONEX")
            sistema_instances = find_text_position(page, "SISTEMA")
            subsistema_instances = find_text_position(page, "SUB SISTEMA")
            subsistema_instances2 = find_text_position(page, "SUB SISTEMA")
            subsistema_instances_alt = find_text_position(page, "SUB SISTEMA :")
            subsistema_instances_alt2 = find_text_position(page, "SUBSISTEMA :")
            
            # Combinar todas las instancias de subsistema
            all_subsistema_instances = subsistema_instances + subsistema_instances_alt + subsistema_instances_alt2 + subsistema_instances2
            
            # Insertar texto después de ACONEX
            if aconex_instances and aconex_text:
                rect = aconex_instances[0]
                x_position = rect.x1 + x_offset
                y_position = rect.y0 + (rect.height / 2) + y_offset
                page.insert_text(
                    (x_position, y_position),
                    str(aconex_text),
                    fontsize=font_size,
                    fontname="helv",
                    color=(0, 0, 0)
                )
            
            # Insertar texto después de SISTEMA
            if sistema_instances and sistema_text:
                rect = sistema_instances[0]
                x_position = rect.x1 + x_offset
                y_position = rect.y0 + (rect.height / 2) + y_offset
                page.insert_text(
                    (x_position, y_position),
                    str(sistema_text),
                    fontsize=font_size,
                    fontname="helv",
                    color=(0, 0, 0)
                )
            
            # Insertar texto después de SUB SISTEMA
            if all_subsistema_instances and subsistema_text:
                rect = all_subsistema_instances[0]
                x_position = rect.x1 + x_offset
                y_position = rect.y0 + (rect.height / 2) + y_offset
                page.insert_text(
                    (x_position, y_position),
                    str(subsistema_text),
                    fontsize=font_size,
                    fontname="helv",
                    color=(0, 0, 0)
                )
        
        # Guardar el PDF modificado en memoria
        output_bytes = pdf_document.tobytes()
        pdf_document.close()
        
        return output_bytes
    
    except Exception as e:
        st.error(f"Error al procesar el PDF: {str(e)}")
        return None

def insert_text_precise(pdf_bytes, text_data, position_settings):
    """Versión avanzada con posicionamiento preciso por coordenadas específicas"""
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Obtener ajustes
        font_size = position_settings['font_size']
        custom_positions = position_settings.get('custom_positions', {})
        
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # Usar posiciones personalizadas si están definidas
            if custom_positions:
                for field, pos in custom_positions.items():
                    if field in text_data and text_data[field]:
                        x, y = pos['x'], pos['y']
                        page.insert_text(
                            (x, y),
                            str(text_data[field]),
                            fontsize=font_size,
                            fontname="helv",
                            color=(0, 0, 0)
                        )
            else:
                # Búsqueda automática con offsets
                def find_and_insert(field_name, search_text, text_value):
                    if not text_value:
                        return
                    
                    instances = page.search_for(search_text)
                    if not instances:
                        instances = page.search_for(search_text.upper())
                    if not instances:
                        instances = page.search_for(search_text.lower())
                    
                    if instances:
                        rect = instances[0]
                        x_position = rect.x1 + position_settings['x_offset']
                        y_position = rect.y0 + (rect.height / 2) + position_settings['y_offset']
                        
                        page.insert_text(
                            (x_position, y_position),
                            str(text_value),
                            fontsize=font_size,
                            fontname="helv",
                            color=(0, 0, 0)
                        )
                
                # Insertar cada campo
                find_and_insert("ACONEX", "ACONEX", text_data.get('aconex'))
                find_and_insert("SISTEMA", "SISTEMA", text_data.get('sistema'))
                find_and_insert("SUB SISTEMA", "SUB SISTEMA", text_data.get('subsistema'))
        
        output_bytes = pdf_document.tobytes()
        pdf_document.close()
        return output_bytes
        
    except Exception as e:
        st.error(f"Error en inserción precisa: {str(e)}")
        return None

def analyze_pdf_coordinates(pdf_bytes, sample_texts):
    """Analiza un PDF y sugiere coordenadas para inserción"""
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        analysis_results = []
        
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            page_results = {'page': page_num + 1, 'fields': {}}
            
            for field_name, search_text in sample_texts.items():
                instances = page.search_for(search_text)
                if instances:
                    rect = instances[0]
                    page_results['fields'][field_name] = {
                        'found': True,
                        'x': rect.x0,
                        'y': rect.y0,
                        'width': rect.width,
                        'height': rect.height,
                        'suggested_x': rect.x1 + 15,  # 15 puntos a la derecha
                        'suggested_y': rect.y0 + (rect.height / 2) - 2  # Centrado vertical
                    }
                else:
                    page_results['fields'][field_name] = {'found': False}
            
            analysis_results.append(page_results)
        
        pdf_document.close()
        return analysis_results
        
    except Exception as e:
        st.error(f"Error analizando coordenadas: {str(e)}")
        return []

def validate_dataframe(df):
    """Valida que el DataFrame tenga las columnas necesarias"""
    required_columns = ['ACONEX', 'SISTEMA', 'SUB SISTEMA']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        st.error(f"Faltan columnas en el Excel: {', '.join(missing_columns)}")
        st.info("Las columnas requeridas son: ACONEX, SISTEMA, SUB SISTEMA")
        return False
    
    return True

def get_position_settings():
    """Obtiene la configuración de posición del usuario"""
    st.sidebar.header("🎯 Ajustes de Posición")
    
    # Método de posicionamiento
    positioning_method = st.sidebar.radio(
        "Método de posicionamiento:",
        ["Automático con offsets", "Coordenadas manuales"]
    )
    
    position_settings = {}
    
    if positioning_method == "Automático con offsets":
        col1, col2, col3 = st.sidebar.columns(3)
        
        with col1:
            position_settings['x_offset'] = st.number_input(
                "Offset X (puntos)",
                min_value=-100.0,
                max_value=100.0,
                value=15.0,
                step=1.0,
                help="Desplazamiento horizontal desde la etiqueta"
            )
        
        with col2:
            position_settings['y_offset'] = st.number_input(
                "Offset Y (puntos)",
                min_value=-50.0,
                max_value=50.0,
                value=2.0,
                step=0.5,
                help="Desplazamiento vertical desde la etiqueta"
            )
        
        with col3:
            position_settings['font_size'] = st.number_input(
                "Tamaño fuente",
                min_value=6.0,
                max_value=20.0,
                value=7.0,
                step=0.5,
                help="Tamaño de la fuente en puntos"
            )
    
    else:  # Coordenadas manuales
        st.sidebar.info("💡 Use el análisis de coordenadas primero para obtener las posiciones")
        
        position_settings['font_size'] = st.sidebar.number_input(
            "Tamaño fuente",
            min_value=6.0,
            max_value=20.0,
            value=9.0,
            step=0.5
        )
        
        position_settings['custom_positions'] = {}
        
        st.sidebar.subheader("Coordenadas por campo:")
        
        # ACONEX
        col1, col2 = st.sidebar.columns(2)
        with col1:
            aconex_x = st.number_input("ACONEX X", value=100.0, step=1.0)
        with col2:
            aconex_y = st.number_input("ACONEX Y", value=100.0, step=1.0)
        position_settings['custom_positions']['aconex'] = {'x': aconex_x, 'y': aconex_y}
        
        # SISTEMA
        col1, col2 = st.sidebar.columns(2)
        with col1:
            sistema_x = st.number_input("SISTEMA X", value=100.0, step=1.0)
        with col2:
            sistema_y = st.number_input("SISTEMA Y", value=120.0, step=1.0)
        position_settings['custom_positions']['sistema'] = {'x': sistema_x, 'y': sistema_y}
        
        # SUB SISTEMA
        col1, col2 = st.sidebar.columns(2)
        with col1:
            subsistema_x = st.number_input("SUB SISTEMA X", value=100.0, step=1.0)
        with col2:
            subsistema_y = st.number_input("SUB SISTEMA Y", value=140.0, step=1.0)
        position_settings['custom_positions']['subsistema'] = {'x': subsistema_x, 'y': subsistema_y}
    
    # 🔥 NUEVA SECCIÓN: Control de zoom para vista previa
    st.sidebar.header("🔍 Ajustes de Vista Previa")
    zoom_level = st.sidebar.slider(
        "Nivel de zoom:",
        min_value=1.0,
        max_value=3.0,
        value=2.0,
        step=0.5,
        help="Ajusta el nivel de zoom para la vista previa de PDF"
    )
    position_settings['zoom_level'] = zoom_level
    
    return position_settings

def create_download_zip(processed_files):
    """Crea un archivo ZIP en memoria para descargar"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for file_info in processed_files:
            # Agregar archivo al ZIP desde bytes
            zip_file.writestr(file_info['filename'], file_info['bytes'])
    
    zip_buffer.seek(0)
    return zip_buffer

def display_pdf_preview(pdf_bytes, width=1000, zoom=2.0):
    """Muestra una vista previa rápida del PDF como imagen (solo primera página)"""
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Convertir primera página a imagen
        page = pdf_document[0]
        mat = fitz.Matrix(zoom, zoom)  # 🔥 ZOOM AJUSTABLE
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        
        pdf_document.close()
        
        # Mostrar imagen
        st.image(img_data, caption=f"Vista previa rápida - Zoom: {zoom}x", width=width)
        return True
        
    except Exception as e:
        st.error(f"Error generando vista previa: {str(e)}")
        return False

def show_quick_preview(processed_files, zoom_level=2.0):
    """Muestra una vista rápida de los archivos procesados con navegación integrada"""
    st.header("👀 Vista Rápida de Resultados")
    
    if not processed_files:
        st.warning("No hay archivos procesados para mostrar")
        return
    
    # Inicializar el índice actual en session_state si no existe
    if 'current_preview_index' not in st.session_state:
        st.session_state.current_preview_index = 0
    
    # Asegurarse de que el índice esté dentro de los límites
    if st.session_state.current_preview_index >= len(processed_files):
        st.session_state.current_preview_index = 0
    
    # Obtener el archivo actual
    current_index = st.session_state.current_preview_index
    current_file = processed_files[current_index]
    
    # Crear dos columnas: una para la vista previa y otra para la información/navegación
    col_preview, col_info = st.columns([2, 1])
    
    with col_preview:
        # Vista previa rápida CON ZOOM
        st.markdown(f"**📄 {current_file['filename']}**")
        if display_pdf_preview(current_file['bytes'], width=500, zoom=zoom_level):
            st.success("✅ Vista previa generada correctamente")
    
    with col_info:
        # Información del archivo actual
        st.subheader("📋 Información")
        st.info(f"**ACONEX:**\n{current_file['aconex']}")
        st.info(f"**SISTEMA:**\n{current_file['sistema']}")
        st.info(f"**SUB SISTEMA:**\n{current_file['subsistema']}")
        
        # Navegación compacta
        st.subheader("🧭 Navegación")
        
        # Indicador de posición
        st.markdown(f"**Archivo {current_index + 1} de {len(processed_files)}**")
        
        # Botones de navegación en una cuadrícula compacta
        nav_col1, nav_col2, nav_col3, nav_col4 = st.columns(4)
        
        with nav_col1:
            if st.button("⏮️", help="Primer archivo", use_container_width=True):
                st.session_state.current_preview_index = 0
                st.rerun()
        
        with nav_col2:
            if st.button("◀️", help="Archivo anterior", use_container_width=True, 
                        disabled=(current_index == 0)):
                if current_index > 0:
                    st.session_state.current_preview_index -= 1
                    st.rerun()
        
        with nav_col3:
            if st.button("▶️", help="Siguiente archivo", use_container_width=True,
                        disabled=(current_index == len(processed_files) - 1)):
                if current_index < len(processed_files) - 1:
                    st.session_state.current_preview_index += 1
                    st.rerun()
        
        with nav_col4:
            if st.button("⏭️", help="Último archivo", use_container_width=True,
                        disabled=(current_index == len(processed_files) - 1)):
                st.session_state.current_preview_index = len(processed_files) - 1
                st.rerun()
        
        # Selector desplegable para navegación rápida
        st.markdown("---")
        file_names = [f['filename'] for f in processed_files]
        selected_file = st.selectbox(
            "Ir a archivo específico:",
            file_names,
            index=current_index,
            key="quick_nav_select"
        )
        
        # Si el usuario selecciona un archivo diferente, actualizar el índice
        if selected_file != file_names[current_index]:
            new_index = file_names.index(selected_file)
            st.session_state.current_preview_index = new_index
            st.rerun()
        
        # Descarga rápida del archivo actual
        st.markdown("---")
        st.download_button(
            label="📥 Descargar este archivo",
            data=current_file['bytes'],
            file_name=current_file['filename'],
            mime="application/pdf",
            use_container_width=True,
            key=f"quick_dl_{current_index}"
        )

def process_files(matched_files, position_settings):
    """Procesa los archivos y los guarda en session_state"""
    st.header("📊 Procesando Archivos...")
    
    progress_bar = st.progress(0)
    processed_files = []
    errors = []
    
    for i, match_info in enumerate(matched_files):
        progress = (i + 1) / len(matched_files)
        progress_bar.progress(progress)
        
        try:
            pdf_bytes = match_info['file'].getvalue()
            
            # Preparar datos de texto
            text_data = {
                'aconex': match_info['aconex'],
                'sistema': match_info['sistema'],
                'subsistema': match_info['subsistema']
            }
            
            # Usar inserción precisa si hay coordenadas personalizadas
            if 'custom_positions' in position_settings:
                modified_pdf_bytes = insert_text_precise(pdf_bytes, text_data, position_settings)
            else:
                modified_pdf_bytes = insert_text_in_pdf(
                    pdf_bytes, 
                    match_info['aconex'], 
                    match_info['sistema'], 
                    match_info['subsistema'],
                    position_settings
                )
            
            if modified_pdf_bytes:
                processed_files.append({
                    'filename': match_info['name'],
                    'aconex': match_info['aconex'],
                    'sistema': match_info['sistema'],
                    'subsistema': match_info['subsistema'],
                    'bytes': modified_pdf_bytes
                })
                
                st.sidebar.success(f"✅ {match_info['name']}")
            else:
                errors.append(f"{match_info['name']}")
                
        except Exception as e:
            errors.append(f"{match_info['name']}: {str(e)}")
    
    # Guardar en session_state
    st.session_state.processed_files = processed_files
    st.session_state.processing_errors = errors
    st.session_state.processing_complete = True
    
    return processed_files, errors

def main():
    st.set_page_config(
        page_title="Procesador Masivo de PDFs",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🎯 Procesador Masivo de PDFs")
    st.markdown("""
    **Procesa múltiples archivos PDF y verifica los resultados antes de descargar.**
    """)
    
    # Inicializar session_state
    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = []
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'processing_errors' not in st.session_state:
        st.session_state.processing_errors = []
    
    # Configuraciones del sidebar
    position_settings = get_position_settings()
    
    # Cargar archivo Excel
    st.header("1. 📊 Cargar Archivo Excel con Datos")
    excel_file = st.file_uploader(
        "Sube el archivo Listado.xlsx", 
        type=["xlsx"],
        help="Debe contener las columnas: ACONEX, SISTEMA, SUB SISTEMA",
        key="excel_uploader"
    )
    
    df = None
    if excel_file is not None:
        try:
            df = pd.read_excel(excel_file)
            st.success(f"✅ Excel cargado - {len(df)} registros encontrados")
            
            if validate_dataframe(df):
                st.dataframe(df.head(10), use_container_width=True)
                
                # Mostrar estadísticas básicas
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total registros", len(df))
                with col2:
                    st.metric("Con SISTEMA", df['SISTEMA'].notna().sum())
                with col3:
                    st.metric("Con SUB SISTEMA", df['SUB SISTEMA'].notna().sum())
    
        except Exception as e:
            st.error(f"❌ Error al cargar el Excel: {str(e)}")
    
    # Cargar múltiples archivos PDF
    st.header("2. 📁 Cargar Archivos PDF Originales")
    
    pdf_files = st.file_uploader(
        "Selecciona múltiples archivos PDF",
        type=["pdf"],
        accept_multiple_files=True,
        help="Selecciona todos los PDFs que quieres procesar",
        key="pdf_uploader"
    )
    
    if pdf_files and df is not None and validate_dataframe(df):
        st.success(f"✅ {len(pdf_files)} archivos PDF cargados")
        
        # Análisis avanzado de coordenadas (opcional)
        if st.checkbox("🔍 Análisis Avanzado de Coordenadas", help="Opcional: Analiza las posiciones exactas en el PDF"):
            st.subheader("Análisis de Coordenadas del Primer PDF")
            
            sample_texts = {
                'ACONEX': 'ACONEX',
                'SISTEMA': 'SISTEMA', 
                'SUB SISTEMA': 'SUB SISTEMA'
            }
            
            first_pdf_bytes = pdf_files[0].getvalue()
            coordinate_analysis = analyze_pdf_coordinates(first_pdf_bytes, sample_texts)
            
            if coordinate_analysis:
                for page_analysis in coordinate_analysis:
                    st.write(f"**Página {page_analysis['page']}:**")
                    
                    for field_name, field_data in page_analysis['fields'].items():
                        if field_data.get('found'):
                            st.write(f"  - **{field_name}:** Encontrado en ({field_data['x']:.1f}, {field_data['y']:.1f})")
                        else:
                            st.write(f"  - **{field_name}:** ❌ No encontrado")
            
            # Reset del archivo
            pdf_files[0].seek(0)
        
        # Crear diccionario de datos del Excel
        excel_data = {}
        for _, row in df.iterrows():
            aconex = row['ACONEX']
            if pd.notna(aconex):
                excel_data[str(aconex).strip()] = {
                    'sistema': row['SISTEMA'] if pd.notna(row['SISTEMA']) else "",
                    'subsistema': row['SUB SISTEMA'] if pd.notna(row['SUB SISTEMA']) else ""
                }
        
        # Emparejamiento de archivos
        st.header("3. 🔍 Emparejamiento de Archivos")
        
        matched_files = []
        unmatched_files = []
        
        for pdf_file in pdf_files:
            pdf_name = pdf_file.name
            base_name = os.path.splitext(pdf_name)[0]
            
            if base_name in excel_data:
                data = excel_data[base_name]
                matched_files.append({
                    'file': pdf_file,
                    'name': pdf_name,
                    'aconex': base_name,
                    'sistema': data['sistema'],
                    'subsistema': data['subsistema']
                })
            else:
                unmatched_files.append(pdf_name)
        
        # Mostrar resultados del emparejamiento
        col1, col2 = st.columns(2)
        
        with col1:
            st.success(f"✅ {len(matched_files)} archivos emparejados")
        
        with col2:
            if unmatched_files:
                st.warning(f"⚠️ {len(unmatched_files)} no emparejados")
        
        # Procesar archivos
        if matched_files:
            st.header("4. ⚙️ Procesar Archivos")
            
            # Mostrar configuración actual de forma simple
            st.info("**Configuración actual:**")
            if 'custom_positions' in position_settings:
                st.write("Usando coordenadas manuales")
            else:
                st.write(f"Offsets: X={position_settings['x_offset']}, Y={position_settings['y_offset']}")
            st.write(f"Tamaño de fuente: {position_settings['font_size']}")
            st.write(f"Zoom vista previa: {position_settings.get('zoom_level', 2.0)}x")
            
            # Botón de procesamiento
            if st.button("🚀 Iniciar Procesamiento", type="primary", key="process_btn"):
                # Reiniciar el estado de vista previa
                if 'current_preview_index' in st.session_state:
                    st.session_state.current_preview_index = 0
                
                processed_files, errors = process_files(matched_files, position_settings)
            
            # Mostrar resultados si el procesamiento está completo
            if st.session_state.processing_complete and st.session_state.processed_files:
                st.header("5. ✅ Resultados Finales")
                
                processed_files = st.session_state.processed_files
                errors = st.session_state.processing_errors
                
                st.success(f"✅ {len(processed_files)} archivos procesados exitosamente")
                
                # Vista rápida con navegación integrada Y ZOOM
                show_quick_preview(processed_files, zoom_level=position_settings.get('zoom_level', 2.0))
                
                # Descargas masivas
                st.header("6. 📥 Descargar Todos los Archivos")
                
                # ZIP con todos los archivos
                zip_buffer = create_download_zip(processed_files)
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                st.download_button(
                    label=f"📦 Descargar ZIP completo ({len(processed_files)} archivos)",
                    data=zip_buffer,
                    file_name=f"pdfs_procesados_{timestamp}.zip",
                    mime="application/zip",
                    type="primary",
                    key="download_zip",
                    use_container_width=True
                )
                
                if errors:
                    st.error(f"❌ {len(errors)} errores:")
                    for error in errors:
                        st.write(f"- {error}")
            
            # Botón para resetear si es necesario
            if st.session_state.processing_complete:
                if st.button("🔄 Procesar Nuevos Archivos", type="secondary"):
                    # Limpiar session_state
                    for key in ['processed_files', 'processing_complete', 'processing_errors', 'current_preview_index']:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.rerun()

if __name__ == "__main__":
    main()