import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
import os
from datetime import datetime
import tempfile
import zipfile
import shutil

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
        
        # Guardar el PDF modificado
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

def select_output_folder():
    """Permite al usuario seleccionar la carpeta de destino"""
    st.sidebar.header("📁 Configuración de Salida")
    
    output_option = st.sidebar.radio(
        "¿Dónde guardar los archivos procesados?",
        ["Carpeta temporal (descarga ZIP)", "Seleccionar carpeta específica"]
    )
    
    if output_option == "Seleccionar carpeta específica":
        folder_path = st.sidebar.text_input(
            "Ruta de la carpeta de destino:",
            value="",
            placeholder="Ej: C:/Usuarios/MiUsuario/Documentos/PDFs_Procesados"
        )
        
        if folder_path:
            if not os.path.exists(folder_path):
                st.sidebar.info(f"La carpeta no existe. Se creará: {folder_path}")
                try:
                    os.makedirs(folder_path, exist_ok=True)
                    st.sidebar.success("✅ Carpeta creada exitosamente")
                except Exception as e:
                    st.sidebar.error(f"❌ Error creando carpeta: {e}")
                    return None
            
            try:
                test_file = os.path.join(folder_path, "test_write.tmp")
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                st.sidebar.success("✅ Permisos de escritura verificados")
                return folder_path
            except Exception as e:
                st.sidebar.error(f"❌ Sin permisos de escritura en: {folder_path}")
                return None
        else:
            return None
    else:
        return None

def save_to_folder(processed_files, output_folder):
    """Guarda los archivos procesados en la carpeta especificada"""
    try:
        saved_files = []
        
        for file_info in processed_files:
            source_path = file_info['filepath']
            destination_path = os.path.join(output_folder, file_info['filename'])
            
            shutil.copy2(source_path, destination_path)
            
            saved_files.append({
                'filename': file_info['filename'],
                'path': destination_path,
                'aconex': file_info['aconex'],
                'sistema': file_info['sistema'],
                'subsistema': file_info['subsistema']
            })
        
        return saved_files
    
    except Exception as e:
        st.error(f"Error guardando archivos en carpeta: {str(e)}")
        return []

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
    
    return position_settings

def main():
    st.set_page_config(
        page_title="Procesador Masivo de PDFs",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("🎯 Procesador Masivo de PDFs - Posicionamiento Preciso")
    st.markdown("""
    Esta aplicación procesa múltiples archivos PDF con control preciso sobre la posición de los datos insertados.
    """)
    
    # Configuraciones del sidebar
    output_folder = select_output_folder()
    position_settings = get_position_settings()
    
    # Cargar archivo Excel
    st.header("1. 📊 Cargar Archivo Excel con Datos")
    excel_file = st.file_uploader(
        "Sube el archivo Listado.xlsx", 
        type=["xlsx"],
        help="Debe contener las columnas: ACONEX, SISTEMA, SUB SISTEMA"
    )
    
    df = None
    if excel_file is not None:
        try:
            df = pd.read_excel(excel_file)
            st.success(f"✅ Excel cargado - {len(df)} registros encontrados")
            
            if validate_dataframe(df):
                st.dataframe(df.head(10), use_container_width=True)
                
                # Mostrar estadísticas
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
        help="Selecciona todos los PDFs que quieres procesar"
    )
    
    if pdf_files and df is not None and validate_dataframe(df):
        st.success(f"✅ {len(pdf_files)} archivos PDF cargados")
        
        # Análisis avanzado de coordenadas
        if st.checkbox("🔍 Análisis Avanzado de Coordenadas", help="Analiza las posiciones exactas en el PDF"):
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
                            st.write(f"  - **{field_name}:**")
                            st.write(f"    - Posición: ({field_data['x']:.1f}, {field_data['y']:.1f})")
                            st.write(f"    - Tamaño: {field_data['width']:.1f} × {field_data['height']:.1f}")
                            st.write(f"    - Sugerido: ({field_data['suggested_x']:.1f}, {field_data['suggested_y']:.1f})")
                        else:
                            st.write(f"  - **{field_name}:** ❌ No encontrado")
                    
                    st.write("---")
            
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
            if matched_files:
                for match in matched_files[:3]:
                    st.write(f"📄 **{match['name']}**")
        
        with col2:
            if unmatched_files:
                st.warning(f"⚠️ {len(unmatched_files)} no emparejados")
        
        # Procesar archivos
        if matched_files:
            st.header("4. ⚙️ Procesar Archivos")
            
            # Mostrar configuración actual
            st.info(f"🎯 **Configuración de posicionamiento:**")
            if 'custom_positions' in position_settings:
                st.write("**Coordenadas manuales:**")
                for field, pos in position_settings['custom_positions'].items():
                    st.write(f"  - {field.upper()}: ({pos['x']}, {pos['y']})")
            else:
                st.write(f"**Offsets automáticos:** X: {position_settings['x_offset']}, Y: {position_settings['y_offset']}")
            st.write(f"**Tamaño de fuente:** {position_settings['font_size']}")
            
            if output_folder:
                st.info(f"📁 **Destino:** `{output_folder}`")
            
            if st.button("🚀 Iniciar Procesamiento con Ajustes Actuales", type="primary"):
                st.header("5. 📊 Procesando Archivos...")
                
                with tempfile.TemporaryDirectory() as temp_dir:
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
                                filename = match_info['name']
                                filepath = os.path.join(temp_dir, filename)
                                
                                with open(filepath, "wb") as f:
                                    f.write(modified_pdf_bytes)
                                
                                processed_files.append({
                                    'filename': filename,
                                    'aconex': match_info['aconex'],
                                    'sistema': match_info['sistema'],
                                    'subsistema': match_info['subsistema'],
                                    'filepath': filepath
                                })
                                
                                st.sidebar.success(f"✅ {match_info['name']}")
                            else:
                                errors.append(f"{match_info['name']}")
                                
                        except Exception as e:
                            errors.append(f"{match_info['name']}: {str(e)}")
                    
                    # Resultados
                    st.header("6. ✅ Resultados Finales")
                    
                    if processed_files:
                        st.success(f"✅ {len(processed_files)} archivos procesados")
                        
                        # Guardar en carpeta
                        if output_folder and processed_files:
                            saved_files = save_to_folder(processed_files, output_folder)
                            if saved_files:
                                st.success(f"📁 {len(saved_files)} archivos guardados en: `{output_folder}`")
                        
                        # Descargas
                        st.header("7. 📥 Descargar Archivos")
                        
                        # ZIP
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                            for file_info in processed_files:
                                zip_file.write(file_info['filepath'], file_info['filename'])
                        
                        zip_buffer.seek(0)
                        
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        st.download_button(
                            label=f"📦 Descargar ZIP ({len(processed_files)} archivos)",
                            data=zip_buffer,
                            file_name=f"pdfs_procesados_{timestamp}.zip",
                            mime="application/zip",
                            type="primary"
                        )
                        
                        # Individuales
                        st.subheader("Descargas Individuales")
                        for i, file_info in enumerate(processed_files[:5]):
                            with open(file_info['filepath'], "rb") as f:
                                st.download_button(
                                    label=f"📄 {file_info['filename']}",
                                    data=f,
                                    file_name=file_info['filename'],
                                    mime="application/pdf",
                                    key=f"dl_{i}"
                                )
                    
                    if errors:
                        st.error(f"❌ {len(errors)} errores:")
                        for error in errors:
                            st.write(f"- {error}")

if __name__ == "__main__":
    main()