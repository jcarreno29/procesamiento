# app_streamlit.py
import streamlit as st
import pandas as pd
import os
import glob
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
import tempfile
import zipfile
from pathlib import Path

class PDFEditorStreamlit:
    def __init__(self):
        self.datos = {}
        self.orientacion = "vertical"
        self.fuente = "Helvetica"
        self.configurar_fuente()
    
    def configurar_fuente(self):
        """Configura la fuente Arial si está disponible"""
        try:
            font_paths = [
                'arial.ttf',
                'Arial.ttf',
                '/usr/share/fonts/truetype/msttcorefonts/Arial.ttf',
                'C:/Windows/Fonts/arial.ttf',
                '/Library/Fonts/Arial.ttf'
            ]
            
            for font_path in font_paths:
                try:
                    pdfmetrics.registerFont(TTFont('Arial', font_path))
                    self.fuente = "Arial"
                    return True
                except:
                    continue
            
            self.fuente = "Helvetica"
            return True
            
        except Exception as e:
            st.warning(f"No se pudo cargar Arial: {e}")
            self.fuente = "Helvetica"
            return True

    def leer_datos_excel(self, uploaded_file):
        """Lee los datos del archivo Excel subido"""
        try:
            # Leer el archivo Excel
            df = pd.read_excel(uploaded_file)
            
            # Mostrar información del archivo
            st.success(f"✅ Archivo Excel cargado correctamente")
            st.info(f"**Información del archivo:**")
            st.info(f"- Total de filas: {len(df)}")
            st.info(f"- Total de columnas: {len(df.columns)}")
            st.info(f"- Columnas encontradas: {list(df.columns)}")
            
            # CORRECCIÓN: Buscar columnas mejorado (incluye acentos y mayúsculas)
            columnas_mapeadas = self._mapear_columnas(df)
            
            # Mostrar mapeo de columnas
            st.info("**Mapeo de columnas detectado:**")
            for campo, idx in columnas_mapeadas.items():
                if idx is not None:
                    st.info(f"- {campo.upper()}: Columna '{df.columns[idx]}' (índice {idx})")
            
            # Procesar datos
            datos = {}
            registros_procesados = 0
            
            for index, row in df.iterrows():
                codigo = self._obtener_valor(row, columnas_mapeadas['codigo'])
                sistema = self._obtener_valor(row, columnas_mapeadas['sistema'])
                subsistema = self._obtener_valor(row, columnas_mapeadas['subsistema'])
                
                if codigo and codigo != 'nan':
                    datos[str(codigo).strip()] = {
                        'sistema': sistema,
                        'subsistema': subsistema
                    }
                    registros_procesados += 1
            
            self.datos = datos
            
            # Mostrar resumen
            st.success(f"**Resumen de datos procesados:**")
            st.success(f"- Total de registros válidos: {len(self.datos)}")
            
            if self.datos:
                # Crear DataFrame para mostrar
                mostrar_df = pd.DataFrame([
                    {
                        'Código': codigo,
                        'Sistema': info['sistema'],
                        'Subsistema': info['subsistema']
                    }
                    for codigo, info in list(self.datos.items())[:20]  # Mostrar primeros 20
                ])
                
                st.dataframe(mostrar_df, use_container_width=True)
                
                if len(self.datos) > 20:
                    st.info(f"... y {len(self.datos) - 20} registros más")
            
            return True
            
        except Exception as e:
            st.error(f"❌ Error al leer el archivo Excel: {str(e)}")
            return False

    def _mapear_columnas(self, df):
        """Mapea automáticamente las columnas del Excel - VERSIÓN MEJORADA"""
        columnas_mapeadas = {'codigo': None, 'sistema': None, 'subsistema': None}
        
        # Listas más completas incluyendo acentos
        nombres_codigo = ['código', 'codigo', 'code', 'id', 'número', 'numero', 'n°', 'no']
        nombres_sistema = ['sistema', 'system', 'sist']
        nombres_subsistema = ['subsistema', 'sub-sistema', 'subsystem', 'subsist']
        
        for idx, col_name in enumerate(df.columns):
            col_name_str = str(col_name)
            col_name_clean = col_name_str.lower().strip()
            
            # Eliminar espacios extra y caracteres especiales
            col_name_clean = col_name_clean.replace('_', ' ').replace('-', ' ').replace('.', ' ')
            
            # Buscar coincidencias
            if any(nombre in col_name_clean for nombre in nombres_codigo):
                if columnas_mapeadas['codigo'] is None:  # Solo asignar si no está asignado
                    columnas_mapeadas['codigo'] = idx
            elif any(nombre in col_name_clean for nombre in nombres_sistema):
                if columnas_mapeadas['sistema'] is None:
                    columnas_mapeadas['sistema'] = idx
            elif any(nombre in col_name_clean for nombre in nombres_subsistema):
                if columnas_mapeadas['subsistema'] is None:
                    columnas_mapeadas['subsistema'] = idx
        
        # Si no se encontraron algunas columnas, usar las primeras disponibles
        if columnas_mapeadas['codigo'] is None and len(df.columns) >= 1:
            columnas_mapeadas['codigo'] = 0
        if columnas_mapeadas['sistema'] is None and len(df.columns) >= 2:
            columnas_mapeadas['sistema'] = 1
        if columnas_mapeadas['subsistema'] is None and len(df.columns) >= 3:
            columnas_mapeadas['subsistema'] = 2
            
        return columnas_mapeadas

    def _obtener_valor(self, row, idx):
        """Obtiene y limpia un valor de la fila"""
        if idx is not None and idx < len(row):
            valor = row.iloc[idx]
            if pd.notna(valor):
                return str(valor).strip()
        return ""

    def generar_pdf_coordenadas(self, pdf_file):
        """Genera un PDF con cuadrícula de coordenadas para referencia"""
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                temp_path = tmp.name
            
            doc = fitz.open(pdf_file)
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont("Helvetica", 6)
            
            # Dibujar cuadrícula
            for x in range(50, 600, 50):
                for y in range(50, 750, 50):
                    can.drawString(x, y, f"({x},{y})")
            
            # Marcadores de referencia
            can.setFillColorRGB(1, 0, 0)
            puntos = [(100, 500), (200, 500), (300, 500), (400, 500)]
            for x, y in puntos:
                can.circle(x, y, 3, fill=1)
                can.setFillColorRGB(0, 0, 0)
                can.drawString(x + 5, y, f"REF({x},{y})")
                can.setFillColorRGB(1, 0, 0)
            
            can.save()
            packet.seek(0)
            
            # Fusionar con PDF original
            nuevo_pdf = PdfReader(packet)
            pdf_existente = PdfReader(open(pdf_file, "rb"))
            output = PdfWriter()
            
            for pagina in pdf_existente.pages:
                pagina.merge_page(nuevo_pdf.pages[0])
                output.add_page(pagina)
            
            with open(temp_path, "wb") as f:
                output.write(f)
            
            doc.close()
            
            return temp_path
            
        except Exception as e:
            st.error(f"Error generando PDF con coordenadas: {e}")
            return None

    def editar_pdf(self, archivo_pdf_entrada, archivo_pdf_salida, posiciones):
        """Edita SOLO la primera página del PDF con los datos correspondientes"""
        try:
            codigo_pdf = os.path.splitext(os.path.basename(archivo_pdf_entrada))[0]
            
            if codigo_pdf not in self.datos:
                st.warning(f"No hay datos para: {codigo_pdf}")
                return False
            
            sistema = self.datos[codigo_pdf]['sistema']
            subsistema = self.datos[codigo_pdf]['subsistema']
            
            # Usar PyMuPDF para edición
            doc = fitz.open(archivo_pdf_entrada)
            
            # SOLO EDITAR LA PRIMERA PÁGINA
            if len(doc) > 0:
                primera_pagina = doc[0]
                
                if self.orientacion == "vertical":
                    # Texto vertical (90° derecha)
                    if sistema:
                        primera_pagina.insert_text(
                            (posiciones[0]['x'], posiciones[0]['y']),
                            sistema, fontsize=8, fontname="helv",
                            color=(0, 0, 0), rotate=90
                        )
                    
                    if subsistema:
                        primera_pagina.insert_text(
                            (posiciones[1]['x'], posiciones[1]['y']),
                            subsistema, fontsize=8, fontname="helv",
                            color=(0, 0, 0), rotate=90
                        )
                    
                    primera_pagina.insert_text(
                        (posiciones[2]['x'], posiciones[2]['y']),
                        codigo_pdf, fontsize=8, fontname="helv",
                        color=(0, 0, 0), rotate=90
                    )
                    
                else:
                    # Texto horizontal
                    if sistema:
                        primera_pagina.insert_text(
                            (posiciones[0]['x'], posiciones[0]['y']),
                            sistema, fontsize=8, fontname="helv",
                            color=(0, 0, 0), rotate=0
                        )
                    
                    if subsistema:
                        primera_pagina.insert_text(
                            (posiciones[1]['x'], posiciones[1]['y']),
                            subsistema, fontsize=8, fontname="helv",
                            color=(0, 0, 0), rotate=0
                        )
                    
                    primera_pagina.insert_text(
                        (posiciones[2]['x'], posiciones[2]['y']),
                        codigo_pdf, fontsize=8, fontname="helv",
                        color=(0, 0, 0), rotate=0
                    )
            else:
                st.error(f"El PDF {archivo_pdf_entrada} no tiene páginas")
                doc.close()
                return False
            
            doc.save(archivo_pdf_salida)
            doc.close()
            
            return True
            
        except Exception as e:
            st.error(f"Error editando PDF: {e}")
            return False

    def procesar_lote(self, pdf_files, posiciones, progress_bar, status_text):
        """Procesa un lote de archivos PDF"""
        resultados = []
        
        for i, pdf_file in enumerate(pdf_files):
            try:
                # Actualizar progreso
                progress_bar.progress((i + 1) / len(pdf_files))
                status_text.text(f"Procesando {i+1}/{len(pdf_files)}: {pdf_file.name}")
                
                # Crear archivo temporal de salida
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    temp_output = tmp.name
                
                # Editar PDF
                if self.editar_pdf(pdf_file, temp_output, posiciones):
                    resultados.append({
                        'nombre': pdf_file.name,
                        'ruta': temp_output,
                        'estado': '✅ Completado'
                    })
                else:
                    resultados.append({
                        'nombre': pdf_file.name,
                        'ruta': None,
                        'estado': '❌ Error'
                    })
                    
            except Exception as e:
                resultados.append({
                    'nombre': pdf_file.name,
                    'ruta': None,
                    'estado': f'❌ Error: {str(e)}'
                })
        
        return resultados

def main():
    """Función principal de la aplicación Streamlit"""
    st.set_page_config(
        page_title="Editor de PDFs con Excel",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("📄 Editor de PDFs con Datos de Excel")
    st.markdown("**Carga un archivo Excel con los datos y archivos PDF para editar**")
    st.markdown("💡 *Los datos se insertarán SOLO en la PRIMERA PÁGINA de cada PDF*")
    
    # Inicializar sesión state
    if 'editor' not in st.session_state:
        st.session_state.editor = PDFEditorStreamlit()
    if 'posiciones' not in st.session_state:
        st.session_state.posiciones = None
    if 'datos_cargados' not in st.session_state:
        st.session_state.datos_cargados = False
    
    editor = st.session_state.editor
    
    # Sidebar para configuración
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        # Configurar orientación
        st.subheader("Orientación del texto")
        orientacion = st.radio(
            "Selecciona la orientación:",
            ["Vertical (90° derecha)", "Horizontal (normal)"],
            index=0
        )
        editor.orientacion = "vertical" if orientacion == "Vertical (90° derecha)" else "horizontal"
        
        st.info(f"**Orientación actual:** {editor.orientacion.upper()}")
        
        # Posiciones por defecto según orientación
        if editor.orientacion == "vertical":
            posiciones_default = [
                {'x': 50, 'y': 500},   # Sistema
                {'x': 50, 'y': 470},   # Subsistema
                {'x': 50, 'y': 440}    # Código
            ]
        else:
            posiciones_default = [
                {'x': 100, 'y': 35},   # Sistema
                {'x': 250, 'y': 35},   # Subsistema
                {'x': 430, 'y': 35}    # Código
            ]
    
    # Contenedor principal
    tab1, tab2, tab3 = st.tabs(["📊 Cargar Datos", "📍 Configurar Posiciones", "🚀 Procesar PDFs"])
    
    with tab1:
        st.header("1. Cargar Datos desde Excel")
        
        # Subir archivo Excel
        uploaded_excel = st.file_uploader(
            "Carga tu archivo Excel (.xlsx o .xls)",
            type=['xlsx', 'xls'],
            key="excel_uploader"
        )
        
        if uploaded_excel is not None:
            if st.button("📥 Procesar Datos del Excel", type="primary"):
                with st.spinner("Procesando archivo Excel..."):
                    if editor.leer_datos_excel(uploaded_excel):
                        st.session_state.datos_cargados = True
                        st.success("✅ Datos cargados correctamente")
                        
                        # Mostrar resumen
                        st.subheader("Resumen de Datos")
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total Registros", len(editor.datos))
                        with col2:
                            st.metric("Fuente", editor.fuente)
                        with col3:
                            st.metric("Orientación", editor.orientacion)
    
    with tab2:
        st.header("2. Configurar Posiciones del Texto")
        
        if not st.session_state.datos_cargados:
            st.warning("⚠ Primero carga los datos desde la pestaña 'Cargar Datos'")
        else:
            # Subir PDF de ejemplo
            st.subheader("Subir PDF de ejemplo")
            pdf_ejemplo = st.file_uploader(
                "Carga un PDF de ejemplo para configurar posiciones",
                type=['pdf'],
                key="pdf_ejemplo_uploader"
            )
            
            if pdf_ejemplo:
                # Guardar PDF temporalmente
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(pdf_ejemplo.getvalue())
                    pdf_path = tmp.name
                
                # Opción para generar PDF con coordenadas
                if st.button("🛠️ Generar PDF con coordenadas", type="secondary"):
                    with st.spinner("Generando PDF con cuadrícula de coordenadas..."):
                        pdf_coords = editor.generar_pdf_coordenadas(pdf_path)
                        
                        if pdf_coords:
                            with open(pdf_coords, "rb") as f:
                                st.download_button(
                                    label="📥 Descargar PDF con coordenadas",
                                    data=f,
                                    file_name="pdf_con_coordenadas.pdf",
                                    mime="application/pdf"
                                )
                            st.info("📋 Abre el PDF descargado para ver las coordenadas y configurar las posiciones")
                
                # Configurar posiciones manualmente
                st.subheader("Configurar Coordenadas")
                st.info(f"Configura las coordenadas para texto **{editor.orientacion.upper()}**")
                
                textos = ["SISTEMA", "SUBSISTEMA", "CÓDIGO"]
                posiciones = []
                
                cols = st.columns(3)
                for i, texto in enumerate(textos):
                    with cols[i]:
                        st.markdown(f"**{texto}**")
                        x = st.number_input(f"Coordenada X", value=posiciones_default[i]['x'], key=f"x_{i}")
                        y = st.number_input(f"Coordenada Y", value=posiciones_default[i]['y'], key=f"y_{i}")
                        posiciones.append({'x': x, 'y': y})
                
                # Guardar posiciones
                if st.button("💾 Guardar Posiciones", type="primary"):
                    st.session_state.posiciones = posiciones
                    st.success("✅ Posiciones guardadas correctamente")
                    
                    # Mostrar resumen
                    st.subheader("Resumen de Posiciones")
                    for i, pos in enumerate(posiciones):
                        st.info(f"**{textos[i]}**: ({pos['x']}, {pos['y']})")
    
    with tab3:
        st.header("3. Procesar Archivos PDF")
        
        if not st.session_state.datos_cargados:
            st.warning("⚠ Primero carga los datos desde la pestaña 'Cargar Datos'")
        elif st.session_state.posiciones is None:
            st.warning("⚠ Configura las posiciones desde la pestaña 'Configurar Posiciones'")
        else:
            # Subir archivos PDF
            st.subheader("Cargar Archivos PDF")
            uploaded_pdfs = st.file_uploader(
                "Selecciona los archivos PDF a procesar",
                type=['pdf'],
                accept_multiple_files=True,
                key="pdfs_uploader"
            )
            
            if uploaded_pdfs:
                st.success(f"✅ {len(uploaded_pdfs)} archivos PDF listos para procesar")
                
                # Mostrar configuración
                st.subheader("Configuración Actual")
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"**Datos cargados:** {len(editor.datos)} registros")
                    st.info(f"**Posiciones configuradas:** Sí")
                with col2:
                    st.info(f"**Orientación:** {editor.orientacion}")
                    st.info(f"**PDFs a procesar:** {len(uploaded_pdfs)}")
                
                # Iniciar procesamiento
                if st.button("🚀 Iniciar Procesamiento", type="primary"):
                    # Barra de progreso
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Procesar archivos
                    resultados = editor.procesar_lote(
                        uploaded_pdfs,
                        st.session_state.posiciones,
                        progress_bar,
                        status_text
                    )
                    
                    # Mostrar resultados
                    st.subheader("📊 Resultados del Procesamiento")
                    
                    # Crear DataFrame de resultados
                    resultados_df = pd.DataFrame(resultados)
                    st.dataframe(resultados_df[['nombre', 'estado']], use_container_width=True)
                    
                    # Estadísticas
                    completados = sum(1 for r in resultados if '✅' in r['estado'])
                    st.success(f"✅ {completados}/{len(resultados)} archivos procesados correctamente")
                    
                    # Crear archivo ZIP con resultados
                    if completados > 0:
                        st.subheader("📦 Descargar Resultados")
                        
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp_zip:
                            zip_path = tmp_zip.name
                        
                        with zipfile.ZipFile(zip_path, 'w') as zipf:
                            for resultado in resultados:
                                if resultado['ruta'] and os.path.exists(resultado['ruta']):
                                    arcname = f"editados/{resultado['nombre']}"
                                    zipf.write(resultado['ruta'], arcname)
                        
                        # Botón de descarga
                        with open(zip_path, "rb") as f:
                            st.download_button(
                                label="📥 Descargar todos los PDFs editados (ZIP)",
                                data=f,
                                file_name="pdfs_editados.zip",
                                mime="application/zip"
                            )
                        
                        # Botones para descargar individuales
                        st.subheader("Descargar individualmente")
                        cols = st.columns(3)
                        for idx, resultado in enumerate(resultados):
                            if resultado['ruta'] and os.path.exists(resultado['ruta']):
                                with cols[idx % 3]:
                                    with open(resultado['ruta'], "rb") as f:
                                        st.download_button(
                                            label=f"📥 {resultado['nombre'][:20]}...",
                                            data=f,
                                            file_name=f"editado_{resultado['nombre']}",
                                            mime="application/pdf"
                                        )

if __name__ == "__main__":
    main()
