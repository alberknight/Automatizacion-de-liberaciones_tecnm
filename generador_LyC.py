import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
import os
import locale

# Configurar locale para meses en español
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

def configurar_fuente_noto_sans(doc):
    # Establecer Noto Sans como fuente principal
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Noto Sans'
    font.size = Pt(10)
    
    # Configurar para caracteres complejos (opcional)
    rFonts = style.element.rPr.rFonts
    rFonts.set(qn('w:eastAsia'), 'Noto Sans')

def generar_constancia(row, carrera):
    # Cargar plantilla Word
    doc = Document('Liberacion.docx')
    
    # Configurar fuente Noto Sans para todo el documento
    configurar_fuente_noto_sans(doc)
    
    # Mapeo de campos del Excel a la plantilla Word
    campos = {
        'NOMBRE_ALUMNO': row['NOMBRE'],  # Notar typo en plantilla original
        'NOMBRE_ALUMNO': row['NOMBRE'],
        'NO_DE_CONTROL': str(row['NO. CONTROL']),
        'CARRERA_ALUMNO': carrera,
        'DEPENDENCIA_ALUMNO': row['DEPENDENCIA'],
        'PROGRAMA_ALUMNO': row['PROGRAMA']
    }
    
    # Formatear fechas en español
    fecha_inicio = pd.to_datetime(row['FECHA DE INICIO']).strftime('%d de %B del %Y').capitalize()
    fecha_termino = pd.to_datetime(row['FECHA DE TERMINO']).strftime('%d de %B del %Y').capitalize()
    
    # Reemplazar campos en el documento manteniendo formato
    for paragraph in doc.paragraphs:
        # Reemplazar campos generales
        for campo, valor in campos.items():
            if campo in paragraph.text:
                for run in paragraph.runs:
                    if campo in run.text:
                        run.text = run.text.replace(campo, valor)
                        # Mantener formato original del run
                        run.font.name = 'Noto Sans'
        
        # Reemplazar fechas específicas
        if '06 de Noviembre del 2023' in paragraph.text:
            for run in paragraph.runs:
                if '06 de Noviembre del 2023' in run.text:
                    run.text = run.text.replace('06 de Noviembre del 2023', fecha_inicio)
                    run.font.name = 'Noto Sans'
        
        if '06 de Mayo del 2024' in paragraph.text:
            for run in paragraph.runs:
                if '06 de Mayo del 2024' in run.text:
                    run.text = run.text.replace('06 de Mayo del 2024', fecha_termino)
                    run.font.name = 'Noto Sans'
    
    # Asegurar que el pie de página mantenga la fuente
    for section in doc.sections:
        for paragraph in section.footer.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Noto Sans'
    
    # Crear directorio por carrera si no existe
    if not os.path.exists(carrera):
        os.makedirs(carrera)
    
    # Guardar documento con nombre adecuado
    nombre_archivo = f"{carrera}/Constancia_{row['NOMBRE'].replace(' ', '_')}_{row['NO. CONTROL']}.docx"
    doc.save(nombre_archivo)
    print(f"Constancia generada para {row['NOMBRE']} en {carrera}")

def procesar_excel(archivo_excel):
    # Leer todas las hojas del Excel
    xls = pd.ExcelFile(archivo_excel)
    
    for sheet_name in xls.sheet_names:
        # Leer datos de la hoja actual
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Normalizar nombres de columnas
        df.columns = df.columns.str.upper().str.strip()
        
        # Corregir nombres específicos
        column_mapping = {
            'NOMBRE': 'NOMBRE',
            'NO. CONTROL': 'NO. CONTROL',
            'NO CONTROL': 'NO. CONTROL',
            'DEPENDENCIA': 'DEPENDENCIA',
            'CARRERA': 'CARRERA',
            'PROGRAMA': 'PROGRAMA',
            'FECHA DE INICIO': 'FECHA DE INICIO',
            'FECHA INICIO': 'FECHA DE INICIO',
            'FECHA DE TERMINO': 'FECHA DE TERMINO',
            'FECHA TERMINO': 'FECHA DE TERMINO'
        }
        
        df = df.rename(columns={col: column_mapping.get(col, col) for col in df.columns})
        
        # Eliminar filas vacías
        df = df.dropna(subset=['NOMBRE'])
        
        # Obtener nombre de la carrera
        carrera = sheet_name.replace('Í', 'I').replace('Ó', 'O').strip().upper()
        
        # Procesar cada estudiante
        for _, row in df.iterrows():
            # Manejar valores NaN
            row = row.fillna('')
            generar_constancia(row, carrera)

if __name__ == "__main__":
    print("Iniciando generación de constancias...")
    archivo_excel = "ENERO - JULIO 2025.xlsx"
    
    try:
        procesar_excel(archivo_excel)
        print("\nProceso completado exitosamente.")
        print("Constancias generadas y organizadas por carrera.")
    except Exception as e:
        print(f"\nError durante el proceso: {str(e)}")