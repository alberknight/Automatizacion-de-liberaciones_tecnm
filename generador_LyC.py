import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import os
import locale

# Configuración inicial
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
FOLIO_INICIAL = 2020
folio_actual = FOLIO_INICIAL

def configurar_fuente_noto_sans(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Noto Sans'
    font.size = Pt(10)
    rFonts = style.element.rPr.rFonts
    rFonts.set(qn('w:eastAsia'), 'Noto Sans')

def verificar_campos_requeridos(row):
    # Columnas H a O (índices 7 a 14 en DataFrame de pandas)
    columnas_requeridas = [
        'CARTA DE PRESENTACION',
        'CARTA DE ACEPTACION',
        'CARTA SOLICITUD',
        'CARTA COMPROMISO',
        'REPORTE 1',
        'CALIFICACION 1',
        'REPORTE 2',
        'CALIFICACION 2'
    ]
    
    for col in columnas_requeridas:
        if col in row and (pd.isna(row[col]) or row[col] == ''):
            return False
    return True

def generar_constancia(row, carrera):
    global folio_actual
    
    doc = Document('Liberacion.docx')
    configurar_fuente_noto_sans(doc)
    
    campos = {
        'NOMBRE_ALUMNO': row['NOMBRE'],
        'NO_DE_CONTROL': str(row['NO. CONTROL']),
        'CARRERA_ALUMNO': carrera,
        'DEPENDENCIA_ALUMNO': row['DEPENDENCIA'],
        'PROGRAMA_ALUMNO': row['PROGRAMA'],
        'NO_DE_FOLIO': str(folio_actual)
    }
    
    fecha_inicio = pd.to_datetime(row['FECHA DE INICIO']).strftime('%d de %B del %Y').capitalize()
    fecha_termino = pd.to_datetime(row['FECHA DE TERMINO']).strftime('%d de %B del %Y').capitalize()
    
    for paragraph in doc.paragraphs:
        for campo, valor in campos.items():
            if campo in paragraph.text:
                for run in paragraph.runs:
                    if campo in run.text:
                        run.text = run.text.replace(campo, valor)
                        run.font.name = 'Noto Sans'
        
        if '06 de Noviembre del 2023' in paragraph.text:
            for run in paragraph.runs:
                if '06 de Noviembre del 2023' in run.text:
                    run.text = run.text.replace('06 de Noviembre del 2023', fecha_inicio)
        
        if '06 de Mayo del 2024' in paragraph.text:
            for run in paragraph.runs:
                if '06 de Mayo del 2024' in run.text:
                    run.text = run.text.replace('06 de Mayo del 2024', fecha_termino)
    
    folio_actual += 1
    
    if not os.path.exists(carrera):
        os.makedirs(carrera)
    
    nombre_archivo = f"{carrera}/Constancia_{row['NOMBRE'].replace(' ', '_')}_{row['NO. CONTROL']}_Folio_{folio_actual-1}.docx"
    doc.save(nombre_archivo)
    print(f"Constancia generada para {row['NOMBRE']} - Folio: {folio_actual-1}")

def procesar_excel(archivo_excel):
    global folio_actual
    
    xls = pd.ExcelFile(archivo_excel)
    print(f"\nIniciando generación con folio inicial: {FOLIO_INICIAL}")
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df.columns = df.columns.str.upper().str.strip()
        
        column_mapping = {
            'NOMBRE': 'NOMBRE',
            'NO. CONTROL': 'NO. CONTROL',
            'DEPENDENCIA': 'DEPENDENCIA',
            'CARRERA': 'CARRERA',
            'PROGRAMA': 'PROGRAMA',
            'FECHA DE INICIO': 'FECHA DE INICIO',
            'FECHA DE TERMINO': 'FECHA DE TERMINO',
            'CARTA DE PRESENTACION': 'CARTA DE PRESENTACION',
            'CARTA DE ACEPTACION': 'CARTA DE ACEPTACION',
            'CARTA SOLICITUD': 'CARTA SOLICITUD',
            'CARTA COMPROMISO': 'CARTA COMPROMISO',
            'REPORTE 1': 'REPORTE 1',
            'CALIFICACION 1': 'CALIFICACION 1',
            'REPORTE 2': 'REPORTE 2',
            'CALIFICACION 2': 'CALIFICACION 2',
            'REPORTE 3': 'REPORTE 3',
            'CALIFICACION 3': 'CALIFICACION 3'
        }
        df = df.rename(columns={col: column_mapping.get(col, col) for col in df.columns})
        
        df = df.dropna(subset=['NOMBRE'])
        carrera = sheet_name.replace('Í', 'I').replace('Ó', 'O').strip().upper()
        
        print(f"\nProcesando carrera: {carrera}")
        print(f"Estudiantes encontrados: {len(df)}")
        
        estudiantes_validos = 0
        estudiantes_omitidos = 0
        
        for _, row in df.iterrows():
            row = row.fillna('')
            
            if verificar_campos_requeridos(row):
                generar_constancia(row, carrera)
                estudiantes_validos += 1
            else:
                print(f"Omitiendo a {row['NOMBRE']} - Faltan documentos requeridos")
                estudiantes_omitidos += 1
        
        print(f"Resumen para {carrera}:")
        print(f"- Constancias generadas: {estudiantes_validos}")
        print(f"- Estudiantes omitidos: {estudiantes_omitidos}")

if __name__ == "__main__":
    print("=== Sistema de Generación de Constancias ===")
    print(f"Folio inicial configurado en: {FOLIO_INICIAL}")
    
    archivo_excel = "ENERO - JULIO 2025.xlsx"
    
    try:
        procesar_excel(archivo_excel)
        print("\nResumen final:")
        print(f"- Folio inicial usado: {FOLIO_INICIAL}")
        print(f"- Último folio asignado: {folio_actual-1}")
        print(f"- Total de constancias generadas: {folio_actual - FOLIO_INICIAL}")
        print("\nProceso completado exitosamente.")
    except Exception as e:
        print(f"\nError durante el proceso: {str(e)}")