# -*- coding: utf-8 -*-
"""
Script de procesamiento académico y generación de boletines PDF para la UEEH.
Cruza las notas de archivos Excel trimestrales y supletorios.
"""

from __future__ import annotations
import os
import sys
import json

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass
import argparse
import pandas as pd
import numpy as np
from pathlib import Path

# ReportLab imports
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# 1. Funciones de cálculo académico
def truncar_2_decimales(valor: float) -> float:
    try:
        # Usar truncamiento simple a 2 decimales para replicar el comportamiento de TRUNC de Excel
        return int(valor * 100) / 100.0
    except (ValueError, TypeError):
        return 0.0

def calcular_promedio_anual(t1: float, t2: float, t3: float) -> float:
    return truncar_2_decimales((t1 + t2 + t3) / 3.0)

def calcular_resultado_con_supletorio(nota_final: float, nota_supletorio: float | None) -> float:
    if nota_final >= 7.0:
        return truncar_2_decimales(nota_final)
    if nota_supletorio is not None and nota_supletorio >= 7.0:
        return 7.0
    return truncar_2_decimales(nota_final)

def determinar_estado_materia(nota_final: float, nota_supletorio: float | None) -> str:
    if nota_final >= 7.0:
        return "APROBADO"
    if nota_supletorio is not None:
        # Ya tomó supletorio y no llegó a 7
        return "REPROBADO"
    return "SUPLETORIO"

# 2. Análisis e Importación de Excels
def normalizar_columna(col_name: str) -> str:
    import unicodedata
    s = str(col_name).strip().lower()
    s = "".join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return s

def detectar_columnas(df: pd.DataFrame) -> tuple[str | None, str | None, list[str]]:
    id_col = None
    name_col = None
    subject_cols = []
    
    ignore_keywords = [
        "nro", "n.", "no.", "lista", "numero", "promedio", "total", 
        "estado", "observacion", "comentario", "firma", "obs"
    ]
    
    for col in df.columns:
        norm = normalizar_columna(col)
        # Identificar columna ID
        if any(k in norm for k in ["cedula", "identificacion", "id", "codigo"]):
            if id_col is None:
                id_col = col
            continue
        # Identificar columna Nombre
        if any(k in norm for k in ["estudiante", "nombre", "apellidos", "alumno", "listado"]):
            if name_col is None:
                name_col = col
            continue
            
    # Asignaturas: Columnas numéricas que no se ignoran
    for col in df.columns:
        if col == id_col or col == name_col:
            continue
        norm = normalizar_columna(col)
        if any(k in norm for k in ignore_keywords):
            continue
        
        # Verificar si la columna tiene valores numéricos
        numeric_series = pd.to_numeric(df[col], errors='coerce')
        if numeric_series.notna().any():
            subject_cols.append(col)
            
    return id_col, name_col, subject_cols

def obtener_hoja_principal(wb):
    """Obtiene la hoja oficial principal del formato de notas."""
    if "Reporte Periodo" in wb.sheetnames:
        return wb["Reporte Periodo"]
    return None


def _celda_con_valor_combinado(sheet, cell_ref):
    """Devuelve la celda directa o la celda superior izquierda del rango combinado."""
    cell = sheet[cell_ref]
    if cell.value not in (None, ""):
        return cell

    for merged_range in sheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell


def _valor_a_texto(value, cell=None):
    """Convierte valores de Excel a texto limpio, respetando formatos con ceros."""
    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        value = int(value)

    if isinstance(value, int):
        number_format = getattr(cell, "number_format", "") if cell is not None else ""
        if number_format and set(number_format) == {"0"} and len(number_format) > len(str(value)):
            return str(value).zfill(len(number_format)).strip()
        return str(value).strip()

    return str(value).strip()


def obtener_valor_celda(sheet, cell_ref, fallback_refs=None):
    """
    Lee una celda, revisa rangos combinados si está vacía y luego prueba fallbacks.
    Devuelve siempre texto limpio.
    """
    refs = [cell_ref, *(fallback_refs or [])]
    for ref in refs:
        cell = _celda_con_valor_combinado(sheet, ref)
        value = _valor_a_texto(cell.value, cell)
        if value:
            return value
    return ""

def cargar_excel_datos(file_path: str) -> dict[str, dict]:
    """
    Lee el excel buscando la hoja oficial 'Reporte Periodo'.
    Busca a partir de la fila donde dice 'LISTADO' en la columna A
    y debajo de 'CEDULA' en la columna B para extraer el nombre y la cédula/ID.
    Retorna un diccionario indexado por cédula/nombre de los estudiantes,
    con 'notas' vacío por ahora.
    """
    if not file_path or not os.path.exists(file_path):
        print(f"[cargar_excel_datos] Ruta inválida o inexistente: {file_path}", file=sys.stderr)
        return {}
        
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = obtener_hoja_principal(wb)
        if sheet is None:
            print(f"[cargar_excel_datos] No existe la hoja 'Reporte Periodo'. Hojas encontradas: {', '.join(wb.sheetnames)}", file=sys.stderr)
            return {}
        print(f"[cargar_excel_datos] Hoja usada: {sheet.title}", file=sys.stderr)
        
        # Buscar la fila de encabezado
        start_row = None
        for r in range(1, 100):
            val_a = sheet.cell(row=r, column=1).value
            val_b = sheet.cell(row=r, column=2).value
            if val_a and val_b:
                str_a = str(val_a).strip().upper()
                str_b = str(val_b).strip().upper()
                if 'LISTADO' in str_a and 'CEDULA' in str_b:
                    start_row = r
                    break
                    
        if start_row is None:
            # Fallback
            for r in range(1, 100):
                val_a = sheet.cell(row=r, column=1).value
                if val_a and 'LISTADO' in str(val_a).strip().upper():
                    start_row = r
                    break
                    
        if start_row is None:
            return {}
            
        ignore_subject_keywords = {"promedio", "total", "observacion", "observación", "estado", "firma", "comportamiento"}
        
        # Detectar si es formato de materia única (donde Col C es T1, Col E es T2, Col G es T3)
        es_materia_unica = False
        header_c = str(sheet.cell(row=start_row, column=3).value or "").strip()
        
        if sheet.max_column >= 5:
            header_d = str(sheet.cell(row=start_row, column=4).value or "").strip().upper()
            header_e = str(sheet.cell(row=start_row, column=5).value or "").strip().upper()
            if any(k in header_d for k in ["EXAM", "PROY", "EVAL"]) or any(k in header_e for k in ["TRIMESTRE 2", "T2", "EXAM", "EVAL"]):
                es_materia_unica = True
                
        subject_columns = []
        if es_materia_unica:
            periodo_detectado = str(sheet.cell(row=6, column=2).value or "").strip().upper()
            col_to_read = 3
            if "TRIMESTRE 2" in periodo_detectado or "T2" in periodo_detectado or "SEGUNDO" in periodo_detectado:
                col_to_read = 5
            elif "TRIMESTRE 3" in periodo_detectado or "T3" in periodo_detectado or "TERCER" in periodo_detectado:
                col_to_read = 7
            
            subject_name = header_c if header_c else "Materia"
            subject_columns.append((col_to_read, subject_name))
            print(f"[cargar_excel_datos] Formato materia única: {subject_name}, leyendo columna {col_to_read} ({periodo_detectado})", file=sys.stderr)
        else:
            # Formato estándar multiasignatura
            for col in range(3, sheet.max_column + 1):
                header_value = sheet.cell(row=start_row, column=col).value
                if header_value is None or str(header_value).strip() == "":
                    continue
                subject_name = str(header_value).strip()
                subject_norm = normalizar_columna(subject_name)
                if any(keyword in subject_norm for keyword in ignore_subject_keywords):
                    continue
                subject_columns.append((col, subject_name))

        print(f"[cargar_excel_datos] Asignaturas detectadas: {[name for _, name in subject_columns]}", file=sys.stderr)

        records = {}
        # Leer a partir de la fila siguiente
        for r in range(start_row + 1, sheet.max_row + 1):
            val_a = sheet.cell(row=r, column=1).value
            val_b = sheet.cell(row=r, column=2).value
            
            # Detenerse en fila vacía o sin nombre en columna A
            if val_a is None or str(val_a).strip() == "":
                break
                
            nombre = str(val_a).strip()
            cedula = ""
            if val_b is not None:
                val_b_str = str(val_b).strip()
                if val_b_str.endswith('.0'):
                    val_b_str = val_b_str[:-2]
                if val_b_str.isdigit() and len(val_b_str) == 9:
                    val_b_str = "0" + val_b_str
                cedula = val_b_str
                
            if not cedula:
                cedula = nombre

            notas = {}
            for col, subject_name in subject_columns:
                grade_value = sheet.cell(row=r, column=col).value
                if grade_value is None or str(grade_value).strip() == "":
                    continue
                try:
                    notas[subject_name] = truncar_2_decimales(float(grade_value))
                except (TypeError, ValueError):
                    continue
                
            records[cedula] = {
                "nombre": nombre,
                "cedula": cedula,
                "notas": notas
            }
            
        return records
    except Exception as e:
        print(f"Error al cargar excel {file_path}: {e}", file=sys.stderr)
        return {}

def consolidar_estudiantes(t1_path, t2_path, t3_path, su_path) -> list[dict]:
    t1_data = cargar_excel_datos(t1_path)
    t2_data = cargar_excel_datos(t2_path)
    t3_data = cargar_excel_datos(t3_path)
    su_data = cargar_excel_datos(su_path)
    
    # Unir todas las cédulas/estudiantes únicos
    all_keys = set(t1_data.keys()) | set(t2_data.keys()) | set(t3_data.keys()) | set(su_data.keys())
    
    lista_consolidada = []
    
    for key in all_keys:
        # Recuperar nombres
        nombre = ""
        for src in [t1_data, t2_data, t3_data, su_data]:
            if key in src:
                nombre = src[key]["nombre"]
                break
        
        # Consolidar asignaturas
        all_subjects = set()
        for src in [t1_data, t2_data, t3_data, su_data]:
            if key in src:
                all_subjects.update(src[key]["notas"].keys())
                
        subjects_grades = {}
        for sub in all_subjects:
            g1 = t1_data.get(key, {}).get("notas", {}).get(sub, 0.0)
            g2 = t2_data.get(key, {}).get("notas", {}).get(sub, 0.0)
            g3 = t3_data.get(key, {}).get("notas", {}).get(sub, 0.0)
            su_grade = su_data.get(key, {}).get("notas", {}).get(sub, None) if su_path else None
            
            p_anual = calcular_promedio_anual(g1, g2, g3)
            nota_final = calcular_resultado_con_supletorio(p_anual, su_grade)
            estado = determinar_estado_materia(nota_final, su_grade)
            
            subjects_grades[sub] = {
                "t1": g1,
                "t2": g2,
                "t3": g3,
                "promedio_anual": p_anual,
                "supletorio": su_grade,
                "nota_final": nota_final,
                "estado": estado
            }
            
        # Calcular promedio general
        if subjects_grades:
            final_grades_list = [val["nota_final"] for val in subjects_grades.values()]
            prom_general = sum(final_grades_list) / len(final_grades_list)
        else:
            prom_general = 0.0
            
        # Determinar estado general del estudiante
        tiene_supletorio = any(val["estado"] == "SUPLETORIO" for val in subjects_grades.values())
        tiene_reprobado = any(val["estado"] == "REPROBADO" for val in subjects_grades.values())
        
        if tiene_reprobado:
            estado_general = "REPROBADO"
        elif tiene_supletorio:
            estado_general = "SUPLETORIO"
        else:
            estado_general = "APROBADO"
            
        lista_consolidada.append({
            "cedula": key if len(key) < 15 else "S/C",  # Mostrar sin cédula si se usó el nombre largo como ID
            "id_real": key,
            "nombre": nombre,
            "promedio": prom_general,
            "estado": estado_general,
            "materias": subjects_grades
        })
        
    # Ordenar por nombre
    lista_consolidada.sort(key=lambda x: x["nombre"])
    return lista_consolidada

# 3. Generación del Reporte PDF con ReportLab
def generar_boletin_pdf(datos_consolidados, institucion_info, logos_paths, output_path):
    # Configuración de hoja y márgenes
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    
    # Colores personalizados (Estilo Premium)
    c_primary = colors.HexColor("#312e81")   # Indigo oscuro
    c_text = colors.HexColor("#1e293b")      # Slate 800
    c_light = colors.HexColor("#f8fafc")     # Slate 50
    c_border = colors.HexColor("#cbd5e1")    # Slate 300
    
    # Estilos de Texto
    styles.add(ParagraphStyle(
        name='MainTitle',
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        textColor=c_primary,
        alignment=1 # Center
    ))
    
    styles.add(ParagraphStyle(
        name='SubTitle',
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#475569"),
        alignment=1 # Center
    ))
    
    styles.add(ParagraphStyle(
        name='TableText',
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        textColor=c_text
    ))

    styles.add(ParagraphStyle(
        name='TableTextBold',
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=c_text
    ))

    story = []
    
    for i, est in enumerate(datos_consolidados):
        # 1. Encabezado con Logos
        logo_left = None
        logo_right = None
        
        # Validar y cargar imagen izquierda
        if logos_paths.get("logo1") and os.path.exists(logos_paths["logo1"]):
            try:
                logo_left = Image(logos_paths["logo1"], width=55, height=55)
            except:
                pass
        # Validar y cargar imagen derecha
        if logos_paths.get("logo2") and os.path.exists(logos_paths["logo2"]):
            try:
                logo_right = Image(logos_paths["logo2"], width=55, height=55)
            except:
                pass
                
        # Texto central del encabezado
        text_encabezado = [
            Paragraph(institucion_info["nombre"].upper(), styles['MainTitle']),
            Spacer(1, 2),
            Paragraph("REPORTE DE CALIFICACIONES TRIMESTRALES", styles['SubTitle']),
            Paragraph(f"Año Lectivo: {institucion_info['anio']}", styles['SubTitle']),
            Paragraph(f"Código AMIE: {institucion_info['amie']}", styles['SubTitle'])
        ]
        
        # Tabla del encabezado
        header_data = [[logo_left or "", text_encabezado, logo_right or ""]]
        header_table = Table(header_data, colWidths=[65, 410, 65])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        
        story.append(header_table)
        story.append(Spacer(1, 8))
        
        # 2. Información del Estudiante
        tutor_nombre = institucion_info.get("tutor") or "N/A"
        info_data = [
            [
                Paragraph(f"<b>Estudiante:</b> {est['nombre']}", styles['TableText']),
                Paragraph(f"<b>Cédula:</b> {est['cedula']}", styles['TableText'])
            ],
            [
                Paragraph(f"<b>Curso:</b> {institucion_info['grado']}", styles['TableText']),
                Paragraph(f"<b>Paralelo:</b> {institucion_info['paralelo']}", styles['TableText'])
            ],
            [
                Paragraph(f"<b>Jornada:</b> {institucion_info['jornada']}", styles['TableText']),
                Paragraph(f"<b>Tutor/a:</b> {tutor_nombre}", styles['TableText'])
            ]
        ]
        
        info_table = Table(info_data, colWidths=[270, 270])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), c_light),
            ('BOX', (0,0), (-1,-1), 1, c_border),
            ('INNERGRID', (0,0), (-1,-1), 0.5, c_border),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
        ]))
        
        story.append(info_table)
        story.append(Spacer(1, 10))
        
        # 3. Tabla de Calificaciones
        table_headers = [
            Paragraph("<b>Asignatura / Área</b>", styles['TableTextBold']),
            Paragraph("<b>Trim 1</b>", styles['TableTextBold']),
            Paragraph("<b>Trim 2</b>", styles['TableTextBold']),
            Paragraph("<b>Trim 3</b>", styles['TableTextBold']),
            Paragraph("<b>Prom Anual</b>", styles['TableTextBold']),
            Paragraph("<b>Supletorio</b>", styles['TableTextBold']),
            Paragraph("<b>Nota Final</b>", styles['TableTextBold']),
            Paragraph("<b>Estado</b>", styles['TableTextBold'])
        ]
        
        grades_data = [table_headers]
        
        def nota_pdf(valor):
            if valor is None or str(valor).strip() in ["", "PENDIENTE"]:
                return None
            try:
                return float(valor)
            except (TypeError, ValueError):
                return None

        for sub, grades in est["materias"].items():
            t1_val = nota_pdf(grades.get('t1'))
            t2_val = nota_pdf(grades.get('t2'))
            t3_val = nota_pdf(grades.get('t3'))
            pa_val = nota_pdf(grades.get('promedio_anual'))
            su_val = nota_pdf(grades.get('supletorio'))
            nf_val = nota_pdf(grades.get('nota_final'))
                
            t1_str = f"{t1_val:.2f}" if t1_val is not None else "PEND."
            t2_str = f"{t2_val:.2f}" if t2_val is not None else "PEND."
            t3_str = f"{t3_val:.2f}" if t3_val is not None else "PEND."
            pa_str = f"{pa_val:.2f}" if pa_val is not None else "PEND."
            su_str = f"{su_val:.2f}" if su_val is not None else "-"
            nf_str = f"{nf_val:.2f}" if nf_val is not None else "PEND."
            
            # Estilos de color para el estado
            est_materia = grades.get('estado', 'APROBADO')
            if est_materia == "APROBADO":
                est_color = colors.HexColor("#065f46") # Emerald oscuro
            elif est_materia == "SUPLETORIO":
                est_color = colors.HexColor("#92400e") # Amber oscuro
            else:
                est_color = colors.HexColor("#991b1b") # Rose oscuro
                
            grades_data.append([
                Paragraph(sub, styles['TableTextBold']),
                Paragraph(t1_str, styles['TableText']),
                Paragraph(t2_str, styles['TableText']),
                Paragraph(t3_str, styles['TableText']),
                Paragraph(pa_str, styles['TableText']),
                Paragraph(su_str, styles['TableText']),
                Paragraph(nf_str, styles['TableTextBold']),
                Paragraph(f"<font color='{est_color}'><b>{est_materia}</b></font>", styles['TableText'])
            ])
            
        grades_table = Table(grades_data, colWidths=[180, 50, 50, 50, 60, 55, 55, 60])
        
        # Aplicar estilos a la tabla
        t_style = [
            ('BACKGROUND', (0,0), (-1,0), c_primary),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('ALIGN', (0,1), (0,-1), 'LEFT'), # Alinear nombres de asignaturas a la izquierda
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, c_border),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]
        
        # Pintar cabecera de blanco para el texto
        for idx in range(len(table_headers)):
            table_headers[idx].style.textColor = colors.white
            
        # Alternar colores de filas
        for r_idx in range(1, len(grades_data)):
            if r_idx % 2 == 0:
                t_style.append(('BACKGROUND', (0, r_idx), (-1, r_idx), c_light))
                
        grades_table.setStyle(TableStyle(t_style))
        story.append(grades_table)
        story.append(Spacer(1, 10))
        
        # 4. Resumen Final
        status_color = "#047857" if est['estado'] == "APROBADO" else ("#b45309" if est['estado'] in ["SUPLETORIO", "PENDIENTE DE NOTAS EXTERNAS"] else "#b91c1c")
        try:
            prom_raw = est.get('promedio')
            prom_val = float(prom_raw) if prom_raw is not None else None
        except:
            prom_val = None
        prom_txt = f"{prom_val:.2f}" if prom_val is not None else "PENDIENTE"
            
        summary_data = [
            [
                Paragraph(f"<b>PROMEDIO GENERAL DEL ESTUDIANTE:</b> {prom_txt}", styles['TableTextBold']),
                Paragraph(f"<b>ESTADO FINAL:</b> <font color='{status_color}'><b>{est['estado']}</b></font>", styles['TableTextBold'])
            ]
        ]
        summary_table = Table(summary_data, colWidths=[270, 270])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f1f5f9")),
            ('BOX', (0,0), (-1,-1), 1, c_border),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 35))
        
        # 5. Firmas (Tutor y Rector)
        firmas_data = [
            [
                Paragraph("________________________________________", styles['SubTitle']),
                Paragraph("________________________________________", styles['SubTitle'])
            ],
            [
                Paragraph(f"<b>{tutor_nombre}</b><br/>Tutor / Docente", styles['SubTitle']),
                Paragraph(f"<b>{institucion_info['rector']}</b><br/>Rector / Director", styles['SubTitle'])
            ]
        ]
        firmas_table = Table(firmas_data, colWidths=[270, 270])
        firmas_table.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ]))
        
        story.append(KeepTogether([firmas_table]))
        
        # Añadir salto de página si no es el último estudiante
        if i < len(datos_consolidados) - 1:
            story.append(PageBreak())
            
    doc.build(story)


def extraer_datos_institucionales(file_path: str) -> dict:
    if not file_path or str(file_path).strip() == "":
        raise ValueError("Ruta inválida: no se recibió la ruta del archivo Excel.")

    file_path = str(file_path).strip()
    print(f"[extraer_datos_institucionales] Ruta recibida: {file_path}", file=sys.stderr)

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Ruta inválida: el archivo Excel no existe: {file_path}")

    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        raise ValueError(f"Excel corrupto o ilegible: no se pudo abrir el archivo. Detalle: {e}") from e

    sheet = obtener_hoja_principal(wb)
    if sheet is None:
        hojas = ", ".join(wb.sheetnames) or "sin hojas"
        raise ValueError(f"Formato incorrecto: no existe la hoja obligatoria 'Reporte Periodo'. Hojas encontradas: {hojas}")

    print(f"[extraer_datos_institucionales] Hoja usada: {sheet.title}", file=sys.stderr)

    valores_log = {
        "A1": obtener_valor_celda(sheet, "A1"),
        "D1/D2": obtener_valor_celda(sheet, "D1", ["D2"]),
        "B7": obtener_valor_celda(sheet, "B7"),
        "B3": obtener_valor_celda(sheet, "B3"),
        "B5": obtener_valor_celda(sheet, "B5"),
        "B4": obtener_valor_celda(sheet, "B4"),
        "D3": obtener_valor_celda(sheet, "D3"),
        "D4": obtener_valor_celda(sheet, "D4"),
        "B6": obtener_valor_celda(sheet, "B6"),
    }
    for ref, valor in valores_log.items():
        print(f"[extraer_datos_institucionales] {ref}: {valor}", file=sys.stderr)

    datos = {
        "nombreInstitucion": valores_log["A1"],
        "nivel": "",
        "anioLectivo": valores_log["D1/D2"],
        "codigoAmie": valores_log["B7"],
        "gradoCurso": valores_log["B3"],
        "paralelo": valores_log["B5"],
        "jornada": valores_log["B4"],
        "rectorDirector": valores_log["D3"],
        "tutorCurso": valores_log["D4"],
        "periodoExcel": valores_log["B6"],
    }

    campos_obligatorios = {
        "nombreInstitucion": "Nombre de la Institución (A1)",
        "anioLectivo": "Año Lectivo (D1/D2)",
        "codigoAmie": "Código AMIE (B7)",
        "gradoCurso": "Grado/Curso (B3)",
        "paralelo": "Paralelo (B5)",
        "jornada": "Jornada (B4)",
        "rectorDirector": "Rector/Director (D3)",
        "periodoExcel": "Período (B6)",
    }
    vacios = [descripcion for key, descripcion in campos_obligatorios.items() if not datos.get(key)]
    if vacios:
        raise ValueError("Formato incorrecto: celdas obligatorias vacías: " + ", ".join(vacios))

    print(f"[extraer_datos_institucionales] datosInstitucion final: {json.dumps(datos, ensure_ascii=False)}", file=sys.stderr)
    return datos


# 4. Inyección de datos en Plantillas HTML de Certificados
import re

PLANTILLAS_PERMITIDAS = {
    "FORMATO INICIAL 1.html",
    "FORMATO INICIAL 2.html",
    "PRIMERO DE EGB.html",
    "FORMALO DE ELEMENTAL.html",
    "FORMATO EGBM.html",
    "FORMATO EGBS.html",
    "FORMATO DE 1 Y 2 DE BGU.html",
    "FORMATO DE 3 DE BGU.html",
}

def _normalizar_grado(texto: str) -> str:
    """Normaliza el texto del grado para comparación robusta."""
    import unicodedata
    s = str(texto).upper().strip()
    # Remover tildes
    s = "".join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    # Colapsar espacios múltiples
    s = " ".join(s.split())
    return s

def _es_egb(g: str) -> bool:
    """Verifica si el texto contiene indicador de EGB."""
    return any(k in g for k in ['EGB', 'EDUCACION GENERAL BASICA', 'EDUCACIÓN GENERAL BÁSICA'])

def _es_bachillerato(g: str) -> bool:
    """Verifica si el texto contiene indicador de Bachillerato."""
    return any(k in g for k in ['BACHILLERATO', 'BGU'])

def mapear_plantilla_python(grado_curso: str) -> str | None:
    """
    Mapea el grado/curso a la plantilla HTML correspondiente.
    ORDEN: Inicial → 3ro Bachillerato → 1ro/2do Bachillerato → 1ro EGB → EGB Elemental → EGB Media → EGB Superior → None
    """
    g = _normalizar_grado(grado_curso)
    if not g:
        return None

    # 1. INICIAL 1
    if 'INICIAL 1' in g or 'INICIAL I' in g and 'INICIAL 2' not in g and 'INICIAL II' not in g:
        return 'FORMATO INICIAL 1.html'

    # 2. INICIAL 2
    if 'INICIAL 2' in g or 'INICIAL II' in g:
        return 'FORMATO INICIAL 2.html'

    # 3. 3RO DE BACHILLERATO (antes del genérico BGU)
    if _es_bachillerato(g) and any(k in g for k in ['TERCERO', '3RO', '3ER', 'TERCER']):
        return 'FORMATO DE 3 DE BGU.html'

    # 4. 1RO y 2DO DE BACHILLERATO
    if _es_bachillerato(g):
        return 'FORMATO DE 1 Y 2 DE BGU.html'

    # 5. 1RO DE EGB / PREPARATORIA
    if any(k in g for k in ['PRIMERO DE EGB', '1RO DE EGB', '1ERO DE EGB', '1RO EGB', '1ER GRADO', 'PREPARATORIA']):
        return 'PRIMERO DE EGB.html'

    # 6. 2DO, 3RO y 4TO DE EGB (ELEMENTAL) - requiere contexto EGB
    if 'ELEMENTAL' in g or (_es_egb(g) and any(k in g for k in ['SEGUNDO', 'TERCERO', 'CUARTO', '2DO', '3RO', '4TO'])):
        return 'FORMALO DE ELEMENTAL.html'

    # 7. 5TO, 6TO y 7MO DE EGB (MEDIA) - requiere contexto EGB
    if 'MEDIA' in g or (_es_egb(g) and any(k in g for k in ['QUINTO', 'SEXTO', 'SEPTIMO', '5TO', '6TO', '7MO'])):
        return 'FORMATO EGBM.html'

    # 8. 8VO, 9NO y 10MO DE EGB (SUPERIOR) - requiere contexto EGB
    if 'SUPERIOR' in g or (_es_egb(g) and any(k in g for k in ['OCTAVO', 'NOVENO', 'DECIMO', '8VO', '9NO', '10MO'])):
        return 'FORMATO EGBS.html'

    # Sin coincidencia
    return None

def inject_student_data_html(html_content, student, inst, logos):
    # 1. Reemplazar datos institucionales y del estudiante
    html_content = re.sub(
        r'(NOMBRE DEL ESTUDIANTE:\s*)[^<\n]+', 
        rf'\g<1>{student["nombre"]}', 
        html_content, 
        flags=re.IGNORECASE
    )
    html_content = re.sub(
        r'(CÉDULA:\s*)[^<\n]+', 
        rf'\g<1>{student["cedula"]}', 
        html_content, 
        flags=re.IGNORECASE
    )
    
    grado_completo = f"{inst.get('grado', '')} PARALELO: {inst.get('paralelo', '')}"
    html_content = re.sub(
        r'(GRADO:\s*)[^<\n]+', 
        rf'\g<1>{grado_completo}', 
        html_content, 
        flags=re.IGNORECASE
    )
    
    html_content = re.sub(
        r'(JORNADA:\s*)[^<\n]+', 
        rf'\g<1>{inst.get("jornada", "")}', 
        html_content, 
        flags=re.IGNORECASE
    )
    
    html_content = re.sub(
        r'(AMIE:\s*)[^<\n\s,]+', 
        rf'\g<1>{inst.get("amie", "")}', 
        html_content, 
        flags=re.IGNORECASE
    )
    
    html_content = re.sub(
        r'(AÑO LECTIVO:\s*)[^<\n]+', 
        rf'\g<1>{inst.get("anio", "")}', 
        html_content, 
        flags=re.IGNORECASE
    )

    if inst.get("nombre"):
        html_content = re.sub(
            r'(<h1[^>]*>)[^<]*(EMILIANO HINOSTROZA|UNIDAD EDUCATIVA)[^<]*(</h1>)',
            rf'\g<1>{inst["nombre"]}\g<3>',
            html_content,
            flags=re.IGNORECASE
        )
        
    # 2. Inyectar logos en base64
    if logos.get("logo1"):
        html_content = re.sub(
            r'id="img-logo-1"\s*class="[^"]*"',
            f'id="img-logo-1" class="absolute inset-0 w-full h-full object-contain p-1"',
            html_content
        )
        if 'src=' in html_content.split('id="img-logo-1"')[1].split('>')[0]:
            html_content = re.sub(
                r'id="img-logo-1"\s*src="[^"]*"',
                f'id="img-logo-1" src="{logos["logo1"]}"',
                html_content
            )
        else:
            html_content = html_content.replace('id="img-logo-1"', f'id="img-logo-1" src="{logos["logo1"]}"')
        html_content = html_content.replace('id="text-logo-1" class="', 'id="text-logo-1" class="hidden ')

    if logos.get("logo2"):
        html_content = re.sub(
            r'id="img-logo-2"\s*class="[^"]*"',
            f'id="img-logo-2" class="absolute inset-0 w-full h-full object-contain p-1"',
            html_content
        )
        if 'src=' in html_content.split('id="img-logo-2"')[1].split('>')[0]:
            html_content = re.sub(
                r'id="img-logo-2"\s*src="[^"]*"',
                f'id="img-logo-2" src="{logos["logo2"]}"',
                html_content
            )
        else:
            html_content = html_content.replace('id="img-logo-2"', f'id="img-logo-2" src="{logos["logo2"]}"')
        html_content = html_content.replace('id="text-logo-2" class="', 'id="text-logo-2" class="hidden ')

    # 3. Firmas
    tutor_name = inst.get("tutor", "")
    if tutor_name:
        tutor_pattern = re.compile(r'<div[^>]*class="[^"]*text-center[^"]*"[\s\S]*?<div[^>]*class="[^"]*border-b[^"]*"[\s\S]*?</div>\s*<p[^>]*>\s*(TUTOR/A|TUTOR|TUTORA)\s*</p>\s*</div>', re.IGNORECASE)
        match = tutor_pattern.search(html_content)
        if match:
            block = match.group(0)
            new_block = block.replace(
                'class="border-b border-slate-800 mb-2 h-8 w-full"',
                'class="border-b border-slate-800 mb-2 w-full" style="height: 3.5rem;"'
            )
            if 'cert-nombre-firma' not in new_block:
                new_block = new_block.replace(
                    '<p>',
                    f'<p class="cert-nombre-firma font-bold text-slate-800 mt-1" style="font-size: 0.75rem;">{tutor_name}</p>\n            <p>'
                )
            html_content = html_content.replace(block, new_block)

    rector_name = inst.get("rector", "")
    if rector_name:
        rector_pattern = re.compile(r'<div[^>]*class="[^"]*text-center[^"]*"[\s\S]*?<div[^>]*class="[^"]*border-b[^"]*"[\s\S]*?</div>\s*<p[^>]*>\s*(RECTOR/A|RECTOR|RECTORA|DIRECTOR|DIRECTORA)\s*</p>\s*</div>', re.IGNORECASE)
        match = rector_pattern.search(html_content)
        if match:
            block = match.group(0)
            new_block = block.replace(
                'class="border-b border-slate-800 mb-2 h-8 w-full"',
                'class="border-b border-slate-800 mb-2 w-full" style="height: 3.5rem;"'
            )
            if 'cert-nombre-firma' not in new_block:
                new_block = new_block.replace(
                    '<p>',
                    f'<p class="cert-nombre-firma font-bold text-slate-800 mt-1" style="font-size: 0.75rem;">{rector_name}</p>\n            <p>'
                )
            html_content = html_content.replace(block, new_block)

    return html_content

def inject_subject_grades(html_content, materias_data):
    for sub_name, data in materias_data.items():
        escaped_sub = re.escape(sub_name)
        tr_pattern = re.compile(rf'<tr[^>]*>\s*<td[^>]*>\s*{escaped_sub}\s*</td>[\s\S]*?</tr>', re.IGNORECASE)
        
        match = tr_pattern.search(html_content)
        if match:
            tr_block = match.group(0)
            td_pattern = re.compile(r'<td[^>]*>([\s\S]*?)</td>', re.IGNORECASE)
            tds = list(td_pattern.finditer(tr_block))
            
            if len(tds) >= 2:
                grade_tds = tds[1:]
                fmt = lambda val, is_supletorio=False: f"{float(val):.2f}" if val is not None else ('0.00' if is_supletorio else '-')
                
                new_tds = []
                if len(grade_tds) == 3:
                    new_tds.append(f'<td>{fmt(data.get("t1"))}</td>')
                    new_tds.append(f'<td>{fmt(data.get("t2"))}</td>')
                    new_tds.append(f'<td>{fmt(data.get("t3"))}</td>')
                elif len(grade_tds) == 4:
                    new_tds.append(f'<td>{fmt(data.get("t1"))}</td>')
                    new_tds.append(f'<td>{fmt(data.get("t2"))}</td>')
                    new_tds.append(f'<td>{fmt(data.get("t3"))}</td>')
                    val_final = data.get("nota_final") if data.get("nota_final") is not None else data.get("promedio_anual")
                    new_tds.append(f'<td class="font-bold bg-slate-50">{fmt(val_final)}</td>')
                elif len(grade_tds) == 5:
                    new_tds.append(f'<td>{fmt(data.get("t1"))}</td>')
                    new_tds.append(f'<td>{fmt(data.get("t2"))}</td>')
                    new_tds.append(f'<td>{fmt(data.get("t3"))}</td>')
                    new_tds.append(f'<td>{fmt(data.get("supletorio"), True)}</td>')
                    val_final = data.get("nota_final") if data.get("nota_final") is not None else data.get("promedio_anual")
                    new_tds.append(f'<td class="font-bold bg-slate-50">{fmt(val_final)}</td>')
                    
                new_tr_block = tr_block[:tds[1].start()] + "".join(new_tds) + "</tr>"
                html_content = html_content.replace(tr_block, new_tr_block)
                
    # Reemplazar promedio general
    prom_pattern = re.compile(r'<tr[^>]*>[\s\S]*?<td[^>]*>\s*(PROMEDIO ANUAL|PROMEDIO GENERAL)\s*</td>[\s\S]*?</tr>', re.IGNORECASE)
    for m in prom_pattern.finditer(html_content):
        tr_block = m.group(0)
        td_pattern = re.compile(r'<td[^>]*>([\s\S]*?)</td>', re.IGNORECASE)
        tds = list(td_pattern.finditer(tr_block))
        if tds:
            try:
                grades = [float(v.get("nota_final", 0.0) or v.get("promedio_anual", 0.0)) for v in materias_data.values() if v.get("nota_final") is not None or v.get("promedio_anual") is not None]
                promedio_val = f"{(sum(grades) / len(grades)):.2f}" if grades else "-"
            except Exception:
                promedio_val = "-"
                
            last_td = tds[-1]
            new_tr_block = tr_block[:last_td.start()] + f'<td class="font-bold bg-slate-100">{promedio_val}</td>' + "</tr>"
            html_content = html_content.replace(tr_block, new_tr_block)

    # Reemplazar cabecera "PROMEDIO ANUAL" a "Nota" en th
    html_content = re.sub(
        r'(<th[^>]*>\s*)PROMEDIO ANUAL(\s*</th>)',
        r'\g<1>Nota\g<2>',
        html_content,
        flags=re.IGNORECASE
    )
    return html_content

def generar_certificados_inicial(payload):
    estudiantes = payload.get("datos_consolidados", [])
    inst = payload.get("institucion", {})
    logos = payload.get("logos", {})
    
    # Directorio de plantillas originales (solo lectura)
    templates_dir = os.path.join(os.path.dirname(__file__), "assets", "certificados")
    
    # Directorio de salida para certificados generados (recibido de Electron o fallback)
    cert_output_dir = payload.get("certOutputDir", os.path.join(os.path.dirname(__file__), "assets", "certificados_generados"))
    os.makedirs(cert_output_dir, exist_ok=True)
    
    # Plantilla: JS tiene prioridad, luego fallback a mapear_plantilla_python
    plantilla_name_js = payload.get("plantillaName")
    grado_canonico = payload.get("gradoCursoCanonico", inst.get("grado", ""))
    
    # Validar plantilla del JS contra lista blanca
    if plantilla_name_js and plantilla_name_js in PLANTILLAS_PERMITIDAS:
        plantilla_name = plantilla_name_js
        print(f"[Certificados-PY] Usando plantilla de JS: {plantilla_name}", file=sys.stderr)
    else:
        plantilla_name = mapear_plantilla_python(grado_canonico)
        print(f"[Certificados-PY] Plantilla calculada por Python: {plantilla_name} (grado: {grado_canonico})", file=sys.stderr)
    
    if not plantilla_name:
        return {"error": f"No se encontró plantilla para el grado: {grado_canonico}", "supletorios": []}
    
    plantilla_path = os.path.join(templates_dir, plantilla_name)
    if not os.path.exists(plantilla_path):
        return {"error": f"Plantilla no encontrada en disco: {plantilla_name} → {plantilla_path}", "supletorios": []}
    
    print(f"[Certificados-PY] Plantilla final: {plantilla_name}", file=sys.stderr)
    print(f"[Certificados-PY] Ruta: {plantilla_path}", file=sys.stderr)
    print(f"[Certificados-PY] Existe: {os.path.exists(plantilla_path)}", file=sys.stderr)
    print(f"[Certificados-PY] Estudiantes: {len(estudiantes)}", file=sys.stderr)
    print(f"[Certificados-PY] Output dir: {cert_output_dir}", file=sys.stderr)
    
    with open(plantilla_path, "r", encoding="utf-8") as f:
        template_html = f.read()
    
    supletorios_detectados = []
    archivos_generados = []
    
    for est in estudiantes:
        materias = est.get("materias", {})
        for sub_name, m_data in list(materias.items()):
            g1 = float(m_data.get("t1") or 0.0)
            g2 = float(m_data.get("t2") or 0.0)
            g3 = float(m_data.get("t3") or 0.0)
            p_anual = calcular_promedio_anual(g1, g2, g3)
            m_data["promedio_anual"] = p_anual
            if m_data.get("nota_final") is None:
                m_data["nota_final"] = p_anual
                
            if 4.01 <= p_anual <= 6.99:
                nivel_str = inst.get("nivel", "")
                es_egb_media = "MEDIA" in str(inst.get("grado", "")).upper() or "MEDIA" in str(nivel_str).upper()
                supletorios_detectados.append({
                    "id": est["id_real"],
                    "nombre": est["nombre"],
                    "curso": inst.get("grado", ""),
                    "asignatura": sub_name,
                    "nivel": nivel_str,
                    "es_egb_media": es_egb_media
                })
        
        student_html = inject_student_data_html(template_html, est, inst, logos)
        student_html = inject_subject_grades(student_html, materias)
        
        # Guardar en directorio de salida, NO en plantillas
        output_filename = f"certificado_{est['id_real']}.html"
        output_path = os.path.join(cert_output_dir, output_filename)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(student_html)
        
        archivos_generados.append(output_path)
            
    return {
        "supletorios": supletorios_detectados,
        "archivos": archivos_generados,
        "plantilla_usada": plantilla_name,
        "cert_output_dir": cert_output_dir
    }

def actualizar_certificado_supletorio(student_id, asignatura, nota_supletorio, cert_output_dir=None):
    """Actualiza un certificado HTML existente con la nota de supletorio."""
    # Buscar en el directorio de salida proporcionado, o fallback
    if cert_output_dir:
        file_path = os.path.join(cert_output_dir, f"certificado_{student_id}.html")
    else:
        # Fallback: buscar en assets/certificados_generados y luego en assets/certificados (legacy)
        base_dir = os.path.dirname(__file__)
        file_path = os.path.join(base_dir, "assets", "certificados_generados", f"certificado_{student_id}.html")
        if not os.path.exists(file_path):
            file_path = os.path.join(base_dir, "assets", "certificados", f"certificado_{student_id}.html")
    
    if not os.path.exists(file_path):
        print(f"[actualizar_certificado_supletorio] No existe el certificado para {student_id} en {file_path}", file=sys.stderr)
        return False
        
    with open(file_path, "r", encoding="utf-8") as f:
        html_content = f.read()
        
    try:
        val_su = float(nota_supletorio)
        val_final = 7.00 if val_su >= 7.00 else val_su
    except (ValueError, TypeError):
        val_su = 0.0
        val_final = 0.0
        
    escaped_sub = re.escape(asignatura)
    tr_pattern = re.compile(rf'<tr[^>]*>\s*<td[^>]*>\s*{escaped_sub}\s*</td>[\s\S]*?</tr>', re.IGNORECASE)
    
    match = tr_pattern.search(html_content)
    if match:
        tr_block = match.group(0)
        td_pattern = re.compile(r'<td[^>]*>([\s\S]*?)</td>', re.IGNORECASE)
        tds = list(td_pattern.finditer(tr_block))
        
        if len(tds) == 6:
            supletorio_td = tds[4]
            final_td = tds[5]
            new_tr_block = (
                tr_block[:supletorio_td.start()] +
                f'<td>{val_su:.2f}</td>' +
                f'<td class="font-bold bg-slate-50">{val_final:.2f}</td>' +
                '</tr>'
            )
            html_content = html_content.replace(tr_block, new_tr_block)
        elif len(tds) == 5:
            final_td = tds[4]
            new_tr_block = (
                tr_block[:final_td.start()] +
                f'<td class="font-bold bg-slate-50">{val_final:.2f}</td>' +
                '</tr>'
            )
            html_content = html_content.replace(tr_block, new_tr_block)
        else:
            return False
            
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        return True
    return False


# 5. Proceso CLI Principal
def main():
    parser = argparse.ArgumentParser(description="Procesador de Notas Académicas de la UEEH")
    parser.add_argument('--analizar', action='store_true', help='Analiza los excels y devuelve la lista de estudiantes en JSON')
    parser.add_argument('--generar', action='store_true', help='Genera los PDF recibiendo el payload por stdin')
    parser.add_argument('--certificados', action='store_true', help='Genera los certificados HTML iniciales y detecta supletorios')
    parser.add_argument('--supletorios', type=str, help='JSON con las notas manuales de supletorio para actualizar')
    parser.add_argument('--t1', type=str, help='Ruta Excel Trimestre 1')
    parser.add_argument('--t2', type=str, help='Ruta Excel Trimestre 2')
    parser.add_argument('--t3', type=str, help='Ruta Excel Trimestre 3')
    parser.add_argument('--su', type=str, help='Ruta Excel Supletorio')
    
    args = parser.parse_args()
    
    if args.supletorios:
        try:
            updates = json.loads(args.supletorios)
            success_count = 0
            # Extract cert_output_dir from the first entry if provided
            cert_output_dir = None
            if isinstance(updates, list) and len(updates) > 0:
                cert_output_dir = updates[0].get("cert_output_dir")
            if isinstance(updates, list):
                for up in updates:
                    student_id = up.get("id")
                    asignatura = up.get("asignatura")
                    nota_su = up.get("nota_supletorio")
                    if student_id and asignatura and nota_su is not None:
                        if actualizar_certificado_supletorio(student_id, asignatura, nota_su, cert_output_dir):
                            success_count += 1
            print(json.dumps({"success": True, "updated": success_count}, ensure_ascii=False))
        except Exception as e:
            print(json.dumps({"success": False, "error": str(e)}, ensure_ascii=False))
        return

    if args.certificados:
        try:
            # Leer el payload JSON enviado desde Electron a través de stdin
            input_data = sys.stdin.read()
            if not input_data:
                print(json.dumps({"success": False, "error": "No se recibió payload en stdin"}, ensure_ascii=False))
                return
                
            payload = json.loads(input_data)
            result = generar_certificados_inicial(payload)
            
            # result is now a dict with supletorios, archivos, plantilla_usada, cert_output_dir
            if isinstance(result, dict) and "error" in result:
                print(json.dumps({"success": False, "error": result["error"]}, ensure_ascii=False))
            else:
                print(json.dumps({"success": True, **result}, ensure_ascii=False))
        except Exception as e:
            print(json.dumps({"success": False, "error": str(e)}, ensure_ascii=False))
        return

    elif args.analizar:
        try:
            first_path = args.t1 or args.t2 or args.t3 or args.su
            datos_inst = extraer_datos_institucionales(first_path)
            
            # Cargar datos del excel específico
            records = cargar_excel_datos(first_path)
            
            # Obtener asignaturas
            asignaturas = []
            if records:
                first_student = next(iter(records.values()))
                asignaturas = list(first_student["notas"].keys())
                
            estudiantes = list(records.values())
            
            # Calcular promedios para la tabla
            estudiantes_tabla = []
            for est in estudiantes:
                grades_list = [v for v in est["notas"].values() if isinstance(v, (int, float))]
                prom = sum(grades_list) / len(grades_list) if grades_list else 0.0
                estudiantes_tabla.append({
                    "cedula": est["cedula"],
                    "nombre": est["nombre"],
                    "promedio": truncar_2_decimales(prom),
                    "estado": "APROBADO" if prom >= 7.0 else "SUPLETORIO"
                })
            
            output_data = {
                "datosInstitucion": datos_inst,
                "periodoExcel": datos_inst.get("periodoExcel", ""),
                "asignaturas": asignaturas,
                "estudiantes": estudiantes,
                "estudiantes_tabla": estudiantes_tabla
            }
            print(json.dumps(output_data, ensure_ascii=False))
        except Exception as e:
            print(json.dumps({"error": str(e)}, ensure_ascii=False))
            
    elif args.generar:
        try:
            # Leer el payload JSON enviado desde Electron a través de stdin
            input_data = sys.stdin.read()
            if not input_data:
                print(json.dumps({"success": False, "error": "No se recibió payload en stdin"}, ensure_ascii=False))
                return
                
            payload = json.loads(input_data)
            
            excels = payload.get("excels", {})
            logos = payload.get("logos", {})
            inst = payload.get("institucion", {})
            estudiantes_seleccionados = payload.get("estudiantes", [])
            
            # Consolidar todos los estudiantes de los excels
            if payload.get("datos_consolidados"):
                seleccionados_data = payload.get("datos_consolidados")
            else:
                datos_completos = consolidar_estudiantes(
                    excels.get("t1"), excels.get("t2"), excels.get("t3"), excels.get("su")
                )
                seleccionados_data = [est for est in datos_completos if est["id_real"] in estudiantes_seleccionados]
            
            if not seleccionados_data:
                print(json.dumps({"success": False, "error": "No se encontraron datos para los estudiantes seleccionados"}, ensure_ascii=False))
                return
                
            # Definir la ruta de salida en Descargas del usuario
            home_dir = str(Path.home())
            descargas_path = os.path.join(home_dir, "Downloads")
            if not os.path.exists(descargas_path):
                descargas_path = home_dir # Fallback al home
                
            pdf_filename = f"Boletines_Consolidados_{inst.get('grado', 'Curso').replace(' ', '_')}_{inst.get('paralelo', 'P')}.pdf"
            output_pdf = os.path.join(descargas_path, pdf_filename)
            
            generar_boletin_pdf(seleccionados_data, inst, logos, output_pdf)
            
            print(json.dumps({"success": True, "path": output_pdf}, ensure_ascii=False))
            
        except Exception as e:
            print(json.dumps({"success": False, "error": str(e)}, ensure_ascii=False))
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
