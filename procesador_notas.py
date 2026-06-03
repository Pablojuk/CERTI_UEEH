# -*- coding: utf-8 -*-
"""
Script de procesamiento académico y generación de boletines PDF para la UEEH.
Cruza las notas de archivos Excel trimestrales y supletorios.
"""

from __future__ import annotations
import os
import sys
import json
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
        if any(k in norm for k in ["estudiante", "nombre", "apellidos", "alumno"]):
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

def cargar_excel_datos(file_path: str) -> dict[str, dict]:
    """Lee el excel y devuelve un diccionario indexado por cédula/nombre con las asignaturas y sus notas."""
    if not file_path or not os.path.exists(file_path):
        return {}
        
    try:
        # Cargar primera hoja del libro
        df = pd.read_excel(file_path, sheet_name=0)
    except Exception as e:
        print(f"Error al leer {file_path}: {e}", file=sys.stderr)
        return {}
        
    id_col, name_col, subject_cols = detectar_columnas(df)
    
    if not name_col and not id_col:
        # Fallback si no detecta nombres o ids, usar las primeras columnas
        if len(df.columns) >= 2:
            name_col = df.columns[1]
            id_col = df.columns[0]
            subject_cols = [c for c in df.columns[2:] if c not in ["Promedio", "Total", "Estado"]]
        else:
            return {}
            
    records = {}
    for _, row in df.iterrows():
        # Capturar el nombre del estudiante
        nombre = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else "Estudiante Desconocido"
        # Capturar la cédula
        cedula = ""
        if id_col and pd.notna(row[id_col]):
            val = row[id_col]
            if isinstance(val, float):
                if val.is_integer():
                    val = int(val)
                else:
                    val = str(val)
            val_str = str(val).strip()
            if val_str.endswith('.0'):
                val_str = val_str[:-2]
            # Si es puramente numérico y tiene 9 dígitos (cédula ecuatoriana sin el 0 inicial), rellenar con 0
            if val_str.isdigit() and len(val_str) == 9:
                val_str = "0" + val_str
            cedula = val_str
            
        if not cedula and nombre:
            # Fallback usar nombre si no hay cédula
            cedula = nombre
            
        if not cedula or nombre == "Estudiante Desconocido":
            continue
            
        student_data = {
            "nombre": nombre,
            "cedula": cedula,
            "notas": {}
        }
        
        for sub in subject_cols:
            val = row[sub]
            try:
                val_float = float(val) if pd.notna(val) else 0.0
            except:
                val_float = 0.0
            student_data["notas"][sub] = val_float
            
        records[cedula] = student_data
        
    return records

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
        
        for sub, grades in est["materias"].items():
            t1_str = f"{grades['t1']:.2f}"
            t2_str = f"{grades['t2']:.2f}"
            t3_str = f"{grades['t3']:.2f}"
            pa_str = f"{grades['promedio_anual']:.2f}"
            su_str = f"{grades['supletorio']:.2f}" if grades['supletorio'] is not None else "-"
            nf_str = f"{grades['nota_final']:.2f}"
            
            # Estilos de color para el estado
            est_materia = grades['estado']
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
        status_color = "#047857" if est['estado'] == "APROBADO" else ("#b45309" if est['estado'] == "SUPLETORIO" else "#b91c1c")
        
        summary_data = [
            [
                Paragraph(f"<b>PROMEDIO GENERAL DEL ESTUDIANTE:</b> {est['promedio']:.2f}", styles['TableTextBold']),
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


# 4. Proceso CLI Principal
def main():
    parser = argparse.ArgumentParser(description="Procesador de Notas Académicas de la UEEH")
    parser.add_argument('--analizar', action='store_true', help='Analiza los excels y devuelve la lista de estudiantes en JSON')
    parser.add_argument('--generar', action='store_true', help='Genera los PDF recibiendo el payload por stdin')
    parser.add_argument('--t1', type=str, help='Ruta Excel Trimestre 1')
    parser.add_argument('--t2', type=str, help='Ruta Excel Trimestre 2')
    parser.add_argument('--t3', type=str, help='Ruta Excel Trimestre 3')
    parser.add_argument('--su', type=str, help='Ruta Excel Supletorio')
    
    args = parser.parse_args()
    
    if args.analizar:
        try:
            estudiantes = consolidar_estudiantes(args.t1, args.t2, args.t3, args.su)
            # Retornar una lista reducida para actualizar la tabla del frontend
            res = []
            for est in estudiantes:
                res.append({
                    "cedula": est["cedula"],
                    "nombre": est["nombre"],
                    "promedio": est["promedio"],
                    "estado": est["estado"]
                })
            print(json.dumps(res, ensure_ascii=False))
        except Exception as e:
            print(json.dumps({"error": str(e)}))
            
    elif args.generar:
        try:
            # Leer el payload JSON enviado desde Electron a través de stdin
            input_data = sys.stdin.read()
            if not input_data:
                print(json.dumps({"success": false, "error": "No se recibió payload en stdin"}))
                return
                
            payload = json.loads(input_data)
            
            excels = payload.get("excels", {})
            logos = payload.get("logos", {})
            inst = payload.get("institucion", {})
            estudiantes_seleccionados = payload.get("estudiantes", [])
            
            # Consolidar todos los estudiantes de los excels
            datos_completos = consolidar_estudiantes(
                excels.get("t1"), excels.get("t2"), excels.get("t3"), excels.get("su")
            )
            
            # Filtrar los estudiantes seleccionados
            seleccionados_data = [est for est in datos_completos if est["id_real"] in estudiantes_seleccionados]
            
            if not seleccionados_data:
                print(json.dumps({"success": False, "error": "No se encontraron datos para los estudiantes seleccionados"}))
                return
                
            # Definir la ruta de salida en Descargas del usuario
            home_dir = str(Path.home())
            descargas_path = os.path.join(home_dir, "Downloads")
            if not os.path.exists(descargas_path):
                descargas_path = home_dir # Fallback al home
                
            pdf_filename = f"Boletines_Consolidados_{inst.get('grado', 'Curso').replace(' ', '_')}_{inst.get('paralelo', 'P')}.pdf"
            output_pdf = os.path.join(descargas_path, pdf_filename)
            
            generar_boletin_pdf(seleccionados_data, inst, logos, output_pdf)
            
            print(json.dumps({"success": True, "path": output_pdf}))
            
        except Exception as e:
            print(json.dumps({"success": False, "error": str(e)}))
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
