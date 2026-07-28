# -*- coding: utf-8 -*-

import re
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from procesador_notas import (
    cargar_excel_datos,
    consolidar_estudiantes,
    generar_certificados_inicial,
    inject_student_data_html,
    inject_subject_grades,
)


CIVICA = "CÍVICA Y ACOMPAÑAMIENTO INTEGRAL EN EL AULA"
PARTICIPACION = "PARTICIPACIÓN ESTUDIANTIL"


class PlantillasNeutrasTests(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.directorio = Path(self.temporal.name)
        self.plantillas = Path(__file__).parents[1] / "assets" / "certificados"

    def tearDown(self):
        self.temporal.cleanup()

    def crear_excel_civica(self, nombre, valor, periodo="TRIMESTRE 1"):
        ruta = self.directorio / nombre
        libro = Workbook()
        hoja = libro.active
        hoja.title = "Reporte Periodo"
        hoja["B3"] = "2DO BGU"
        hoja["B6"] = periodo
        hoja["A9"] = "LISTADO"
        hoja["B9"] = "CEDULA"
        hoja["M9"] = CIVICA
        hoja["A10"] = "ALVARADO CORONEL"
        hoja["B10"] = "0123456789"
        hoja["M10"] = valor
        libro.save(ruta)
        return ruta

    def fila_civica(self, html):
        coincidencia = re.search(
            rf'<tr[^>]*data-subject="{re.escape(CIVICA)}"[^>]*>[\s\S]*?</tr>',
            html,
            re.IGNORECASE,
        )
        self.assertIsNotNone(coincidencia)
        return coincidencia.group(0)

    def valores_civica(self, html):
        fila = self.fila_civica(html)
        return {
            periodo: re.search(
                rf'<td[^>]*data-academic-field="{periodo}"[^>]*>([\s\S]*?)</td>',
                fila,
                re.IGNORECASE,
            ).group(1).strip()
            for periodo in ("t1", "t2", "t3")
        }

    def test_alvarado_excel_b_mas_llega_literal_a_consolidacion(self):
        ruta = self.crear_excel_civica("alvarado_t1.xlsx", "  B+  ")
        registro = cargar_excel_datos(ruta, "2DO BGU")["0123456789"]
        self.assertEqual(registro["nombre"], "ALVARADO CORONEL")
        self.assertEqual(registro["notas"][CIVICA], "B+")
        self.assertEqual(registro["tipos_asignaturas"][CIVICA], "cualitativa")

        consolidado = consolidar_estudiantes(ruta, None, None, None, "2DO BGU")[0]
        self.assertEqual(consolidado["materias"][CIVICA]["tipo"], "cualitativa")
        self.assertEqual(
            (
                consolidado["materias"][CIVICA]["t1"],
                consolidado["materias"][CIVICA]["t2"],
                consolidado["materias"][CIVICA]["t3"],
            ),
            ("B+", "", ""),
        )

    def test_inyeccion_python_conserva_trimestres_y_limpia_ausentes(self):
        plantilla = (self.plantillas / "FORMATO DE 1 Y 2 DE BGU.html").read_text(
            encoding="utf-8"
        )
        html = inject_subject_grades(
            plantilla,
            {CIVICA: {"tipo": "cualitativa", "t1": "B+", "t2": "A-", "t3": "C+"}},
        )
        self.assertEqual(
            self.valores_civica(html),
            {"t1": "B+", "t2": "A-", "t3": "C+"},
        )

        html_sin_materia = inject_subject_grades(plantilla, {})
        self.assertEqual(
            self.valores_civica(html_sin_materia),
            {"t1": "", "t2": "", "t3": ""},
        )

    def test_dos_estudiantes_generados_no_comparten_valores(self):
        salida = self.directorio / "certificados"
        resultado = generar_certificados_inicial({
            "datos_consolidados": [
                {
                    "id_real": "estudiante_1",
                    "cedula": "1",
                    "nombre": "ESTUDIANTE UNO",
                    "materias": {
                        CIVICA: {
                            "tipo": "cualitativa",
                            "t1": "A+",
                            "t2": "",
                            "t3": "",
                        }
                    },
                },
                {
                    "id_real": "estudiante_2",
                    "cedula": "2",
                    "nombre": "ESTUDIANTE DOS",
                    "materias": {
                        CIVICA: {
                            "tipo": "cualitativa",
                            "t1": "B+",
                            "t2": "",
                            "t3": "",
                        }
                    },
                },
            ],
            "institucion": {"grado": "2DO BGU", "nivel": "BGU"},
            "logos": {},
            "gradoCursoCanonico": "2DO BGU",
            "plantillaName": "FORMATO DE 1 Y 2 DE BGU.html",
            "certOutputDir": str(salida),
        })
        primero = Path(resultado["archivos"][0]).read_text(encoding="utf-8")
        segundo = Path(resultado["archivos"][1]).read_text(encoding="utf-8")
        self.assertEqual(self.valores_civica(primero)["t1"], "A+")
        self.assertEqual(self.valores_civica(segundo)["t1"], "B+")

    def test_bgu_1_y_2_omiten_participacion_y_muestran_asistencia_a_ancho_completo(self):
        plantilla = (self.plantillas / "FORMATO DE 1 Y 2 DE BGU.html").read_text(
            encoding="utf-8"
        )
        self.assertNotIn(PARTICIPACION, plantilla)
        self.assertRegex(
            plantilla,
            r'<div data-cert-section="asistencia-anual" class="w-full">',
        )

        for grado in ("1RO BGU", "2DO BGU"):
            with self.subTest(grado=grado):
                salida = self.directorio / f"certificado_{grado[:1]}_sin_participacion"
                resultado = generar_certificados_inicial({
                    "datos_consolidados": [{
                        "id_real": "estudiante_bgu",
                        "cedula": "0102030405",
                        "nombre": "ESTUDIANTE BGU",
                        "materias": {
                            PARTICIPACION: {
                                "tipo": "cuantitativa",
                                "t1": 9.0,
                                "t2": 9.0,
                                "t3": 9.0,
                            }
                        },
                        "asistencia": {
                            "configurada": True,
                            "totalFaltas": 6,
                            "justificadas": 0,
                            "injustificadas": 6,
                            "totalAsistencia": 62,
                        },
                    }],
                    "institucion": {"grado": grado, "nivel": "BGU"},
                    "logos": {},
                    "gradoCursoCanonico": grado,
                    "plantillaName": "FORMATO DE 1 Y 2 DE BGU.html",
                    "certOutputDir": str(salida),
                })
                html = Path(resultado["archivos"][0]).read_text(encoding="utf-8")
                self.assertNotIn(PARTICIPACION, html)
                for campo, valor in (
                    ("registro", "6"),
                    ("justificadas", "0"),
                    ("injustificadas", "6"),
                    ("total", "62"),
                ):
                    self.assertRegex(
                        html,
                        rf'data-asistencia="{campo}"[^>]*>{valor}</td>',
                    )

    def test_todas_las_plantillas_inyectan_tutor_y_rector_con_claves_compatibles(self):
        estudiante = {"nombre": "ESTUDIANTE", "cedula": "0102030405"}
        instituciones = (
            {
                "tutor": "TUTOR DEL CURSO",
                "rector": "RECTOR DE LA INSTITUCIÓN",
            },
            {
                "tutorCurso": "TUTOR DEL CURSO",
                "rectorDirector": "RECTOR DE LA INSTITUCIÓN",
            },
        )
        for ruta in sorted(self.plantillas.glob("*.html")):
            plantilla = ruta.read_text(encoding="utf-8")
            self.assertEqual(plantilla.count('data-cert-field="tutor-name"'), 1)
            self.assertEqual(plantilla.count('data-cert-field="rector-name"'), 1)
            for institucion in instituciones:
                with self.subTest(
                    plantilla=ruta.name,
                    claves=tuple(institucion),
                ):
                    html = inject_student_data_html(
                        plantilla,
                        estudiante,
                        institucion,
                        {},
                    )
                    self.assertRegex(
                        html,
                        r'data-cert-field="tutor-name"[^>]*>TUTOR DEL CURSO</p>',
                    )
                    self.assertRegex(
                        html,
                        r'data-cert-field="rector-name"[^>]*>'
                        r'RECTOR DE LA INSTITUCIÓN</p>',
                    )

    def test_las_ocho_plantillas_son_neutras_y_conservan_estructura(self):
        rutas = sorted(self.plantillas.glob("*.html"))
        self.assertEqual(len(rutas), 8)
        for ruta in rutas:
            html = ruta.read_text(encoding="utf-8")
            with self.subTest(plantilla=ruta.name):
                self.assertNotIn("[cite:", html)
                self.assertIn("<table", html)
                self.assertIn("data-cert-field=\"institution-name\"", html)
                self.assertIn("data-cert-field=\"student-name\"", html)
                self.assertIn("data-cert-field=\"student-id\"", html)
                self.assertIn("data-cert-field=\"grade\"", html)
                self.assertIn("data-cert-field=\"schedule\"", html)
                self.assertIn(CIVICA, html)

                campos = re.findall(
                    r'<(?:td|span|div|h1)\b[^>]*(?:data-academic-value="true"|'
                    r'data-cert-field="[^"]+")[^>]*>([\s\S]*?)</(?:td|span|div|h1)>',
                    html,
                    re.IGNORECASE,
                )
                self.assertGreater(len(campos), 0)
                self.assertTrue(all(not re.sub(r"<[^>]+>", "", valor).strip() for valor in campos))


if __name__ == "__main__":
    unittest.main()
