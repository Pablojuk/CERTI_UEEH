# -*- coding: utf-8 -*-

import re
import tempfile
import unittest
from pathlib import Path

from procesador_notas import (
    generar_boletin_pdf,
    generar_certificados_inicial,
    inject_asistencia_anual,
)


class AsistenciaReportesTests(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.directorio = Path(self.temporal.name)
        self.plantillas = Path(__file__).parents[1] / "assets" / "certificados"

    def tearDown(self):
        self.temporal.cleanup()

    def extraer_valor(self, html, periodo, campo):
        patron = re.compile(
            rf'<td\b'
            rf'(?=[^>]*data-asistencia-periodo="{periodo}")'
            rf'(?=[^>]*data-asistencia-campo="{campo}")'
            rf'[^>]*>([\s\S]*?)</td>',
            re.IGNORECASE,
        )
        coincidencia = patron.search(html)
        self.assertIsNotNone(coincidencia, f"Falta {periodo}/{campo}")
        return re.sub(r"<[^>]+>", "", coincidencia.group(1)).strip()

    @staticmethod
    def resumen_pruebas():
        return {
            "cursoId": "curso_a",
            "estudianteId": "estudiante_a",
            "T1": {
                "configurado": True,
                "totalFaltas": 3,
                "justificadas": 0,
                "injustificadas": 3,
                "totalAsistencia": 19,
            },
            "T2": {
                "configurado": True,
                "totalFaltas": 2,
                "justificadas": 1,
                "injustificadas": 1,
                "totalAsistencia": 20,
            },
            "T3": {
                "configurado": True,
                "totalFaltas": 1,
                "justificadas": 0,
                "injustificadas": 1,
                "totalAsistencia": 21,
            },
            "anual": {
                "configurada": True,
                "totalFaltas": 6,
                "justificadas": 1,
                "injustificadas": 5,
                "totalAsistencia": 60,
            },
        }

    def test_todas_las_plantillas_de_asistencia_tienen_identificadores_y_sin_ejemplos(self):
        encontradas = 0
        for ruta in self.plantillas.glob("*.html"):
            html = ruta.read_text(encoding="utf-8")
            if "data-asistencia-periodo" not in html:
                continue
            encontradas += 1
            with self.subTest(plantilla=ruta.name):
                for periodo in ("T1", "T2", "T3", "ANUAL"):
                    for campo in (
                        "registro",
                        "justificacion",
                        "injustificado",
                        "total",
                    ):
                        self.assertEqual(
                            self.extraer_valor(html, periodo, campo),
                            "",
                        )
                for etiqueta in (
                    "PRIMER TRIMESTRE",
                    "SEGUNDO TRIMESTRE",
                    "TERCER TRIMESTRE",
                    "TOTAL ANUAL",
                ):
                    self.assertIn(etiqueta, html)
        self.assertEqual(encontradas, 6)

    def test_inyecta_trimestres_y_acumulado_anual_en_las_celdas_correctas(self):
        plantilla = (self.plantillas / "FORMATO DE 3 DE BGU.html").read_text(encoding="utf-8")
        html = inject_asistencia_anual(
            plantilla,
            self.resumen_pruebas(),
            "curso_a",
            "estudiante_a",
        )
        esperados = {
            "T1": ("3", "0", "3", "19"),
            "T2": ("2", "1", "1", "20"),
            "T3": ("1", "0", "1", "21"),
            "ANUAL": ("6", "1", "5", "60"),
        }
        for periodo, valores in esperados.items():
            with self.subTest(periodo=periodo):
                for campo, valor in zip(
                    ("registro", "justificacion", "injustificado", "total"),
                    valores,
                ):
                    self.assertEqual(
                        self.extraer_valor(html, periodo, campo),
                        valor,
                    )

    def test_no_inyecta_resumen_de_otro_curso_o_estudiante(self):
        plantilla = (self.plantillas / "FORMATO DE 3 DE BGU.html").read_text(encoding="utf-8")
        resumen = self.resumen_pruebas()
        for curso, estudiante in (
            ("curso_b", "estudiante_a"),
            ("curso_a", "estudiante_b"),
        ):
            with self.subTest(curso=curso, estudiante=estudiante):
                html = inject_asistencia_anual(
                    plantilla,
                    resumen,
                    curso,
                    estudiante,
                )
                for periodo in ("T1", "T2", "T3", "ANUAL"):
                    for campo in (
                        "registro",
                        "justificacion",
                        "injustificado",
                        "total",
                    ):
                        self.assertEqual(
                            self.extraer_valor(html, periodo, campo),
                            "",
                        )

    def test_trimestre_sin_configuracion_queda_vacio_y_configurado_sin_faltas_muestra_ceros(self):
        plantilla = (self.plantillas / "FORMATO DE 3 DE BGU.html").read_text(encoding="utf-8")
        resumen = self.resumen_pruebas()
        resumen["T2"] = {
            "configurado": True,
            "totalFaltas": 0,
            "justificadas": 0,
            "injustificadas": 0,
            "totalAsistencia": 22,
        }
        resumen["T3"] = {
            "configurado": False,
            "totalFaltas": 0,
            "justificadas": 0,
            "injustificadas": 0,
            "totalAsistencia": 0,
        }
        html = inject_asistencia_anual(plantilla, resumen)
        self.assertEqual(
            tuple(
                self.extraer_valor(html, "T2", campo)
                for campo in ("registro", "justificacion", "injustificado", "total")
            ),
            ("0", "0", "0", "22"),
        )
        for campo in ("registro", "justificacion", "injustificado", "total"):
            self.assertEqual(self.extraer_valor(html, "T3", campo), "")
        self.assertNotRegex(html, r">\s*(None|NaN|null|undefined|200|199)\s*</td>")

    def test_sin_configuracion_mantiene_todas_las_celdas_vacias(self):
        plantilla = (self.plantillas / "FORMATO DE 3 DE BGU.html").read_text(encoding="utf-8")
        html = inject_asistencia_anual(plantilla, None)
        for periodo in ("T1", "T2", "T3", "ANUAL"):
            for campo in ("registro", "justificacion", "injustificado", "total"):
                self.assertEqual(self.extraer_valor(html, periodo, campo), "")

    def test_generador_de_certificados_propaga_asistencia(self):
        salida = self.directorio / "certificados"
        resumen = self.resumen_pruebas()
        resumen["estudianteId"] = "estudiante_1"
        resultado = generar_certificados_inicial({
            "datos_consolidados": [{
                "id_real": "estudiante_1",
                "cedula": "0123456789",
                "nombre": "ESTUDIANTE PRUEBA",
                "materias": {},
                "evaluacion_comportamental": {},
                "asistencia": resumen,
            }],
            "institucion": {"grado": "3RO BGU", "nivel": "BGU"},
            "logos": {},
            "gradoCursoCanonico": "3RO BGU",
            "plantillaName": "FORMATO DE 3 DE BGU.html",
            "certOutputDir": str(salida),
        })
        html = Path(resultado["archivos"][0]).read_text(encoding="utf-8")
        self.assertEqual(self.extraer_valor(html, "T1", "registro"), "3")
        self.assertEqual(self.extraer_valor(html, "T2", "justificacion"), "1")
        self.assertEqual(self.extraer_valor(html, "T3", "total"), "21")
        self.assertEqual(self.extraer_valor(html, "ANUAL", "registro"), "6")
        self.assertEqual(self.extraer_valor(html, "ANUAL", "total"), "60")

    def test_boletin_pdf_se_genera_con_asistencia(self):
        salida = self.directorio / "boletin_asistencia.pdf"
        generar_boletin_pdf(
            [{
                "nombre": "ESTUDIANTE PRUEBA",
                "cedula": "0123456789",
                "promedio": 9,
                "estado": "APROBADO",
                "materias": {},
                "evaluacion_comportamental": {},
                "asistencia": {
                    "configurada": True,
                    "justificadas": 3,
                    "injustificadas": 4,
                    "totalFaltas": 7,
                    "diasLectivos": 190,
                    "totalAsistencia": 183,
                },
            }],
            {
                "nombre": "UEEH",
                "anio": "2026-2027",
                "amie": "01H00000",
                "grado": "3RO BGU",
                "paralelo": "A",
                "jornada": "MATUTINA",
                "tutor": "TUTOR",
                "rector": "RECTOR",
            },
            {},
            str(salida),
        )
        self.assertTrue(salida.exists())
        self.assertGreater(salida.stat().st_size, 1000)
        self.assertEqual(salida.read_bytes()[:4], b"%PDF")


if __name__ == "__main__":
    unittest.main()
