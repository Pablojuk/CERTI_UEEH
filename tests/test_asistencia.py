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

    def extraer_valor(self, html, clave):
        patron = re.compile(
            rf'<td[^>]*data-asistencia="{clave}"[^>]*>([\s\S]*?)</td>',
            re.IGNORECASE,
        )
        return re.sub(r"<[^>]+>", "", patron.search(html).group(1)).strip()

    def test_todas_las_plantillas_de_asistencia_tienen_identificadores_y_sin_ejemplos(self):
        encontradas = 0
        for ruta in self.plantillas.glob("*.html"):
            html = ruta.read_text(encoding="utf-8")
            if "Asistencia Anual" not in html:
                continue
            encontradas += 1
            with self.subTest(plantilla=ruta.name):
                for clave in ("registro", "justificadas", "injustificadas", "total"):
                    self.assertIn(f'data-asistencia="{clave}"', html)
                    self.assertEqual(self.extraer_valor(html, clave), "")
        self.assertEqual(encontradas, 6)

    def test_inyecta_acumulado_anual_en_las_celdas_correctas(self):
        plantilla = (self.plantillas / "FORMATO DE 3 DE BGU.html").read_text(encoding="utf-8")
        html = inject_asistencia_anual(plantilla, {
            "configurada": True,
            "justificadas": 3,
            "injustificadas": 4,
            "totalFaltas": 7,
            "diasLectivos": 190,
            "totalAsistencia": 183,
        })
        self.assertEqual(self.extraer_valor(html, "registro"), "7")
        self.assertEqual(self.extraer_valor(html, "justificadas"), "3")
        self.assertEqual(self.extraer_valor(html, "injustificadas"), "4")
        self.assertEqual(self.extraer_valor(html, "total"), "183")

    def test_sin_configuracion_mantiene_celdas_vacias(self):
        plantilla = (self.plantillas / "FORMATO DE 3 DE BGU.html").read_text(encoding="utf-8")
        html = inject_asistencia_anual(plantilla, None)
        for clave in ("registro", "justificadas", "injustificadas", "total"):
            self.assertEqual(self.extraer_valor(html, clave), "")
        self.assertNotRegex(html, r">\s*(None|NaN|null|undefined|200|199)\s*</td>")

    def test_generador_de_certificados_propaga_asistencia(self):
        salida = self.directorio / "certificados"
        resultado = generar_certificados_inicial({
            "datos_consolidados": [{
                "id_real": "estudiante_1",
                "cedula": "0123456789",
                "nombre": "ESTUDIANTE PRUEBA",
                "materias": {},
                "evaluacion_comportamental": {},
                "asistencia": {
                    "configurada": True,
                    "justificadas": 1,
                    "injustificadas": 2,
                    "totalFaltas": 3,
                    "diasLectivos": 60,
                    "totalAsistencia": 57,
                },
            }],
            "institucion": {"grado": "3RO BGU", "nivel": "BGU"},
            "logos": {},
            "gradoCursoCanonico": "3RO BGU",
            "plantillaName": "FORMATO DE 3 DE BGU.html",
            "certOutputDir": str(salida),
        })
        html = Path(resultado["archivos"][0]).read_text(encoding="utf-8")
        self.assertEqual(self.extraer_valor(html, "registro"), "3")
        self.assertEqual(self.extraer_valor(html, "justificadas"), "1")
        self.assertEqual(self.extraer_valor(html, "injustificadas"), "2")
        self.assertEqual(self.extraer_valor(html, "total"), "57")

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
