# -*- coding: utf-8 -*-

import re
import unittest
from pathlib import Path

from catalogo_asignaturas import (
    convertir_nota_a_escala_cualitativa,
    grado_usa_escala_cualitativa,
    normalizar_grado,
)
from procesador_notas import inject_subject_grades


class ConversionEgbElementalTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.plantilla = (
            Path(__file__).parents[1]
            / "assets"
            / "certificados"
            / "FORMALO DE ELEMENTAL.html"
        ).read_text(encoding="utf-8")

    def valores_materia(self, html, nombre):
        fila = re.search(
            rf'<tr[^>]*data-subject="{re.escape(nombre)}"[^>]*>[\s\S]*?</tr>',
            html,
            re.IGNORECASE,
        ).group(0)
        return {
            campo: re.search(
                rf'<td[^>]*data-academic-field="{campo}"[^>]*>([\s\S]*?)</td>',
                fila,
                re.IGNORECASE,
            ).group(1).strip()
            for campo in ("t1", "t2", "t3", "final")
        }

    def test_reconoce_solo_primero_a_cuarto_egb(self):
        for grado in (
            "1RO DE EGB",
            "PRIMERO DE EDUCACIÓN GENERAL BÁSICA",
            "PREPARATORIA",
            "2DO EGB",
            "SEGUNDO DE EGB",
            "3RO DE EGB",
            "TERCER EGB",
            "4TO DE EGB",
            "CUARTO DE EDUCACIÓN GENERAL BÁSICA",
        ):
            with self.subTest(grado=grado):
                self.assertTrue(grado_usa_escala_cualitativa(grado))
        for grado in ("5TO DE EGB", "EGB ELEMENTAL", "1RO BGU", "3RO BGU"):
            with self.subTest(grado=grado):
                self.assertFalse(grado_usa_escala_cualitativa(grado))
        self.assertEqual(normalizar_grado("1ERO DE EGB"), "EGB_1")
        self.assertEqual(normalizar_grado("TERCER EGB"), "EGB_ELEMENTAL")

    def test_conversion_central_respeta_840_950_vacios_y_literales(self):
        self.assertEqual(convertir_nota_a_escala_cualitativa(8.40), "B+")
        self.assertEqual(convertir_nota_a_escala_cualitativa(9.50), "A+")
        self.assertEqual(convertir_nota_a_escala_cualitativa(None), "")
        self.assertEqual(convertir_nota_a_escala_cualitativa(""), "")
        self.assertEqual(convertir_nota_a_escala_cualitativa("B+"), "B+")
        self.assertEqual(convertir_nota_a_escala_cualitativa(0.50), "0.5")
        self.assertEqual(convertir_nota_a_escala_cualitativa(10.50), "10.5")

    def test_segundo_egb_convierte_solo_la_presentacion(self):
        materia = {
            "tipo": "cuantitativa",
            "t1": 8.40,
            "t2": 9.50,
            "t3": None,
            "promedio_anual": 8.95,
        }
        html = inject_subject_grades(
            self.plantilla,
            {"MATEMÁTICA": materia},
            "2DO DE EGB",
        )
        self.assertEqual(
            self.valores_materia(html, "MATEMÁTICA"),
            {"t1": "B+", "t2": "A+", "t3": "", "final": "A-"},
        )
        self.assertEqual(materia["t1"], 8.40)
        self.assertEqual(materia["t2"], 9.50)
        self.assertIsNone(materia["t3"])

    def test_cuarto_egb_convierte_numeros_y_conserva_literales(self):
        html = inject_subject_grades(
            self.plantilla,
            {
                "MATEMÁTICA": {
                    "tipo": "cuantitativa",
                    "t1": 8.40,
                    "t2": "B+",
                    "t3": None,
                    "promedio_anual": 8.40,
                }
            },
            "4TO DE EGB",
        )
        self.assertEqual(
            self.valores_materia(html, "MATEMÁTICA"),
            {"t1": "B+", "t2": "B+", "t3": "", "final": "B+"},
        )

    def test_quinto_egb_conserva_numeros_y_literales(self):
        html = inject_subject_grades(
            self.plantilla,
            {
                "MATEMÁTICA": {
                    "tipo": "cuantitativa",
                    "t1": 8.40,
                    "t2": "B+",
                    "t3": None,
                    "promedio_anual": 8.40,
                }
            },
            "5TO DE EGB",
        )
        self.assertEqual(
            self.valores_materia(html, "MATEMÁTICA"),
            {"t1": "8.40", "t2": "B+", "t3": "", "final": "8.40"},
        )


if __name__ == "__main__":
    unittest.main()
