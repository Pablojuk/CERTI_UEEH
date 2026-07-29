import tempfile
import unittest
from pathlib import Path

from procesador_notas import (
    es_segmento_ruta_seguro,
    inject_student_data_html,
    resolver_ruta_hija,
    validar_archivo_excel,
)


class SeguridadProcesadorTests(unittest.TestCase):
    def test_segmentos_y_rutas_bloquean_traversal(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal)
            self.assertEqual(
                resolver_ruta_hija(base, "certificado_123.html"),
                (base / "certificado_123.html").resolve(),
            )
            self.assertIsNone(resolver_ruta_hija(base, "..", "fuera.html"))
        self.assertTrue(es_segmento_ruta_seguro("cedula_0102030405"))
        self.assertFalse(es_segmento_ruta_seguro("../../estudiante"))
        self.assertFalse(es_segmento_ruta_seguro("estudiante\\otro"))
        self.assertFalse(es_segmento_ruta_seguro("estudiante\0otro"))

    def test_datos_personales_se_escapan_en_html(self):
        plantilla = """
        <h1 data-cert-field="institution-name"></h1>
        <span data-cert-field="student-name"></span>
        <span data-cert-field="student-id"></span>
        <span data-cert-field="grade"></span>
        <span data-cert-field="schedule"></span>
        <p data-cert-field="tutor-name"></p>
        <p data-cert-field="rector-name"></p>
        """
        resultado = inject_student_data_html(
            plantilla,
            {
                "nombre": '<img src=x onerror="alert(1)">',
                "cedula": "0102<script>",
            },
            {
                "nombre": "<b>Institución</b>",
                "amie": "AMIE",
                "anio": "2026",
                "grado": "2DO EGB",
                "paralelo": "A",
                "jornada": "MATUTINA",
                "tutor": "<svg onload=alert(1)>",
                "rector": "<script>alert(1)</script>",
            },
            {},
        )
        self.assertNotIn("<img src=x", resultado)
        self.assertNotIn("<script>alert(1)</script>", resultado)
        self.assertIn("&lt;img src=x", resultado)
        self.assertIn("&lt;script&gt;alert(1)&lt;/script&gt;", resultado)

    def test_excel_debe_ser_absoluto_regular_y_con_extension_permitida(self):
        with tempfile.TemporaryDirectory() as temporal:
            base = Path(temporal)
            excel = base / "notas.xlsx"
            texto = base / "notas.txt"
            excel.write_bytes(b"xlsx")
            texto.write_bytes(b"texto")
            self.assertTrue(validar_archivo_excel(excel))
            self.assertFalse(validar_archivo_excel(texto))
            self.assertFalse(validar_archivo_excel(Path("notas.xlsx")))
            self.assertFalse(validar_archivo_excel(base / "ausente.xlsx"))


if __name__ == "__main__":
    unittest.main()
