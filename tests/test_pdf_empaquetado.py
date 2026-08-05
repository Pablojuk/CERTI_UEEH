import io
import re
import tempfile
import unittest
from pathlib import Path

from procesador_notas import (
    escribir_json_stdout,
    generar_certificados_inicial,
    leer_stdin_utf8,
)


class FlujoBinario:
    def __init__(self, contenido=b""):
        self.buffer = io.BytesIO(contenido)


class TestPdfEmpaquetado(unittest.TestCase):
    def test_stdin_se_decodifica_como_utf8_sin_depender_de_windows(self):
        flujo = FlujoBinario(
            '{"nombre":"Carlos Andrés","materia":"FÍSICA"}'.encode("utf-8")
        )
        self.assertEqual(
            leer_stdin_utf8(flujo),
            '{"nombre":"Carlos Andrés","materia":"FÍSICA"}',
        )

    def test_stdout_escribe_json_utf8_sin_mojibake(self):
        flujo = FlujoBinario()
        escribir_json_stdout(
            {"nombre": "Carlos Andrés", "materia": "MATEMÁTICA"},
            flujo,
        )
        self.assertEqual(
            flujo.buffer.getvalue().decode("utf-8"),
            '{"nombre": "Carlos Andrés", "materia": "MATEMÁTICA"}\n',
        )

    def test_segundo_bgu_conserva_ceros_decimales_tildes_y_filas(self):
        materias = {
            "MATEMÁTICA": (0, 8.61, 9.2),
            "FÍSICA": (4.82, 7.35, 8),
            "BIOLOGÍA": (8.1, 8.2, 8.3),
            "QUÍMICA": (7.1, 7.2, 7.3),
            "FILOSOFÍA": (9.1, 9.2, 9.3),
            "EDUCACIÓN PARA LA CIUDADANÍA": (8.4, 8.5, 8.6),
            "HISTORIA": (7.4, 7.5, 7.6),
            "LENGUA Y LITERATURA": (8.7, 8.8, 8.9),
            "EDUCACIÓN CULTURAL Y ARTÍSTICA": (9.4, 9.5, 9.6),
            "INGLÉS": (7.7, 7.8, 7.9),
            "EDUCACIÓN FÍSICA": (8.11, 8.22, 8.33),
            "EMPRENDIMIENTO Y GESTIÓN": (9.11, 9.22, 9.33),
        }
        with tempfile.TemporaryDirectory() as directorio:
            resultado = generar_certificados_inicial({
                "cursoActivoId": "curso_bgu2_a",
                "plantillaName": "FORMATO DE 1 Y 2 DE BGU.html",
                "gradoCursoCanonico": "SEGUNDO DE BACHILLERATO",
                "institucion": {
                    "grado": "SEGUNDO DE BACHILLERATO",
                    "paralelo": "A",
                    "nivel": "BGU",
                },
                "logos": {},
                "certOutputDir": directorio,
                "datos_consolidados": [{
                    "id_real": "estudiante_bgu2",
                    "cedula": "0000000000",
                    "nombre": "Carlos Andrés Mendoza Ruiz",
                    "materias": {
                        nombre: {
                            "t1": notas[0],
                            "t2": notas[1],
                            "t3": notas[2],
                            "tipo": "cuantitativa",
                        }
                        for nombre, notas in materias.items()
                    },
                }],
            })
            html = Path(resultado["archivos"][0]).read_text(encoding="utf-8")

        self.assertIn(
            'data-cert-field="student-name">Carlos Andrés Mendoza Ruiz</span>',
            html,
        )
        self.assertIn(
            '>EDUCACIÓN CULTURAL Y ARTÍSTICA </td>',
            html,
        )
        for nombre in materias:
            fila = re.search(
                rf'<tr data-subject="{re.escape(nombre)}">([\s\S]*?)</tr>',
                html,
            )
            self.assertIsNotNone(fila, nombre)
            valores = re.findall(
                r'data-academic-value="true"[^>]*>([^<]*)</td>',
                fila.group(1),
            )
            self.assertEqual(len([valor for valor in valores if valor]), 4, nombre)

        self.assertRegex(
            html,
            r'data-subject="MATEMÁTICA">[\s\S]*?>0\.00</td>[\s\S]*?>8\.61</td>',
        )
        self.assertRegex(
            html,
            r'data-subject="FÍSICA">[\s\S]*?>4\.82</td>',
        )


if __name__ == "__main__":
    unittest.main()
