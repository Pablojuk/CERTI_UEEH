# -*- coding: utf-8 -*-

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from catalogo_asignaturas import (
    clasificar_asignatura,
    grados_equivalentes,
    normalizar_texto_asignatura,
)
from procesador_notas import (
    ErrorGradoExcel,
    cargar_excel_datos,
    consolidar_estudiantes,
    crear_diagnostico_asignaturas,
)


class ReconocimientoAsignaturasTests(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.directorio = Path(self.temporal.name)

    def tearDown(self):
        self.temporal.cleanup()

    def crear_excel(self, nombre, grado, materias, cedula="0123456789", nombre_estudiante="ESTUDIANTE PRUEBA"):
        ruta = self.directorio / nombre
        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte Periodo"
        ws["A1"] = "UNIDAD EDUCATIVA PRUEBA"
        ws["D1"] = "2026-2027"
        ws["B3"] = grado
        ws["D3"] = "RECTOR PRUEBA"
        ws["B4"] = "MATUTINA"
        ws["D4"] = "TUTOR PRUEBA"
        ws["B5"] = "A"
        ws["B6"] = "TRIMESTRE 1"
        ws["B7"] = "01H00000"
        ws.cell(row=9, column=1, value="LISTADO")
        ws.cell(row=9, column=2, value="CEDULA")
        ws.cell(row=10, column=1, value=nombre_estudiante)
        ws.cell(row=10, column=2, value=cedula)
        for indice, (materia, nota) in enumerate(materias, start=3):
            ws.cell(row=9, column=indice, value=materia)
            ws.cell(row=10, column=indice, value=nota)
        wb.save(ruta)
        return ruta

    def test_normaliza_tildes_espacios_mayusculas_y_separadores(self):
        self.assertEqual(normalizar_texto_asignatura("  Biología --  "), "BIOLOGIA")
        self.assertEqual(normalizar_texto_asignatura("educación / física"), "EDUCACION FISICA")

    def test_bologia_se_reconoce_como_biologia(self):
        resultado = clasificar_asignatura("Bologia", "1RO BGU")
        self.assertEqual(resultado["canonica"], "BIOLOGÍA")
        self.assertEqual(resultado["metodo"], "coincidencia_aproximada")

    def test_biolgia_se_reconoce_como_biologia(self):
        self.assertEqual(clasificar_asignatura("Biolgia", "1RO BGU")["canonica"], "BIOLOGÍA")

    def test_matemtica_se_reconoce_como_matematica(self):
        self.assertEqual(clasificar_asignatura("Matemtica", "1RO BGU")["canonica"], "MATEMÁTICA")

    def test_matematica_superior_se_ignora_en_primero_bgu(self):
        resultado = clasificar_asignatura("MATEMÁTICA SUPERIOR", "1RO BGU")
        self.assertEqual(resultado["estado"], "ignorada_por_curso")
        self.assertEqual(resultado["canonica"], "MATEMÁTICA SUPERIOR")

    def test_matematica_superior_se_admite_en_tercero_bgu(self):
        resultado = clasificar_asignatura("MATEMÁTICA SUPERIOR", "3RO BGU")
        self.assertEqual(resultado["estado"], "reconocida")
        self.assertEqual(resultado["canonica"], "MATEMÁTICA SUPERIOR")

    def test_matematica_se_admite_en_primero_bgu(self):
        resultado = clasificar_asignatura("MATEMÁTICA", "1RO BGU")
        self.assertEqual(resultado["estado"], "reconocida")

    def test_matematica_superior_nunca_se_convierte_en_matematica(self):
        for texto in ("MATEMÁTICA SUPERIOR", "MATEMATICA SUPRIOR"):
            for grado in ("1RO BGU", "3RO BGU"):
                self.assertEqual(clasificar_asignatura(texto, grado)["canonica"], "MATEMÁTICA SUPERIOR")

    def test_quimica_superior_nunca_se_convierte_en_quimica(self):
        for texto in ("QUÍMICA SUPERIOR", "QUIMICA SUPERIO"):
            resultado = clasificar_asignatura(texto, "1RO BGU")
            self.assertEqual(resultado["estado"], "ignorada_por_curso")
            self.assertEqual(resultado["canonica"], "QUÍMICA SUPERIOR")

    def test_variantes_de_tres_trimestres_se_consolidan(self):
        t1 = self.crear_excel("t1.xlsx", "1RO BGU", [("BIOLOGÍA", 8.25)])
        t2 = self.crear_excel("t2.xlsx", "PRIMERO DE BACHILLERATO", [("Bologia", 8.5)])
        t3 = self.crear_excel("t3.xlsx", "PRIMER BGU", [("Biologia", 8.75)])
        resultado = consolidar_estudiantes(t1, t2, t3, None, "1RO DE BACHILLERATO CIENCIAS")
        self.assertEqual(list(resultado[0]["materias"]), ["BIOLOGÍA"])
        self.assertEqual(resultado[0]["materias"]["BIOLOGÍA"]["t1"], 8.25)
        self.assertEqual(resultado[0]["materias"]["BIOLOGÍA"]["t2"], 8.5)
        self.assertEqual(resultado[0]["materias"]["BIOLOGÍA"]["t3"], 8.75)

    def test_encabezado_desconocido_no_se_reconoce(self):
        self.assertEqual(clasificar_asignatura("TALLER DE ROBÓTICA", "1RO BGU")["estado"], "no_reconocida")

    def test_grado_distinto_rechaza_sin_alterar_datos_previos(self):
        ruta = self.crear_excel("tercero.xlsx", "3RO BGU", [("MATEMÁTICA SUPERIOR", 9)])
        datos_previos = {"curso": {"promedio": 8.5}}
        copia = {"curso": dict(datos_previos["curso"])}
        with self.assertRaisesRegex(ErrorGradoExcel, "El archivo corresponde a 3RO BGU"):
            cargar_excel_datos(ruta, "1RO BGU")
        self.assertEqual(datos_previos, copia)

    def test_materia_ignorada_no_participa_en_promedio(self):
        rutas = [
            self.crear_excel(f"p{n}.xlsx", "1RO BGU", [("MATEMÁTICA", 9), ("MATEMÁTICA SUPERIOR", 1)])
            for n in range(1, 4)
        ]
        resultado = consolidar_estudiantes(*rutas, None, "1RO BGU")[0]
        self.assertEqual(list(resultado["materias"]), ["MATEMÁTICA"])
        self.assertEqual(resultado["promedio"], 9)
        self.assertEqual(resultado["estado"], "APROBADO")

    def test_cedula_conserva_ceros_iniciales_segun_formato(self):
        ruta = self.crear_excel("cedula.xlsx", "1RO BGU", [("MATEMÁTICA", 9)], cedula=123456789)
        from openpyxl import load_workbook
        libro = load_workbook(ruta)
        libro["Reporte Periodo"]["B10"].number_format = "0000000000"
        libro.save(ruta)
        datos = cargar_excel_datos(ruta, "1RO BGU")
        self.assertIn("0123456789", datos)

    def test_nombres_exactos_existentes_siguen_funcionando(self):
        materias = ["MATEMÁTICA", "FÍSICA", "BIOLOGÍA", "QUÍMICA", "HISTORIA", "LENGUA Y LITERATURA", "INGLÉS", "EDUCACIÓN FÍSICA", "EMPRENDIMIENTO Y GESTIÓN"]
        for materia in materias:
            with self.subTest(materia=materia):
                self.assertEqual(clasificar_asignatura(materia, "3RO BGU")["estado"], "reconocida")

    def test_materias_especiales_de_tercero_no_se_extienden_a_primero(self):
        especiales = [
            "MATEMÁTICA SUPERIOR", "REDACCIÓN CREATIVA", "QUÍMICA SUPERIOR",
            "INVESTIGACIÓN EN CIENCIA Y TECNOLOGÍA", "SOCIOLOGÍA",
        ]
        for materia in especiales:
            with self.subTest(materia=materia):
                self.assertEqual(clasificar_asignatura(materia, "3RO BGU")["estado"], "reconocida")
                self.assertEqual(clasificar_asignatura(materia, "1RO BGU")["estado"], "ignorada_por_curso")

    def test_diagnostico_separa_reconocidas_ignoradas_y_desconocidas(self):
        ruta = self.crear_excel("diagnostico.xlsx", "1RO BGU", [("Bologia", 8), ("MATEMÁTICA SUPERIOR", 9), ("TALLER X", 7)])
        diagnostico = crear_diagnostico_asignaturas()
        datos = cargar_excel_datos(ruta, "1RO BGU", diagnostico)
        self.assertEqual(list(next(iter(datos.values()))["notas"]), ["BIOLOGÍA"])
        self.assertEqual(diagnostico["asignaturasReconocidas"][0]["canonica"], "BIOLOGÍA")
        self.assertEqual(diagnostico["asignaturasIgnoradasPorCurso"][0]["canonica"], "MATEMÁTICA SUPERIOR")
        self.assertEqual(diagnostico["asignaturasNoReconocidas"][0]["original"], "TALLER X")

    def test_equivalencias_de_grado(self):
        self.assertTrue(grados_equivalentes("1RO DE BACHILLERATO CIENCIAS", "PRIMER BGU"))
        self.assertFalse(grados_equivalentes("1RO BGU", "3RO BGU"))

    def test_cli_conserva_contrato_y_entrega_diagnostico(self):
        ruta = self.crear_excel("cli.xlsx", "PRIMER BGU", [("Bologia", 8), ("MATEMÁTICA SUPERIOR", 9)])
        proceso = subprocess.run(
            [sys.executable, str(Path(__file__).parents[1] / "procesador_notas.py"), "--analizar", "--grado", "1RO DE BACHILLERATO CIENCIAS", "--t1", str(ruta)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=True,
        )
        respuesta = json.loads(proceso.stdout)
        self.assertIn("datosInstitucion", respuesta)
        self.assertIn("estudiantes", respuesta)
        self.assertEqual(respuesta["asignaturas"], ["BIOLOGÍA"])
        self.assertEqual(respuesta["asignaturasIgnoradasPorCurso"][0]["canonica"], "MATEMÁTICA SUPERIOR")

    def test_cli_rechaza_grado_distinto_con_mensaje_comprensible(self):
        ruta = self.crear_excel("cli_otro_grado.xlsx", "3RO BGU", [("MATEMÁTICA SUPERIOR", 9)])
        proceso = subprocess.run(
            [sys.executable, str(Path(__file__).parents[1] / "procesador_notas.py"), "--analizar", "--grado", "1RO BGU", "--t1", str(ruta)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=True,
        )
        respuesta = json.loads(proceso.stdout)
        self.assertEqual(respuesta["error"], "El archivo corresponde a 3RO BGU, pero el curso seleccionado es 1RO BGU.")


if __name__ == "__main__":
    unittest.main()
