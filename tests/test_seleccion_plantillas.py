import unittest

from procesador_notas import mapear_plantilla_python


class SeleccionPlantillasTests(unittest.TestCase):
    def test_tercero_egb_y_equivalencias_usan_elemental(self):
        for grado in ("3RO DE EGB", "TERCERO DE EGB", "TERCER EGB"):
            with self.subTest(grado=grado):
                self.assertEqual(mapear_plantilla_python(grado), "FORMALO DE ELEMENTAL.html")

    def test_noveno_egb_y_equivalencias_usan_egb_superior(self):
        for grado in ("9NO DE EGB", "NOVENO DE EGB", "9NO EGB"):
            with self.subTest(grado=grado):
                self.assertEqual(mapear_plantilla_python(grado), "FORMATO EGBS.html")

    def test_primero_bachillerato_y_equivalencias_usan_plantilla_bgu(self):
        for grado in (
            "1RO DE BACHILLERATO CIENCIAS",
            "PRIMERO DE BACHILLERATO",
            "1RO DE BGU",
            "PRIMER BGU",
        ):
            with self.subTest(grado=grado):
                self.assertEqual(mapear_plantilla_python(grado), "FORMATO DE 1 Y 2 DE BGU.html")

    def test_bachillerato_no_reconocido_no_recibe_plantilla_arbitraria(self):
        for grado in ("4TO DE BGU", "BACHILLERATO DESCONOCIDO", "CURSO DESCONOCIDO"):
            with self.subTest(grado=grado):
                self.assertIsNone(mapear_plantilla_python(grado))


if __name__ == "__main__":
    unittest.main()
