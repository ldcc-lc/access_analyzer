import unittest

from PySide6.QtWidgets import QApplication, QDialog

from interfaz import ConfigDialog
from src.sistemas_internos.comparador import comparar
from src.sistemas_internos.lectores import CuentaSistema, Persona, ReglaCargo, SistemaConfig


class ConfigDialogTests(unittest.TestCase):
    """Comprueba que los cambios del diálogo llegan al análisis sin reiniciar la app."""
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_analisis_despues_de_guardar_configuracion_sin_reiniciar(self):
        config = SistemaConfig("Prueba", "prueba.txt", "S", "Prueba", 1, 4, 7)
        for prefijo, sufijo, login in [("", "", "ana"), ("bdp", "_fg", "bdpana_fg"),
                                      ("", "", "ana")]:
            with self.subTest(prefijo=prefijo, sufijo=sufijo):
                dialog = ConfigDialog(config)
                dialog.inputs["campo_roles"].setText("5")
                dialog.inputs["prefijo_login"].setText(prefijo)
                dialog.inputs["sufijo_login"].setText(sufijo)
                dialog._accept()
                self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
                config = dialog.config
                self.assertEqual(config.prefijo_login, prefijo.upper())
                self.assertEqual(config.sufijo_login, sufijo.upper())
                self.assertEqual(config.campo_roles, 5)
                resultado = comparar([Persona("ana", "Analista")], [CuentaSistema(login, "1")],
                                     [ReglaCargo("Analista", "1")],
                                     {"sistema": "prueba.txt", "dominio": "dominio.txt"}, config)
                self.assertEqual(resultado["observaciones"], [])
                self.assertEqual(resultado["detalle"][0]["resultado"], "Conforme")
                dialog.close()
