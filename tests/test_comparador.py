import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from src.sistemas_internos.comparador import comparar
from src.sistemas_internos.proceso import ejecutar_revision
from src.sistemas_internos.reporte import generar_reporte
from src.sistemas_internos.lectores import (CuentaSistema, Persona, ReglaCargo, SistemaConfig,
                          leer_accesos_temporales, leer_personas_txt, leer_matrices,
                          leer_salar, leer_sistema_txt, normalizar_cargo, normalizar_login,
                          normalizar_login_sistema)


# Datos ficticios compartidos; las pruebas pasan la configuración de sistema explícitamente.
ARCHIVOS = {"matriz": "matriz_ficticia.xlsx", "dominio": "dominio_ficticio.txt",
            "argosnet": "argosnet_ficticio.txt"}
ARGOSNET_CONFIG = SistemaConfig("ArgosNet", "ArgosNet.txt", "S", "Argos Net", 1, 4, 7,
                                null_sin_roles=True)


class ComparadorTests(unittest.TestCase):
    """Reglas de cruce, matriz, roles, casos externos y accesos temporales."""
    def test_argosnet_compara_roles_de_matriz_y_oficina_con_salar(self):
        config = SistemaConfig("ArgosNet", "ArgosNet.txt", "S", "Argos Net", 1, 4, 7,
            null_sin_roles=True, campo_oficina=6)
        persona = Persona("ana", "Analista", "ANA PERSONA", oficina="Agencia Norte")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "argosnet.txt"
            path.write_text("ana|x|x|1|x|Agencia Sur|x\n", encoding="utf-8")
            cuentas, errores = leer_sistema_txt(path, config)
        salida = comparar([persona], cuentas, [ReglaCargo("Analista", "1")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, config)
        self.assertEqual(errores, [])
        self.assertEqual(cuentas[0].oficina, "Agencia Sur")
        self.assertEqual([o["categoria"] for o in salida["observaciones"]],
                         ["Diferencia de Oficina/Gerencia"])
        self.assertEqual(salida["detalle"][0]["resultado"], "Conforme")

    def test_srg_nc_cruza_por_correo_y_acepta_cualquier_estado_no_ina(self):
        config = SistemaConfig("SRGNC", "SRGNC.txt", "AO", "SRGNC", 3, 5, 6,
            campo_estado=4, estado_inactivo="INA", clave_cruce="correo")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "SRGNC.txt"
            path.write_text(
                "Ana|BDPANA|ana@example.com|BLO|2|Operador\n"
                "Ana|BDPANA|otro@example.com|INA|2|Operador\n",
                encoding="utf-8")
            cuentas, errores = leer_sistema_txt(path, config)
        persona = Persona("ana", "Analista", correo="ANA@EXAMPLE.COM")
        sin_correo = Persona("sin_correo", "Analista", fila=3)
        salida = comparar([persona, sin_correo], cuentas, [ReglaCargo("Analista", "2")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, config)
        self.assertEqual([cuenta.estado for cuenta in cuentas], ["BLO"])
        self.assertEqual(errores, [])
        self.assertEqual(salida["detalle"][0]["resultado"], "Conforme")
        self.assertEqual(salida["observaciones"], [])
        self.assertTrue(any(error["campo"] == "correo" for error in salida["errores"]))

    def test_punto_reclamo_valida_existencia_y_estado_sin_comparar_roles(self):
        config = SistemaConfig("PUNTO RECLAMO", "PUNTO RECLAMO.txt", "AD",
            "Punto de Reclamo (PR) BDP", 3, 6, 6, campo_estado=6,
            estado_activo="Activo", campo_oficina=4, validar_solo_existencia=True)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "PUNTO RECLAMO.txt"
            path.write_text(
                "x|x|ana|Agencia Norte|x|Inactivo\n"
                "x|x|luis|Agencia Sur|x|Activo\n"
                "x|x|externo|Agencia Este|x|Inactivo\n",
                encoding="utf-8")
            cuentas, errores = leer_sistema_txt(path, config)
        personas = [Persona("ana", "Analista", oficina="Agencia Norte"),
                    Persona("luis", "Analista", oficina="Agencia Sur"),
                    Persona("maria", "Analista")]
        salida = comparar(personas, cuentas, [ReglaCargo("Analista", "101")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, config)
        self.assertEqual([cuenta.estado for cuenta in cuentas], ["Inactivo", "Activo", "Inactivo"])
        self.assertEqual(errores, [])
        self.assertEqual([o["categoria"] for o in salida["observaciones"]],
            ["Estado distinto de Activo", "Usuarios en Salar sin registro en sistema"])
        self.assertEqual([d["resultado"] for d in salida["detalle"]],
            ["Estado distinto de Activo", "Conforme",
             "Cuenta inactiva fuera de la fuente de personas", "Sin registro en sistema"])
        self.assertEqual(salida["observaciones"][0]["esperado"], "Con Acceso")
        self.assertEqual(salida["observaciones"][1]["esperado"], "Con Acceso")

        sin_acceso = comparar([Persona("ana", "Auxiliar")],
            [CuentaSistema("ana", "", estado="Activo")],
            [ReglaCargo("Auxiliar", " - ")], ARCHIVOS, config)
        self.assertEqual(sin_acceso["observaciones"][0]["categoria"],
                         "Registro en sistema sin acceso esperado")
        self.assertEqual(sin_acceso["observaciones"][0]["esperado"], "Sin Acceso")

    def test_normalizacion_admite_ajustes_de_login_nulos(self):
        config = SistemaConfig("Prueba", "prueba.txt", "S", "Prueba", 1, 4, 7,
                               prefijo_login=None, sufijo_login=None)
        self.assertEqual(normalizar_login_sistema(" ana ", config), "ANA")

    def test_normalizacion_no_cambia_el_valor_original(self):
        self.assertEqual(normalizar_login("  jLopez  "), "JLOPEZ")
        self.assertEqual(normalizar_cargo("  Analista   de Áreas "), "analista de areas")
        self.assertEqual(
            normalizar_cargo("ANALISTA DE INFORMACION Y DESARROLLO EN APLICACIONES NTIC'S"),
            normalizar_cargo("ANALISTA DE INFORMACION Y DESARROLLO EN APLICACIONES NTICS"))

    def test_normalizadores_especificos_solo_quitan_prefijo_y_sufijo_configurados(self):
        bdp = SistemaConfig("BDPSAM", "BDPSAM.txt", "W", "Argos Formularios 2P", 1, 5, 6,
                            prefijo_login="bdp")
        fg = SistemaConfig("FG", "FG.txt", "X", "Fondo de Garantias", 1, 5, 6,
                           prefijo_login="bdp", sufijo_login="_fg")
        self.assertEqual(normalizar_login_sistema(" BDPusuario ", bdp), "USUARIO")
        self.assertEqual(normalizar_login_sistema(" BDPaborja_FG ", fg), "ABORJA")

    def test_diferencias_de_roles_se_consolidan_y_el_orden_no_importa(self):
        salida = comparar([Persona("ana", "Analista")], [CuentaSistema(" ANA ", "3, 9")],
            [ReglaCargo("Analista", "1-3")], ARCHIVOS, ARGOSNET_CONFIG)
        cats = [o["categoria"] for o in salida["observaciones"]]
        self.assertEqual(cats, ["Diferencia de roles"])
        self.assertEqual(salida["observaciones"][0]["esperado"], "1, 3")
        self.assertEqual(salida["observaciones"][0]["encontrado"], "3, 9")

    def test_cuenta_sin_dominio_no_se_empareja_por_nombre(self):
        salida = comparar([Persona("otra", "Analista", "Nombre parecido")],
            [CuentaSistema("tecnica", "Null")], [ReglaCargo("Analista", "1")], ARCHIVOS, ARGOSNET_CONFIG)
        self.assertEqual(salida["sin_coincidencia"], 1)
        self.assertEqual(salida["observaciones"][0]["categoria"],
                         "Usuario fuera de Dominio con acceso pendiente de retirar")
        self.assertIn("Fuera de Dominio", salida["detalle"][0]["resultado"])

    def test_dominio_sin_cuenta_solo_se_reporta_si_hay_roles(self):
        personas = [Persona("con_acceso", "Analista"), Persona("sin_acceso", "Consulta")]
        reglas = [ReglaCargo("Analista", "4"), ReglaCargo("Consulta", "-")]
        salida = comparar(personas, [], reglas, ARCHIVOS, ARGOSNET_CONFIG)
        self.assertEqual([o["login"] for o in salida["observaciones"]], ["con_acceso"])
        self.assertEqual(salida["observaciones"][0]["categoria"],
                         "Usuario en Dominio sin registro en sistema")

    def test_dominio_fuera_de_salar_solo_genera_alta_si_no_tiene_cuenta(self):
        archivos = {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}
        dominio = [Persona("externo", "Analista", "Persona Externa")]
        reglas = [ReglaCargo("Analista", "4")]
        sin_cuenta = comparar([], [], reglas, archivos, ARGOSNET_CONFIG,
            personas_dominio=dominio)
        self.assertEqual([o["categoria"] for o in sin_cuenta["observaciones"]],
            ["Usuario en Dominio sin registro en sistema"])
        self.assertEqual(sin_cuenta["observaciones"][0]["esperado"], "4")

        con_cuenta = comparar([], [CuentaSistema("externo", "4")], reglas,
            archivos, ARGOSNET_CONFIG, personas_dominio=dominio)
        self.assertNotIn("Usuario en Dominio sin registro en sistema",
            [o["categoria"] for o in con_cuenta["observaciones"]])

    def test_null_y_guion_son_sin_roles(self):
        salida = comparar([Persona("a", "A"), Persona("b", "B")],
            [CuentaSistema("a", "Null"), CuentaSistema("b", "")],
            [ReglaCargo("A", "-"), ReglaCargo("B", "")], ARCHIVOS, ARGOSNET_CONFIG)
        self.assertEqual(salida["observaciones"], [])
        self.assertEqual({d["resultado"] for d in salida["detalle"]}, {"Conforme"})

    def test_cargo_no_encontrado_y_formato_invalido(self):
        salida = comparar([Persona("a", "Otro"), Persona("b", "Analista")],
            [CuentaSistema("a", "1"), CuentaSistema("b", "1,x")],
            [ReglaCargo("Analista", "1")], ARCHIVOS, ARGOSNET_CONFIG)
        cats = [o["categoria"] for o in salida["observaciones"]]
        self.assertIn("Cargo sin correspondencia en matriz", cats)
        self.assertIn("Duplicados o casos especiales", cats)

    def test_cargo_fuera_de_matriz_se_observa_solo_si_hay_acceso_encontrado(self):
        persona = Persona("a", "Externo")
        regla = [ReglaCargo("Analista", "1")]
        sin_acceso = comparar([persona], [CuentaSistema("a", "Null")], regla, ARCHIVOS, ARGOSNET_CONFIG)
        self.assertEqual(sin_acceso["observaciones"], [])
        con_acceso = comparar([persona], [CuentaSistema("a", "8")], regla, ARCHIVOS, ARGOSNET_CONFIG)
        self.assertEqual(con_acceso["observaciones"][0]["categoria"],
                         "Cargo sin correspondencia en matriz")

    def test_null_no_se_asume_sin_roles_en_otro_sistema(self):
        config = SistemaConfig("ArgosAPX", "ArgosAPX.txt", "R", "ArgosAPX", 2, 5, 20)
        salida = comparar([Persona("a", "Analista")], [CuentaSistema("a", "Null")],
            [ReglaCargo("Analista", "1")], ARCHIVOS, config)
        self.assertEqual(salida["observaciones"][0]["categoria"], "Duplicados o casos especiales")
        self.assertIn("significado", salida["observaciones"][0]["motivo"])

    def test_null_confirmado_se_compara_como_conjunto_vacio(self):
        config = SistemaConfig("ArgosAPX", "ArgosAPX.txt", "R", "ArgosAPX", 2, 5, 20,
                               null_sin_roles=True)
        salida = comparar([Persona("a", "Analista")], [CuentaSistema("a", "Null")],
            [ReglaCargo("Analista", "-")], ARCHIVOS, config)
        self.assertEqual(salida["detalle"][0]["resultado"], "Conforme")

    def test_logins_duplicados_se_marcan_y_no_se_elige_registro(self):
        salida = comparar([Persona("x", "A"), Persona(" X ", "A")],
            [CuentaSistema("x", "1")], [ReglaCargo("A", "1")], ARCHIVOS, ARGOSNET_CONFIG)
        self.assertTrue(any(o["categoria"] == "Duplicados o casos especiales" for o in salida["observaciones"]))
        self.assertTrue(salida["detalle"][0]["resultado"].startswith("No evaluable"))

    def test_cia_compara_rol_y_organigrama_en_una_observacion(self):
        config = SistemaConfig("CIA", "CIA.txt", "AA", "CIA", 3, 1, 6,
            columna_matriz_organigrama="AB", encabezado_matriz_organigrama="ORGANIGRAMA - CIA",
            campo_organigrama=6)
        salida = comparar([Persona("ana", "Analista")], [CuentaSistema("ana", "1", 1, "4,9")],
            [ReglaCargo("Analista", "1", 1, "2-4")], ARCHIVOS, config)
        self.assertEqual(len(salida["observaciones"]), 1)
        self.assertEqual(salida["observaciones"][0]["categoria"], "Diferencia de roles")
        self.assertEqual(salida["observaciones"][0]["esperado"], "Roles: 1; Organigrama: 2, 4")
        self.assertEqual(salida["observaciones"][0]["encontrado"], "Roles: 1; Organigrama: 4, 9")

    def test_cia_ignora_nota_entre_parentesis_en_organigrama_de_matriz(self):
        config = SistemaConfig("CIA", "CIA.txt", "AA", "CIA", 3, 1, 6,
            columna_matriz_organigrama="AB", encabezado_matriz_organigrama="ORGANIGRAMA - CIA",
            campo_organigrama=6)
        regla = ReglaCargo("Jefe Regional Zonal", "201", 1,
            "4855 - 4865 - 4870\n(solo a la Gerencia Regional a la que corresponde)")
        cuenta = CuentaSistema("ana", "201", 1, "4855, 4865, 4870")
        salida = comparar([Persona("ana", "Jefe Regional Zonal")], [cuenta],
            [regla], ARCHIVOS, config)
        self.assertEqual(salida["observaciones"], [])
        self.assertEqual(salida["detalle"][0]["esperados"],
                         "Roles: 201; Organigrama: 4855, 4865, 4870")
        self.assertEqual(salida["detalle"][0]["resultado"], "Conforme")

    def test_argosapx_normaliza_variantes_de_campos_y_reporte_nombre_sistema(self):
        config = SistemaConfig("ArgosAPX", "ArgosAPX.txt", "R", "ArgosAPX", 2, 5, 20,
            null_sin_roles=True, campo_nombre=3, campo_cargo=14,
            campo_oficina=15, campo_gerencia=16)
        persona = Persona("jchoque", "SUB GERENTE DE FIDEICOMISOS", "JHONNY CHOQUE",
                          oficina="AGENCIA SALINAS DE GARCI MENDOZA",
                          gerencia="GERENCIA DE FIDEICOMISOS")
        cuenta = CuentaSistema("jchoque", "2", nombre="JHONNY . CHOQUE DE COLLAO",
            cargo="SUBGERENTE DE FIDEICOMISOS",
            oficina="AGE. SALINAS DE GARCI MENDOZA",
            gerencia="GERENCIA DE FIDEICOMISOS")
        salida = comparar([persona], [cuenta],
            [ReglaCargo("SUBGERENTE DE FIDEICOMISOS", "1")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, config)
        self.assertEqual([o["categoria"] for o in salida["observaciones"]],
                         ["Diferencia de roles"])
        self.assertEqual(salida["observaciones"][0]["nombre"], "JHONNY CHOQUE")
        self.assertEqual(salida["observaciones"][0]["nombre_sistema"],
                         "JHONNY . CHOQUE DE COLLAO")

    def test_acceso_temporal_vencido_se_clasifica_y_conserva_fechas(self):
        temporal = {"login": "ana", "cargo_actual": "Analista",
                    "cargo_designado": "Analista",
                    "oficina_designada": "Oficina Central",
                    "fecha_inicio": "2020-01-01", "fecha_fin": "2020-12-31"}
        salida = comparar([Persona("ana", "Analista")], [CuentaSistema("ana", "1")],
            [ReglaCargo("Analista", "1")], ARCHIVOS,
            accesos_temporales=[temporal], config=ARGOSNET_CONFIG)
        observacion = salida["observaciones"][0]
        self.assertEqual(observacion["categoria"], "Fin Acceso Temporal")
        self.assertEqual((observacion["fecha_inicio"], observacion["fecha_fin"]),
                         ("2020-01-01", "2020-12-31"))

    def test_acceso_temporal_usa_cargo_designado_para_roles_y_compara_oficina(self):
        config = SistemaConfig("ArgosAPX", "ArgosAPX.txt", "R", "ArgosAPX", 2, 5, 20,
            null_sin_roles=True, campo_nombre=3, campo_cargo=14,
            campo_oficina=15, campo_gerencia=16)
        temporal = {"login": "ana", "cargo_actual": "Analista", "cargo_designado": "Jefe",
                    "cargo_designado_pintado": True,
                    "oficina_actual": "Oficina Central", "oficina_designada": "Agencia Norte",
                    "fecha_inicio": "2026-01-01", "fecha_fin": "2099-12-31"}
        cuenta = CuentaSistema("ana", "2", cargo="Jefe", oficina="Agencia Norte")
        salida = comparar([Persona("ana", "Analista")], [cuenta],
            [ReglaCargo("Analista", "1"), ReglaCargo("Jefe", "2")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, config,
            accesos_temporales=[temporal])
        observacion = salida["observaciones"][0]
        self.assertEqual(observacion["categoria"], "Con Acceso Temporal")
        self.assertEqual(observacion["esperado"], "2")
        self.assertEqual(observacion["resultado_roles"], "Conforme")
        self.assertEqual(observacion["resultado_cargo"], "Conforme")
        self.assertEqual(observacion["resultado_oficina"], "Conforme")

    def test_acceso_temporal_con_ambos_cargos_pintados_conserva_y_adiciona_roles(self):
        temporal = {"login": "ana", "cargo_actual": "Analista", "cargo_designado": "Jefe",
                    "cargo_actual_pintado": True, "cargo_designado_pintado": True}
        salida = comparar([Persona("ana", "Analista")],
            [CuentaSistema("ana", "1,2", cargo="Jefe")],
            [ReglaCargo("Analista", "1"), ReglaCargo("Jefe", "2")], ARCHIVOS,
            accesos_temporales=[temporal], config=ARGOSNET_CONFIG)
        observacion = salida["observaciones"][0]
        self.assertEqual(observacion["resultado_roles"], "Conforme")
        self.assertEqual(observacion["esperado"], "1 + 2")

    def test_alta_salar_sin_cuenta_usa_cargo_designado_temporal(self):
        temporal = {"login": "ana", "cargo_actual": "Analista", "cargo_designado": "Jefe",
                    "cargo_designado_pintado": True, "fecha_fin": "2099-12-31"}
        salida = comparar([Persona("ana", "Analista", "ANA PERSONA")], [],
            [ReglaCargo("Analista", "1"), ReglaCargo("Jefe", "2")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, ARGOSNET_CONFIG,
            accesos_temporales=[temporal])
        observacion = salida["observaciones"][0]
        self.assertEqual(observacion["categoria"], "Usuarios en Salar sin registro en sistema")
        self.assertEqual(observacion["cargo"], "Jefe")
        self.assertEqual(observacion["esperado"], "2")

    def test_alta_salar_con_temporal_vencido_usa_cargo_original(self):
        temporal = {"login": "ana", "cargo_actual": "Analista", "cargo_designado": "Jefe",
                    "cargo_actual_pintado": True, "cargo_designado_pintado": True,
                    "fecha_fin": "2020-12-31"}
        salida = comparar([Persona("ana", "Analista", "ANA PERSONA")], [],
            [ReglaCargo("Analista", "1"), ReglaCargo("Jefe", "2")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, ARGOSNET_CONFIG,
            accesos_temporales=[temporal])
        observacion = salida["observaciones"][0]
        self.assertEqual(observacion["cargo"], "Analista")
        self.assertEqual(observacion["esperado"], "1")

    def test_alta_salar_con_ambos_cargos_pintados_suma_roles_temporales(self):
        temporal = {"login": "ana", "cargo_actual": "Analista", "cargo_designado": "Jefe",
                    "cargo_actual_pintado": True, "cargo_designado_pintado": True,
                    "fecha_fin": "2099-12-31"}
        salida = comparar([Persona("ana", "Analista", "ANA PERSONA")], [],
            [ReglaCargo("Analista", "1"), ReglaCargo("Jefe", "2")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, ARGOSNET_CONFIG,
            accesos_temporales=[temporal])
        observacion = salida["observaciones"][0]
        self.assertEqual(observacion["cargo"], "(Actual) Analista\n(Designado) Jefe")
        self.assertEqual(observacion["esperado"], "1, 2")

    def test_alta_dominio_sin_salar_usa_cargo_designado_temporal(self):
        archivos = {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}
        dominio = [Persona("ana", "Analista", "ANA PERSONA")]
        temporal = {"login": "ana", "cargo_actual": "Analista", "cargo_designado": "Jefe",
                    "cargo_designado_pintado": True, "fecha_fin": "2099-12-31"}
        salida = comparar([], [],
            [ReglaCargo("Analista", "1"), ReglaCargo("Jefe", "2")],
            archivos, ARGOSNET_CONFIG, accesos_temporales=[temporal],
            personas_dominio=dominio)
        observacion = salida["observaciones"][0]
        self.assertEqual(observacion["categoria"], "Usuario en Dominio sin registro en sistema")
        self.assertEqual(observacion["cargo"], "Jefe")
        self.assertEqual(observacion["esperado"], "2")

    def test_acceso_temporal_punto_reclamo_muestra_valor_matriz(self):
        config = SistemaConfig("PUNTO RECLAMO", "PUNTO RECLAMO.txt", "AD",
            "Punto de Reclamo (PR) BDP", 3, 6, 6, campo_estado=6,
            estado_activo="Activo", validar_solo_existencia=True)
        temporal = {"login": "ana", "cargo_actual": "Analista",
                    "cargo_designado": "JEFE DE AGENCIA MOVIL"}
        cuenta = CuentaSistema("ana", "", estado="Activo", cargo="JEFE DE AGENCIA MOVIL")
        salida = comparar([Persona("ana", "Analista")], [cuenta],
            [ReglaCargo("Analista", "-"), ReglaCargo("JEFE DE AGENCIA MOVIL", "122")],
            ARCHIVOS, config, accesos_temporales=[temporal])
        observacion = salida["observaciones"][0]
        self.assertEqual(observacion["esperado"], "122")
        self.assertEqual(observacion["encontrado"], "Activo")
        self.assertEqual(observacion["resultado_roles"], "Conforme")

    def test_diferencias_de_oficina_y_gerencia_se_agrupan_y_guardan_datos_de_salar(self):
        config = SistemaConfig("ArgosAPX", "ArgosAPX.txt", "R", "ArgosAPX", 2, 5, 20,
            null_sin_roles=True, campo_nombre=3, campo_cargo=14,
            campo_oficina=15, campo_gerencia=16)
        persona = Persona("ana", "Analista", "ANA PERSONA", oficina="Agencia Norte",
                          gerencia="Gerencia A")
        cuenta = CuentaSistema("ana", "1", nombre="ANA PERSONA", cargo="Analista",
                             oficina="Agencia Sur", gerencia="Gerencia B")
        salida = comparar([persona], [cuenta], [ReglaCargo("Analista", "1")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, config)
        diferencias = [o for o in salida["observaciones"]
                       if o["categoria"] == "Diferencia de Oficina/Gerencia"]
        self.assertEqual(len(diferencias), 1)
        self.assertEqual((diferencias[0]["oficina_sistema"], diferencias[0]["oficina_salar"]),
                         ("Agencia Sur", "Agencia Norte"))
        self.assertEqual((diferencias[0]["gerencia_sistema"], diferencias[0]["gerencia_salar"]),
                         ("Gerencia B", "Gerencia A"))

    def test_reporte_separa_columnas_normales_y_temporales(self):
        config = SistemaConfig("ArgosAPX", "ArgosAPX.txt", "R", "ArgosAPX", 2, 5, 20,
            null_sin_roles=True, campo_nombre=3, campo_cargo=14,
            campo_oficina=15, campo_gerencia=16)
        salida = comparar([Persona("ana", "Analista", "ANA PERSONA")],
            [CuentaSistema("ana", "2", nombre="ANA PERSONA", cargo="Analista")],
            [ReglaCargo("Analista", "1")],
            {**ARCHIVOS, "etiqueta_fuente_personas": "Salar"}, config,
            accesos_temporales=[{"login": "ana", "cargo_actual": "Analista",
                "cargo_designado": "Analista",
                "oficina_designada": "Norte", "fecha_inicio": "2026-01-01",
                "fecha_fin": "2099-12-31"}])
        resultado = {**salida, "sistemas": [], "dominios": [1], "fuente_personas": "Salar"}
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "reporte.xlsx"
            generar_reporte(resultado, {"matriz": "matriz.xlsx", "dominio": "personas.xlsx",
                "fuente_personas": "salar.xlsx"}, path, ("ArgosAPX",))
            report = load_workbook(path, data_only=True, read_only=True)
            try:
                normales = report["Observaciones ArgosAPX"]
                temporales = report["Temporal ArgosAPX"]
                self.assertEqual(normales.max_column, 12)
                self.assertNotIn("Cargo (temporal, designado)",
                                 [c.value for c in normales[1]])
                self.assertEqual(temporales.max_column, 12)
                self.assertIn("Cargo (temporal, designado)",
                              [c.value for c in temporales[1]])
            finally:
                report.close()


class LectoresTests(unittest.TestCase):
    """Lectura y validación de TXT, matrices Excel y accesos temporales."""
    def test_lector_filtra_estado_inactivo_si_se_configura(self):
        config = SistemaConfig("PUNTO RECLAMO", "PUNTO RECLAMO.txt", "AD",
            "Punto de Reclamo (PR) BDP", 3, 6, 6, campo_estado=6,
            estado_activo="Activo", estado_inactivo="Inactivo",
            validar_solo_existencia=True)
        filas = [
            (1, ["x", "x", "ana", "Oficina", "x", "Inactivo"]),
            (2, ["x", "x", "luis", "Oficina", "x", "Activo"]),
        ]
        with patch("src.sistemas_internos.lectores._leer_txt", return_value=filas):
            cuentas, errores = leer_sistema_txt("PUNTO RECLAMO.txt", config)
        self.assertEqual([cuenta.login for cuenta in cuentas], ["luis"])
        self.assertEqual(errores, [])

    def test_lector_salar_carga_correo_columna_19(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "salar.txt"
            row = [""] * 20
            row[0:5] = ["Ana", "Persona", "", "", ""]
            row[13], row[14], row[16], row[18], row[19] = (
                "Gerencia A", "Agencia Norte", "Analista", "ana@example.com", "ana")
            path.write_text("|".join(row) + "\n", encoding="utf-8")
            personas, errores = leer_salar(path)
        self.assertEqual(personas[0].correo, "ana@example.com")
        self.assertEqual(errores, [])

    def test_lector_temporales_lee_fechas_en_f_y_g(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "temporales.xlsx"
            wb = Workbook(); ws = wb.active
            ws.append(["Usuario", "Nombre", "Tipo", "Cargo Actual", "Cargo Asignado",
                       "Oficina Actual", "Oficina Asignada", "Fecha inicio", "Fecha fin"])
            ws.append(["ana", "Ana Persona", "Cambio", "Analista", "Jefe",
                       "Sucursal A", "Sucursal B", "2026-01-02", "2026-12-31"])
            wb.save(path)
            filas = leer_accesos_temporales(path)
        self.assertEqual((filas[0]["fecha_inicio"], filas[0]["fecha_fin"]),
                         ("2026-01-02", "2026-12-31"))
        self.assertEqual((filas[0]["cargo_designado"], filas[0]["oficina_designada"]),
                         ("Jefe", "Sucursal B"))

    def test_srg_nc_ignora_estados_no_activos_y_normaliza_prefijo(self):
        config = SistemaConfig("SRGNC", "SRGNC.txt", "AO", "SRGNC", 2, 5, 6,
                               prefijo_login="BDP", campo_estado=4, estado_activo="ACT")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "SRGNC.txt"
            path.write_text(
                "Ana|BDPANA|ana@example.com|ACT|2|Operador\n"
                "Ana|BDPANA|ana@example.com|INA|1|Administrador\n",
                encoding="utf-8")
            cuentas, errores = leer_sistema_txt(path, config)
        self.assertEqual([(c.login, c.roles_texto) for c in cuentas], [("BDPANA", "2")])
        self.assertEqual(errores, [])
        salida = comparar([Persona("ana", "Analista")], cuentas,
            [ReglaCargo("Analista", "2")], ARCHIVOS, config)
        self.assertEqual(salida["detalle"][0]["resultado"], "Conforme")

    def test_presupuesto_normaliza_codigos_y_agrupa_roles_por_login(self):
        config = SistemaConfig("PRESUPUESTO", "PRESUPUESTO.txt", "BD", "PRESUPUESTO",
                               1, 6, 6, agrupar_roles_por_login=True,
                               normalizar_codigos_rol=True)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "PRESUPUESTO.txt"
            path.write_text(
                "ANA|Ana Persona|TRANSPARENCIA|1|Registrador|3.000,00-3.001,00\n"
                "ANA|Ana Persona|TRANSPARENCIA|2|Solicitante|3.001,00\n"
                "LUIS|Luis Persona|TRANSPARENCIA|1|Registrador|3.000,00\n",
                encoding="utf-8")
            cuentas, errores = leer_sistema_txt(path, config)
        self.assertEqual([(c.login, c.roles_texto) for c in cuentas],
                         [("ANA", "3000,3001"), ("LUIS", "3000")])
        self.assertEqual(errores, [])
        salida = comparar([Persona("ana", "Analista")], [cuentas[0]],
            [ReglaCargo("Analista", "3000,3001")], ARCHIVOS, config)
        self.assertEqual(salida["detalle"][0]["resultado"], "Conforme")

    def test_presupuesto_codigo_rol_invalido_se_conserva_como_no_evaluable(self):
        config = SistemaConfig("PRESUPUESTO", "PRESUPUESTO.txt", "BD", "PRESUPUESTO",
                               1, 6, 6, agrupar_roles_por_login=True,
                               normalizar_codigos_rol=True)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "PRESUPUESTO.txt"
            path.write_text("ANA|Ana Persona|TRANSPARENCIA|1|Registrador|3.000,50\n",
                            encoding="utf-8")
            cuentas, errores = leer_sistema_txt(path, config)
        self.assertEqual(len(cuentas), 1)
        self.assertEqual(len(errores), 1)
        salida = comparar([Persona("ana", "Analista")], cuentas,
            [ReglaCargo("Analista", "3000")], ARCHIVOS, config)
        self.assertEqual(salida["detalle"][0]["resultado"], "No evaluable: formato o significado de roles")

    def test_domino_cp1252_lee_nombre_cargo_login_y_fila(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dominio.txt"
            cols = [""] * 13
            cols[0], cols[3], cols[12] = "José Álvarez", "Analista", " jlopez "
            path.write_bytes(("|".join(cols) + "\r\n").encode("cp1252"))
            personas, errores = leer_personas_txt(path)
        self.assertEqual((personas[0].nombre, personas[0].cargo, personas[0].login),
                         ("José Álvarez", "Analista", " jlopez "))
        self.assertEqual(personas[0].fila, 1)
        self.assertEqual(errores, [])

    def test_argosnet_se_lee_con_configuracion_generica(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "argos.txt"
            path.write_text(" user ||| Null ||| BLOQUEADO\n", encoding="utf-8")
            config = SistemaConfig("ArgosNet", "ArgosNet.txt", "S", "Argos Net",
                                   1, 4, 7, null_sin_roles=True)
            cuentas, errores = leer_sistema_txt(path, config)
        self.assertEqual((cuentas[0].login, cuentas[0].roles_texto), (" user ", " Null "))
        self.assertEqual(errores, [])

    def test_matriz_busca_encabezado_en_sin_fila_fija(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "matriz.xlsx"
            wb = Workbook(); ws = wb.active; ws.title = " Matriz de Accesos"
            ws["B1"] = "Título de matriz"; ws["S4"] = "Argos Net"
            ws["B6"] = "Analista"; ws["S6"] = "1-3"
            ws["B7"] = "Consulta"; ws["S7"] = "-"
            wb.save(path)
            config = SistemaConfig("ArgosNet", "ArgosNet.txt", "S", "Argos Net",
                                   1, 4, 7, null_sin_roles=True)
            reglas = leer_matrices(path, (config,))["ArgosNet"]
        self.assertEqual([(r.cargo, r.roles_texto) for r in reglas],
                         [("Analista", "1-3"), ("Consulta", "-")])

    def test_matriz_cia_lee_rol_y_organigrama_en_columnas_separadas(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "matriz.xlsx"
            wb = Workbook(); ws = wb.active; ws.title = "Matriz de Accesos"
            ws["AA6"] = "CIA"; ws["AB6"] = "ORGANIGRAMA - CIA"
            ws["B7"] = "Analista"; ws["AA7"] = "1"; ws["AB7"] = "2-4"
            wb.save(path)
            config = SistemaConfig("CIA", "CIA.txt", "AA", "CIA", 3, 1, 6,
                columna_matriz_organigrama="AB", encabezado_matriz_organigrama="ORGANIGRAMA - CIA",
                campo_organigrama=6)
            reglas = leer_matrices(path, (config,))["CIA"]
        self.assertEqual((reglas[0].roles_texto, reglas[0].organigrama_texto), ("1", "2-4"))

    def test_txt_con_cantidad_de_campos_incorrecta_falla_con_fila(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dominio.txt"
            path.write_text("solo|tres|campos\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "fila 1: se esperaban 13 campos"):
                leer_personas_txt(path)

    def test_lector_de_sistema_usa_campos_configurados(self):
        config = SistemaConfig("CIA", "CIA.txt", "AA", "CIA", 3, 1, 6,
            columna_matriz_organigrama="AB", encabezado_matriz_organigrama="ORGANIGRAMA - CIA",
            campo_organigrama=6)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "cia.txt"
            path.write_text("1|b|login|d|e|4,9\n", encoding="utf-8")
            cuentas, errores = leer_sistema_txt(path, config)
        self.assertEqual((cuentas[0].login, cuentas[0].roles_texto, cuentas[0].organigrama_texto),
                         ("login", "1", "4,9"))
        self.assertEqual(errores, [])


class ProcesoTests(unittest.TestCase):
    """Validaciones y coordinación de la ejecución completa."""
    def test_configuracion_rechaza_posicion_de_campo_fuera_del_txt(self):
        config = SistemaConfig("CIA", "CIA.txt", "AA", "CIA", 3, 1, 6,
            columna_matriz_organigrama="AB", encabezado_matriz_organigrama="ORGANIGRAMA - CIA",
            campo_organigrama=7)
        with self.assertRaisesRegex(ValueError, r"organigrama \(7\).*entre 1 y 6"):
            ejecutar_revision("matriz.xlsx", "dominio.txt", {"CIA": "cia.txt"},
                              "reporte.xlsx", (config,))

    def test_ejecucion_acepta_seleccion_fragmentada_de_sistemas(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            matrix = root / "matriz.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Matriz de Accesos"
            ws["B1"] = "Cargo"
            ws["S1"] = "Argos Net"
            ws["R1"] = "ArgosAPX"
            ws["B2"] = "Analista"
            ws["S2"] = "1-2"
            ws["R2"] = "1-2"
            wb.save(matrix)
            domain = root / "dominio.txt"
            domain_fields = [""] * 13
            domain_fields[0], domain_fields[3], domain_fields[12] = "Persona Ficticia", "Analista", "usr1"
            domain.write_text("|".join(domain_fields) + "\n", encoding="utf-8")
            system_file = root / "argosnet.txt"
            system_file.write_text("usr1||x|1,2||x|BLOQUEADO\n", encoding="utf-8")
            apx_fields = [""] * 20
            apx_fields[1], apx_fields[4] = "usr1", "1,2"
            apx_file = root / "argosapx.txt"
            apx_file.write_text("|".join(apx_fields) + "\n", encoding="utf-8")
            output = root / "reporte.xlsx"
            argosnet = SistemaConfig("ArgosNet", "ArgosNet.txt", "S", "Argos Net", 1, 4, 7,
                                     null_sin_roles=True)
            apx = SistemaConfig("ArgosAPX", "ArgosAPX.txt", "R", "ArgosAPX", 2, 5, 20,
                                null_sin_roles=True)
            generated, result = ejecutar_revision(matrix, domain,
                {"ArgosNet": system_file, "ArgosAPX": apx_file}, output, (argosnet, apx),
                fuente_personas="Dominio")
            # No mantener el libro temporal bloqueado al salir de la carpeta.
            report = load_workbook(generated, data_only=True)
        self.assertEqual([s["nombre"] for s in result["sistemas"]], ["ArgosNet", "ArgosAPX"])
        self.assertIn("Observaciones ArgosNet", report.sheetnames)
        self.assertIn("Observaciones ArgosAPX", report.sheetnames)
        self.assertNotIn("Observaciones CIA", report.sheetnames)
        self.assertEqual([d["resultado"] for d in result["detalle"]], ["Conforme", "Conforme"])

if __name__ == "__main__":
    unittest.main()
