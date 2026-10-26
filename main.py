"""Coordina la revisión de los sistemas configurados para esta etapa."""
import argparse
from pathlib import Path

from src.sistemas_internos.lectores import ErrorCarga, SistemaConfig
from src.sistemas_internos.proceso import ejecutar_revision


# Configuración reutilizable: define columnas y reglas propias de cada sistema institucional.
SISTEMAS = (
    SistemaConfig("ArgosAPX", "ArgosAPX.txt", "R", "ArgosAPX", 2, 5, 20,
                  null_sin_roles=True, campo_nombre=3, campo_cargo=14,
                  campo_oficina=15, campo_gerencia=16),
    SistemaConfig("ArgosNet", "ArgosNet.txt", "S", "Argos Net", 1, 4, 7,
                  null_sin_roles=True, campo_oficina=6),
    SistemaConfig("BDPSAM", "BDPSAM.txt", "W", "Argos Formularios 2P", 1, 5, 6,
                  prefijo_login="bdp"),
    SistemaConfig("CIA", "CIA.txt", "AA", "CIA", 3, 1, 6,
                  columna_matriz_organigrama="AB",
                  encabezado_matriz_organigrama="ORGANIGRAMA - CIA",
                  campo_organigrama=6),
    SistemaConfig("FG", "FG.txt", "X", "Fondo de Garantias", 1, 5, 6,
                  prefijo_login="bdp", sufijo_login="_fg"),
    SistemaConfig("SEGA", "SEGA.txt", "Z", "SEGA", 3, 1, 4),
    SistemaConfig("SRGNC", "SRGNC.txt", "AO", "SRGNC", 3, 5, 6,
                  campo_estado=4, estado_inactivo="INA", clave_cruce="correo",
                  campo_login_respaldo=2, prefijo_login_respaldo="BDP"),
    SistemaConfig("PRESUPUESTO", "PRESUPUESTO.txt", "BD", "PRESUPUESTO", 1, 6, 6,
                  agrupar_roles_por_login=True, normalizar_codigos_rol=True),
    SistemaConfig("PUNTO RECLAMO", "PUNTO RECLAMO.txt", "AD", "Punto de Reclamo (PR) BDP", 3, 6, 6,
                  campo_estado=6, estado_activo="Activo", estado_inactivo="Inactivo",
                  campo_oficina=4,
                  validar_solo_existencia=True),
)


def main() -> int:
    # Permite ejecutar desde consola con rutas por defecto o sobrescritas por argumentos.
    raiz = Path(__file__).resolve().parent
    docs = raiz / "documentos base"
    parser = argparse.ArgumentParser(description="Compara SALAR con los sistemas institucionales configurados y usa Dominio como fuente complementaria.")
    parser.add_argument("--matriz", default=docs / "prueba.xlsx", type=Path)
    parser.add_argument("--dominio", default=docs / "DOMINIO.txt", type=Path,
                        help="Dominio complementario para conciliar usuarios que no aparecen en SALAR.")
    parser.add_argument("--entrada", default=docs, type=Path,
                        help="Carpeta con salar.txt y los TXT de sistemas configurados.")
    parser.add_argument("--salida", default=raiz / "data/salida/revision_accesos.xlsx", type=Path)
    archivo_temporales = docs / "usuarios accesos temporales sga.xlsx"
    parser.add_argument("--temporales", default=archivo_temporales if archivo_temporales.is_file() else None,
                        type=Path, help="Excel opcional de usuarios con acceso temporal.")
    args = parser.parse_args()
    try:
        archivos_sistema = {config.nombre: args.entrada / config.archivo for config in SISTEMAS}
        salar = args.entrada / "salar.txt"
        generado, resultado = ejecutar_revision(args.matriz, salar,
            archivos_sistema, args.salida, SISTEMAS, args.temporales, "SALAR",
            archivo_dominio_externo=args.dominio)
    except ErrorCarga as exc:
        parser.exit(2, f"Error de carga: {exc}\n")
    print(f"Reporte generado: {generado}")
    print(f"Sistemas: {len(SISTEMAS)}; observaciones: {len(resultado['observaciones'])}; "
          f"errores de carga: {len(resultado['errores'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
