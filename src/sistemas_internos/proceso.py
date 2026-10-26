"""Coordinación del análisis compartida por consola e interfaz gráfica."""
from pathlib import Path

from .comparador import comparar
from .lectores import (SistemaConfig, leer_accesos_temporales, leer_personas_txt,
                       DOMINIO_CONFIG, MatrizConfig, PersonasConfig, SALAR_CONFIG,
                       TemporalesConfig, leer_matrices, leer_salar, leer_sistema_txt,
                       normalizar_cargo)
from .reporte import generar_reporte


def ejecutar_revision(matriz: str | Path, dominio: str | Path,
                     archivos_sistema: dict[str, str | Path], salida: str | Path,
                     sistemas: tuple[SistemaConfig, ...],
                     archivo_temporales: str | Path | None = None,
                     fuente_personas: str = "SALAR",
                     config_matriz: MatrizConfig = MatrizConfig(),
                     config_personas: PersonasConfig | None = None,
                     config_temporales: TemporalesConfig = TemporalesConfig(),
                     archivo_dominio_externo: str | Path | None = None,
                     config_dominio: PersonasConfig = DOMINIO_CONFIG,
                     incluir_cargos_matriz: bool = False,
                     permitir_sin_sistemas: bool = False) -> tuple[Path, dict]:
    """Ejecuta solo los sistemas seleccionados y genera un único reporte."""
    # Comprueba que la selección de sistemas y sus posiciones de columnas sean utilizables.
    if not sistemas and not permitir_sin_sistemas:
        raise ValueError("Selecciona al menos un sistema para analizar.")
    configuraciones = {config.nombre: config for config in sistemas}
    faltantes = [nombre for nombre in configuraciones if nombre not in archivos_sistema]
    if faltantes:
        raise ValueError("Falta seleccionar el TXT para: " + ", ".join(faltantes))
    for config in sistemas:
        posiciones = {"login": config.campo_login, "roles": config.campo_roles}
        if config.campo_organigrama:
            posiciones["organigrama"] = config.campo_organigrama
        if config.campo_estado:
            posiciones["estado"] = config.campo_estado
        for campo in ("campo_nombre", "campo_cargo", "campo_oficina", "campo_gerencia"):
            posicion = getattr(config, campo)
            if posicion:
                posiciones[campo.removeprefix("campo_")] = posicion
        invalidas = [f"{campo} ({posicion})" for campo, posicion in posiciones.items()
                     if not isinstance(posicion, int) or posicion < 1 or posicion > config.campos_txt]
        if invalidas:
            raise ValueError(f"Configuración de {config.nombre}: los campos {', '.join(invalidas)} "
                             f"deben estar entre 1 y {config.campos_txt}.")

    # Lee la fuente principal; al usar SALAR también puede cargar Dominio para conciliar personas.
    es_salar = fuente_personas.strip().casefold() == "salar"
    if config_personas is None:
        config_personas = SALAR_CONFIG if es_salar else DOMINIO_CONFIG
    personas, errores_dominio = (leer_salar(dominio, config_personas) if es_salar
                                 else leer_personas_txt(dominio, config_personas))
    personas_dominio = []
    errores_dominio_externo = []
    if es_salar and archivo_dominio_externo:
        personas_dominio, errores_dominio_externo = leer_personas_txt(
            archivo_dominio_externo, config_dominio)
        dominio_por_nombre = {}
        for persona_dominio in personas_dominio:
            nombre_key = normalizar_cargo(persona_dominio.nombre)
            if nombre_key:
                dominio_por_nombre.setdefault(nombre_key, []).append(persona_dominio)
        # SALAR sin correo puede tomar login y correo de una coincidencia única por nombre en Dominio.
        for persona in personas:
            if persona.correo.strip():
                continue
            coincidencias = dominio_por_nombre.get(normalizar_cargo(persona.nombre), [])
            if len(coincidencias) == 1:
                persona.login_salar = persona.login
                persona.login = coincidencias[0].login
                persona.correo = coincidencias[0].correo
            elif not coincidencias:
                persona.sin_coincidencia_dominio = True
            else:
                errores_dominio_externo.append({
                    "archivo": Path(archivo_dominio_externo).name,
                    "fila": persona.fila,
                    "campo": "nombre (coincidencia SALAR/Dominio)",
                    "descripcion": (f"El nombre {persona.nombre!r} aparece en varias filas de Dominio; "
                                    "no se eligió un login automáticamente."),
                })
    # Carga una vez las reglas de matriz y los accesos temporales para los sistemas seleccionados.
    reglas_sistemas = leer_matrices(matriz, sistemas, config_matriz)
    accesos_temporales = (leer_accesos_temporales(archivo_temporales, config_temporales)
                          if archivo_temporales else [])
    resultado = {"observaciones": [],
                 "errores": list(errores_dominio) + list(errores_dominio_externo), "detalle": [],
                 "dominios": personas, "fuente_personas": fuente_personas, "sistemas": []}
    if incluir_cargos_matriz:
        # Reutiliza los cargos que ya leyÃ³ la matriz para clasificar registros de Dominio.
        resultado["_cargos_matriz"] = {
            normalizar_cargo(regla.cargo)
            for reglas in reglas_sistemas.values()
            for regla in reglas
        }
    archivo_dominio_nombre = (Path(archivo_dominio_externo).name
                               if archivo_dominio_externo else Path(dominio).name)
    archivos = {"matriz": Path(matriz).name, "dominio": archivo_dominio_nombre,
                "fuente_personas": Path(dominio).name, "sistemas": {},
                "dominio_externo_consultado": bool(es_salar and archivo_dominio_externo)}
    if archivo_temporales:
        archivos["temporales"] = Path(archivo_temporales).name
    # Procesa cada sistema con su propia configuración y acumula sus resultados.
    for config in sistemas:
        path_sistema = Path(archivos_sistema[config.nombre])
        cuentas, errores_txt = leer_sistema_txt(path_sistema, config)
        reglas = reglas_sistemas[config.nombre]
        archivos["sistemas"][config.nombre] = path_sistema.name
        parcial = comparar(personas, cuentas, reglas,
            {"dominio": archivo_dominio_nombre, "fuente_personas": Path(dominio).name,
             "etiqueta_fuente_personas": fuente_personas, "sistema": path_sistema.name}, config,
            accesos_temporales, personas_dominio)
        resultado["observaciones"].extend(parcial["observaciones"])
        resultado["detalle"].extend(parcial["detalle"])
        resultado["errores"].extend(errores_txt + parcial["errores"])
        resultado["sistemas"].append({"nombre": config.nombre, "archivo": path_sistema.name,
            "registros": len(cuentas), "reglas": len(reglas),
            "coincidentes": parcial["coincidentes"],
            "sin_coincidencia": parcial["sin_coincidencia"]})
    # Genera un único Excel con resumen, observaciones, detalle y errores de carga.
    generado = generar_reporte(resultado, archivos, salida,
                               tuple(config.nombre for config in sistemas),
                               {config.nombre: {
                                   "oficina": bool(config.campo_oficina),
                                   "gerencia": bool(config.campo_gerencia),
                               } for config in sistemas})
    return generado, resultado
