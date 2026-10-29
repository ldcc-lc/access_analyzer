"""Reglas de comparación reutilizables para todos los sistemas."""
from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from datetime import date

from .lectores import (CuentaSistema, Persona, ReglaCargo, SistemaConfig,
                      _normalizar_codigos_rol, normalizar_cargo,
                      normalizar_login, normalizar_login_sistema)


# Interpreta el texto de roles como conjunto de codigos y reporta formatos no reconocidos.
def _roles(texto: str, separador: str, null_sin_roles: bool = False,
           normalizar_codigos: bool = False) -> tuple[set[str], str | None]:
    valor = (texto or "").strip()
    if not valor or valor == "-":
        return set(), None
    if valor.casefold() == "null":
        return (set(), None) if null_sin_roles else (set(), "El significado de 'Null' para este sistema no está confirmado.")
    if normalizar_codigos:
        valor = _normalizar_codigos_rol(valor)
        if valor is None: 
            return set(), "Formato de codigos de rol no reconocido."
        separador = ","
    partes = [parte.strip() for parte in valor.split(separador)]
    if any(not parte for parte in partes) or any(not re.fullmatch(r"\d+", p) for p in partes):
        return set(), f"Formato de roles no reconocido: {valor!r}."
    return set(partes), None


def comparar(personas: list[Persona], cuentas: list[CuentaSistema], reglas: list[ReglaCargo],
             archivos: dict[str, str], config: SistemaConfig,
             accesos_temporales: list[dict[str, str]] | None = None,
             personas_dominio: list[Persona] | None = None) -> dict:
    """Compara las cuentas del sistema con las personas y los accesos esperados según la matriz."""
    # Obtiene etiquetas y nombres de archivo que se incluirán en errores y observaciones.
    sistema = config.nombre
    fuente_personas = archivos.get("etiqueta_fuente_personas", "Dominio")
    archivo_personas = archivos.get("fuente_personas", archivos["dominio"])
    archivo_sistema = archivos.get("sistema", config.archivo)
    # Solo se aceptan los dos criterios de identidad soportados por los lectores.
    if config.clave_cruce not in {"login", "correo"}:
        raise ValueError(f"Clave de cruce no reconocida para {sistema}: {config.clave_cruce!r}.")
    # Acumula por separado las alertas, los errores de datos y el detalle de cada registro.
    observaciones, errores, detalle = [], [], []
    # Prepara índices de la matriz por cargo para consultar rápidamente los accesos esperados.
    # La clave de cada índice es el cargo normalizado.
    reglas_por_cargo: dict[str, list[ReglaCargo]] = defaultdict(list)
    roles_matrix, problemas_matrix, acceso_esperado_matriz = {}, {}, {}
    # Normaliza cargos y convierte los accesos de matriz a conjuntos comparables.
    for regla in reglas:
        key = normalizar_cargo(regla.cargo)
        reglas_por_cargo[key].append(regla)
        acceso_esperado_matriz[key] = _valor_indica_acceso(regla.roles_texto)
        roles, problema = (set(), None) if config.validar_solo_existencia else _roles(
            regla.roles_texto, "-", normalizar_codigos=config.normalizar_codigos_rol)
        # La matriz CIA agrega notas aclaratorias entre paréntesis en la misma
        # celda. Esas notas no forman parte de los códigos que se comparan.
        texto_organigrama = re.sub(r"\([^)]*\)", "", regla.organigrama_texto)
        organigrama, problema_org = _roles(texto_organigrama,
            config.separador_organigrama_matriz) if config.columna_matriz_organigrama else (set(), None)
        if problema or problema_org:
            problemas_matrix[key] = problema or problema_org
        else:
            roles_matrix[key] = roles
            if config.columna_matriz_organigrama:
                roles_matrix[(key, "organigrama")] = organigrama
    # Si un cargo aparece varias veces en la matriz, su regla es ambigua.
    for key, items in reglas_por_cargo.items():
        if len(items) > 1:
            problemas_matrix[key] = "Cargo duplicado en la matriz; no se puede elegir una regla sin ambigüedad."

    # Indexa la fuente principal de personas y las cuentas por la clave configurada.
    # Este diccionario contiene `personas` (SALAR o Dominio), pese a su nombre histórico.
    dominio_por_login: dict[str, list[Persona]] = defaultdict(list)
    sistema_por_login: dict[str, list[CuentaSistema]] = defaultdict(list)
    sistema_por_alias_temporal: dict[str, set[str]] = defaultdict(set)
    # Construye claves de personas y registra correos vacíos cuando el cruce depende del correo.
    for persona in personas:
        key = _clave_persona(persona, config)
        if key:
            dominio_por_login[key].append(persona)
        elif config.clave_cruce == "correo" and not config.campo_login_respaldo:
            errores.append(_error(archivo_personas, persona.fila, "correo",
                "Correo vacío; no se puede cruzar esta persona con SRGNC."))
    # Construye claves de cuentas y guarda aliases para relacionar después accesos temporales.
    for cuenta in cuentas:
        key = _clave_cuenta(cuenta, config, dominio_por_login)
        if key:
            sistema_por_login[key].append(cuenta)
            for alias in _aliases_temporales(cuenta, config):
                sistema_por_alias_temporal[alias].add(key)
        elif config.clave_cruce == "correo":
            errores.append(_error(archivo_sistema, cuenta.fila, "correo",
                "Correo vacío en el campo configurado; no se puede cruzar la cuenta."))

    # Detecta claves duplicadas en la fuente de personas tras normalizarlas.
    for key, rows in dominio_por_login.items():
        if len(rows) > 1:
            for p in rows:
                campo = ("login" if key.startswith("login:") else config.clave_cruce)
                valor = p.login if campo == "login" else p.correo
                errores.append(_error(archivo_personas, p.fila, campo,
                                      f"Clave {campo} duplicada tras normalizar: {valor!r}."))
    # Detecta claves duplicadas en el archivo del sistema.
    for key, rows in sistema_por_login.items():
        if len(rows) > 1:
            for c in rows:
                campo = "login respaldo" if key.startswith("login:") else config.clave_cruce
                valor = c.login_respaldo if key.startswith("login:") else c.login
                errores.append(_error(archivo_sistema, c.fila, campo,
                                      f"Clave {campo} duplicada tras normalizar: {valor!r}."))

    # Revisa cada cuenta del sistema contra la fuente principal de personas.
    for key, rows in sistema_por_login.items():
        c = rows[0]
        dominio_rows = dominio_por_login.get(key, [])
        # Varias cuentas con la misma clave hacen que el caso sea no evaluable.
        if len(rows) > 1:
            encontrado = " | ".join(_access_text(x.roles_texto, x.organigrama_texto, config) for x in rows)
            detalle.append(_detalle(c.login, "", "", sistema, encontrado,
                                    "No evaluable: login duplicado en sistema"))
            observaciones.append(_obs("Duplicados o casos especiales", c.login, sistema=sistema,
                encontrado=encontrado,
                motivo="Hay varias filas con este login normalizado; no se eligió una cuenta para comparar.",
                origen_acceso=archivo_sistema))
            continue
        # Si no hay persona en la fuente principal, intenta encontrarla en Dominio.
        if not dominio_rows:
            # En sistemas con control de estado, una cuenta inactiva externa no requiere más cruce.
            if (config.validar_solo_existencia
                    and c.estado.strip().casefold() != config.estado_activo.strip().casefold()):
                detalle.append(_detalle(c.login, "", "", sistema, c.estado,
                    "Cuenta inactiva fuera de la fuente de personas"))
                continue
            personas_externas = _buscar_en_dominio(c, personas_dominio or [], config)
            # Una única coincidencia permite consultar la matriz con el cargo de Dominio.
            if len(personas_externas) == 1:
                persona_externa = personas_externas[0]
                cargo_dominio = normalizar_cargo(persona_externa.cargo)
                esperado_dominio, resultado_dominio, encontrado_dominio = _evaluar_cuenta_por_cargo(
                    persona_externa.cargo, c, reglas_por_cargo, roles_matrix,
                    problemas_matrix, acceso_esperado_matriz, config)
                encontrado = encontrado_dominio
                detalle_dominio = _detalle(c.login, persona_externa.nombre,
                    persona_externa.cargo, sistema, encontrado,
                    "Usuario Externo (Dominio)" if not esperado_dominio else resultado_dominio)
                detalle_dominio["esperados"] = esperado_dominio
                detalle.append(detalle_dominio)
                motivo_externo = (
                    "El usuario no aparece en SALAR, sí en Dominio y su cargo se evaluó con la matriz."
                    if cargo_dominio in reglas_por_cargo else
                    "El usuario no aparece en SALAR, sí en Dominio, pero su cargo no está en la matriz."
                )
                observaciones.append(_obs("Usuario Externo (Dominio)", c.login,
                    nombre=persona_externa.nombre, cargo=persona_externa.cargo,
                    sistema=sistema, esperado=esperado_dominio, encontrado=encontrado,
                    nombre_sistema=c.nombre, cargo_sistema=c.cargo,
                    oficina_sistema=c.oficina, oficina_salar=persona_externa.oficina,
                    gerencia_sistema=c.gerencia, gerencia_salar=persona_externa.gerencia,
                    motivo=motivo_externo,
                    origen_cargo=archivo_personas, origen_acceso=archivo_sistema,
                    cargo_en_matriz=cargo_dominio in reglas_por_cargo
                        and cargo_dominio not in problemas_matrix))
                continue
            # Varias coincidencias son ambiguas; no se asigna la cuenta a una persona.
            if len(personas_externas) > 1:
                encontrados = " | ".join(p.nombre for p in personas_externas if p.nombre)
                detalle.append(_detalle(c.login, encontrados, "", sistema,
                    _access_text(c.roles_texto, c.organigrama_texto, config),
                    "No evaluable: identidad duplicada en Dominio"))
                observaciones.append(_obs("Duplicados o casos especiales", c.login,
                    nombre=encontrados, sistema=sistema,
                    encontrado=_access_text(c.roles_texto, c.organigrama_texto, config),
                    motivo="El login o correo coincide con varias filas de Dominio; no se clasificó como usuario externo.",
                    origen_cargo=archivo_personas, origen_acceso=archivo_sistema))
                continue
            # Sin coincidencia única, registra la cuenta como fuera de la fuente principal.
            persona_en_dominio = (personas_externas[0] if personas_externas else None)
            encontrado = _access_text(c.roles_texto, c.organigrama_texto, config)
            esperado_dominio = resultado_dominio = ""
            if persona_en_dominio:
                esperado_dominio, resultado_dominio, encontrado = _evaluar_cuenta_por_cargo(
                    persona_en_dominio.cargo, c, reglas_por_cargo, roles_matrix,
                    problemas_matrix, acceso_esperado_matriz, config)
            detalle_dominio = _detalle(c.login,
                                    persona_en_dominio.nombre if persona_en_dominio else "",
                                    persona_en_dominio.cargo if persona_en_dominio else "",
                                    sistema, encontrado,
                                    (f"Fuera de {fuente_personas}: acceso pendiente de retirar; {resultado_dominio}"
                                     if esperado_dominio else
                                     f"Fuera de {fuente_personas}: acceso pendiente de retirar"))
            detalle_dominio["esperados"] = esperado_dominio
            detalle.append(detalle_dominio)
            motivo_fuera = f"El login no aparece en {fuente_personas}; el acceso queda pendiente de revisar o retirar."
            if persona_en_dominio:
                motivo_fuera = ("El usuario no aparece en SALAR, pero sí en Dominio; "
                    "su cargo coincide con uno de la matriz y requiere revisión como posible personal de planta.")
            observaciones.append(_obs(f"Usuario fuera de {fuente_personas} con acceso pendiente de retirar", c.login,
                nombre=persona_en_dominio.nombre if persona_en_dominio else "",
                cargo=persona_en_dominio.cargo if persona_en_dominio else "",
                oficina_salar=persona_en_dominio.oficina if persona_en_dominio else "",
                gerencia_salar=persona_en_dominio.gerencia if persona_en_dominio else "",
                sistema=sistema, esperado=esperado_dominio, encontrado=encontrado,
                nombre_sistema=c.nombre, cargo_sistema=c.cargo, oficina_sistema=c.oficina,
                gerencia_sistema=c.gerencia,
                motivo=motivo_fuera,
                origen_acceso=archivo_sistema))
            continue
        # Si la persona está duplicada en la fuente principal, no se comparan sus accesos.
        if len(dominio_rows) > 1:
            encontrado = _access_text(c.roles_texto, c.organigrama_texto, config)
            detalle.append(_detalle(c.login, "", "", sistema, encontrado,
                                    f"No evaluable: login duplicado en {fuente_personas}"))
            observaciones.append(_obs("Duplicados o casos especiales", c.login, sistema=sistema,
                encontrado=encontrado,
                motivo=f"Hay varias filas de {fuente_personas} con este login normalizado; no se eligió una persona para comparar.",
                origen_cargo=archivo_personas, origen_acceso=archivo_sistema))
            continue

        # Desde aquí existe una sola persona asociada a esta cuenta.
        p = dominio_rows[0]
        # La persona de SALAR sin coincidencia de Dominio se alerta aunque sí tenga cuenta.
        if (fuente_personas.strip().casefold() == "salar"
                and p.sin_coincidencia_dominio):
            observaciones.append(_obs("Usuario sin Dominio pendiente de creación",
                p.login, nombre=p.nombre, cargo=p.cargo, sistema=sistema,
                encontrado=_access_text(c.roles_texto, c.organigrama_texto, config),
                motivo="La persona aparece en SALAR sin correo y no se encontró un nombre idéntico en Dominio.",
                login_salar=p.login_salar or p.login,
                origen_cargo=archivo_personas, origen_acceso=archivo_sistema))
        # Compara los datos organizacionales de SALAR con los campos configurados del sistema.
        if (fuente_personas.strip().casefold() == "salar"
                and any((config.campo_cargo, config.campo_oficina, config.campo_gerencia))):
            _comparar_datos_persona(p, c, sistema, archivo_personas, archivo_sistema,
                                    observaciones, config, fuente_personas)
        key_cargo = normalizar_cargo(p.cargo)
        # Guarda los datos comunes para adjuntarlos luego a observaciones del caso.
        base = dict(login=c.login, login_salar=p.login_salar, nombre=p.nombre, cargo=p.cargo, sistema=sistema,
                    nombre_sistema=c.nombre, cargo_sistema=c.cargo,
                    oficina_sistema=c.oficina, oficina_salar=p.oficina,
                    gerencia_sistema=c.gerencia, gerencia_salar=p.gerencia,
                    origen_cargo=archivo_personas, origen_acceso=archivo_sistema)
        # Algunos sistemas se validan por existencia/estado, no comparando códigos de rol.
        if config.validar_solo_existencia:
            if key_cargo not in reglas_por_cargo or key_cargo in problemas_matrix:
                motivo = problemas_matrix.get(key_cargo,
                    "El cargo no aparece en la matriz; no se puede evaluar la existencia esperada.")
                detalle.append(_detalle(c.login, p.nombre, p.cargo, sistema, c.estado,
                                        "No evaluable: cargo sin regla de matriz"))
                observaciones.append(_obs("Duplicados o casos especiales", p.login,
                    nombre=p.nombre, cargo=p.cargo, sistema=sistema, motivo=motivo,
                    **_origen(base)))
                continue
            # Para este tipo de sistema solo importa si se espera cuenta y si está activa.
            esperado_existe = acceso_esperado_matriz[key_cargo]
            esperado_matriz = _mostrar_acceso_esperado(
                key_cargo, reglas_por_cargo, acceso_esperado_matriz)
            activo = c.estado.strip().casefold() == config.estado_activo.strip().casefold()
            if esperado_existe and not activo:
                categoria = "Estado distinto de Activo"
                observaciones.append(_obs(categoria, p.login, nombre=p.nombre, cargo=p.cargo,
                    sistema=sistema, esperado=esperado_matriz,
                    encontrado=c.estado or "(vacío)",
                    motivo="La matriz exige una cuenta y su estado debe ser Activo.",
                    **_origen(base)))
                resultado_estado = categoria
            elif not esperado_existe and activo:
                categoria = "Registro en sistema sin acceso esperado"
                observaciones.append(_obs(categoria, p.login, nombre=p.nombre, cargo=p.cargo,
                    sistema=sistema, esperado=esperado_matriz,
                    encontrado=c.estado or "(vacío)",
                    motivo="Existe una cuenta, pero la matriz no indica acceso para el cargo.",
                    **_origen(base)))
                resultado_estado = categoria
            elif not esperado_existe:
                resultado_estado = "Conforme"
            else:
                resultado_estado = "Conforme"
            detalle.append({**_detalle(c.login, p.nombre, p.cargo, sistema, c.estado,
                resultado_estado), "esperados": esperado_matriz})
            continue
        # Sin regla para el cargo en matriz, no hay roles esperados que comparar.
        if key_cargo not in reglas_por_cargo:
            encontrado = _access_text(c.roles_texto, c.organigrama_texto, config)
            if _tiene_acceso(c, config):
                detalle.append(_detalle(c.login, p.nombre, p.cargo, sistema, encontrado,
                                        "Cargo sin correspondencia en matriz"))
                observaciones.append(_obs("Cargo sin correspondencia en matriz", p.login,
                    nombre=p.nombre, cargo=p.cargo,
                    sistema=sistema, encontrado=encontrado,
                    motivo="El cargo no aparece en la matriz y se encontró acceso en el sistema.",
                    **_origen(base)))
            else:
                detalle.append(_detalle(c.login, p.nombre, p.cargo, sistema, encontrado,
                                        "Sin acceso encontrado; cargo sin correspondencia en matriz"))
            continue
        # Evita comparar si la regla de matriz está duplicada o tiene formato inválido.
        if key_cargo in problemas_matrix:
            encontrado = _access_text(c.roles_texto, c.organigrama_texto, config)
            detalle.append(_detalle(c.login, p.nombre, p.cargo, sistema, encontrado,
                                    "No evaluable: matriz ambigua o inválida"))
            observaciones.append(_obs("Duplicados o casos especiales", p.login, nombre=p.nombre,
                cargo=p.cargo, sistema=sistema,
                encontrado=encontrado, motivo=problemas_matrix[key_cargo], **_origen(base)))
            continue
        # Interpreta roles encontrados y esperados; también organigrama si este sistema lo usa.
        encontrados, problema = _roles(c.roles_texto, config.separador_roles,
            config.null_sin_roles, config.normalizar_codigos_rol)
        esperados = roles_matrix[key_cargo]
        encontrados_org, problema_org = _roles(c.organigrama_texto, config.separador_organigrama_txt) \
            if config.campo_organigrama else (set(), None)
        esperados_org = roles_matrix.get((key_cargo, "organigrama"), set())
        # Un formato desconocido impide concluir si los accesos son conformes.
        if problema or problema_org:
            encontrado = _access_text(c.roles_texto, c.organigrama_texto, config)
            detalle.append(_detalle(c.login, p.nombre, p.cargo, sistema, encontrado,
                                    "No evaluable: formato o significado de roles"))
            observaciones.append(_obs("Duplicados o casos especiales", p.login, nombre=p.nombre,
                cargo=p.cargo, sistema=sistema,
                esperado=_expected_access_text(_fmt(esperados), _fmt(esperados_org), config),
                encontrado=encontrado, motivo=problema or problema_org, **_origen(base)))
            continue
        # Calcula accesos faltantes y adicionales como diferencias entre conjuntos.
        faltantes, adicionales = esperados - encontrados, encontrados - esperados
        faltantes_org, adicionales_org = esperados_org - encontrados_org, encontrados_org - esperados_org
        acceso_esperado = _expected_access_text(_fmt(esperados), _fmt(esperados_org), config)
        acceso_encontrado = _access_text(_fmt(encontrados), _fmt(encontrados_org), config)
        if faltantes or adicionales or faltantes_org or adicionales_org:
            observaciones.append(_obs("Diferencia de roles", p.login, nombre=p.nombre, cargo=p.cargo,
                sistema=sistema, esperado=acceso_esperado, encontrado=acceso_encontrado,
                motivo="El acceso encontrado difiere del esperado en la matriz.", **_origen(base)))
        detalle.append({"login": c.login, "nombre": p.nombre, "cargo": p.cargo, "sistema": sistema,
            "esperados": acceso_esperado, "encontrados": acceso_encontrado,
            "resultado": "Conforme" if not (faltantes or adicionales or faltantes_org or adicionales_org)
                        else "Diferencia de roles"})

    # Busca personas de la fuente principal que no tienen una cuenta en este sistema.
    for key, rows in dominio_por_login.items():
        if key in sistema_por_login or len(rows) != 1:
            continue
        p = rows[0]
        expectativa_temporal = _expectativa_temporal_para_alta(
            p, accesos_temporales, reglas_por_cargo, roles_matrix,
            problemas_matrix, acceso_esperado_matriz, config)
        cargo_observado = p.cargo
        motivo_alta = (f"La matriz espera acceso y el login no aparece en el sistema; "
                       f"la fuente consultada es {fuente_personas}.")
        if expectativa_temporal is not None:
            cargo_observado = expectativa_temporal["cargo"] or p.cargo
            if expectativa_temporal["problema"]:
                problema = expectativa_temporal["problema"]
                detalle.append(_detalle(p.login, p.nombre, cargo_observado, sistema, "",
                                        f"No evaluable: {problema}"))
                observaciones.append(_obs("Duplicados o casos especiales", p.login,
                    nombre=p.nombre, cargo=cargo_observado, sistema=sistema,
                    esperado=problema,
                    motivo="La asignacion temporal impide determinar un cargo esperado sin ambiguedad.",
                    origen_cargo=archivo_personas))
                continue
            esperado = expectativa_temporal["esperado"]
            tiene_acceso_esperado = expectativa_temporal["tiene_acceso"]
            motivo_alta = ("No existe cuenta en el sistema. La expectativa se calculo "
                           "con la asignacion temporal vigente o el cargo de origen si ya vencio.")
        else:
            key_cargo = normalizar_cargo(p.cargo)
            if key_cargo not in reglas_por_cargo:
                detalle.append(_detalle(p.login, p.nombre, p.cargo, sistema, "",
                                        "Sin registro de sistema; cargo sin correspondencia en matriz"))
                observaciones.append(_obs("Cargo sin correspondencia en matriz", p.login,
                    nombre=p.nombre, cargo=p.cargo, sistema=sistema,
                    motivo="El cargo no aparece en la matriz y el login no tiene registro en el sistema.",
                    origen_cargo=archivo_personas))
                continue
            if key_cargo in problemas_matrix:
                detalle.append(_detalle(p.login, p.nombre, p.cargo, sistema, "",
                                        "No evaluable: matriz ambigua o invalida"))
                observaciones.append(_obs("Duplicados o casos especiales", p.login,
                    nombre=p.nombre, cargo=p.cargo, sistema=sistema,
                    motivo=problemas_matrix[key_cargo], origen_cargo=archivo_personas))
                continue
            tiene_acceso_esperado = (
                acceso_esperado_matriz.get(key_cargo, False)
                if config.validar_solo_existencia else
                bool(roles_matrix[key_cargo]
                     or roles_matrix.get((key_cargo, "organigrama"), set())))
            esperado = (_mostrar_acceso_esperado(
                key_cargo, reglas_por_cargo, acceso_esperado_matriz)
                if config.validar_solo_existencia else
                _expected_access_text(_fmt(roles_matrix[key_cargo]),
                    _fmt(roles_matrix.get((key_cargo, "organigrama"), set())), config))

        if tiene_acceso_esperado:
            detalle.append({**_detalle(p.login, p.nombre, cargo_observado, sistema, "",
                                       "Sin registro en sistema"),
                            "esperados": esperado})
            categoria_sin_registro = (f"Usuarios en {fuente_personas} sin registro en sistema"
                if fuente_personas.strip().casefold() == "salar"
                else f"Usuario en {fuente_personas} sin registro en sistema")
            if (fuente_personas.strip().casefold() == "salar"
                    and p.sin_coincidencia_dominio):
                categoria_sin_registro = "Usuario sin Dominio pendiente de creacion"
            observaciones.append(_obs(categoria_sin_registro, p.login,
                nombre=p.nombre, cargo=cargo_observado, sistema=sistema, esperado=esperado,
                login_salar=p.login_salar,
                nombre_sistema="", cargo_sistema="", oficina_sistema="", gerencia_sistema="",
                oficina_salar=p.oficina, gerencia_salar=p.gerencia,
                motivo=motivo_alta, origen_cargo=archivo_personas))
        else:
            detalle.append(_detalle(p.login, p.nombre, cargo_observado, sistema, "",
                "Sin Acceso" if config.validar_solo_existencia else "No se espera cuenta"))
            if (fuente_personas.strip().casefold() == "salar"
                    and p.sin_coincidencia_dominio):
                observaciones.append(_obs("Usuario sin Dominio pendiente de creacion",
                    p.login, nombre=p.nombre, cargo=cargo_observado, sistema=sistema,
                    motivo="La persona aparece en SALAR sin correo y no se encontro un nombre identico en Dominio.",
                    login_salar=p.login_salar or p.login,
                    origen_cargo=archivo_personas))

    # Si SALAR es la fuente principal, busca además personas de Dominio no conciliadas con SALAR.
    if fuente_personas.strip().casefold() == "salar" and personas_dominio:
        for persona_dominio in _dominio_sin_persona_salar(personas, personas_dominio):
            # Solo alerta altas para personas de Dominio sin una cuenta activa en este sistema.
            if _persona_dominio_tiene_cuenta(persona_dominio, cuentas, config):
                continue
            aliases = _aliases_dominio(persona_dominio, config)
            if any(sistema_por_alias_temporal.get(alias) for alias in aliases):
                continue
            expectativa_temporal = _expectativa_temporal_para_alta(
                persona_dominio, accesos_temporales, reglas_por_cargo, roles_matrix,
                problemas_matrix, acceso_esperado_matriz, config)
            cargo_observado = persona_dominio.cargo
            if expectativa_temporal is not None:
                cargo_observado = expectativa_temporal["cargo"] or persona_dominio.cargo
                if expectativa_temporal["problema"] or not expectativa_temporal["tiene_acceso"]:
                    continue
                esperado = expectativa_temporal["esperado"]
                motivo_temporal_alta = (
                    "La expectativa se calculo con la asignacion temporal vigente "
                    "o con el cargo de origen si ya vencio.")
            else:
                key_cargo = normalizar_cargo(persona_dominio.cargo)
                if key_cargo not in reglas_por_cargo or key_cargo in problemas_matrix:
                    continue
                if config.validar_solo_existencia:
                    if not acceso_esperado_matriz.get(key_cargo, False):
                        continue
                    esperado = _mostrar_acceso_esperado(
                        key_cargo, reglas_por_cargo, acceso_esperado_matriz)
                else:
                    esperados = roles_matrix[key_cargo]
                    esperados_org = roles_matrix.get((key_cargo, "organigrama"), set())
                    # Si no hay roles esperados, no hace falta crear una alerta de alta.
                    if not esperados and not esperados_org:
                        continue
                    esperado = _expected_access_text(_fmt(esperados), _fmt(esperados_org), config)
                motivo_temporal_alta = ""
            categoria_alta = "Usuario en Dominio sin registro en sistema"
            detalle_alta = _detalle(persona_dominio.login, persona_dominio.nombre,
                cargo_observado, sistema, "Sin registro en sistema", categoria_alta)
            detalle_alta["esperados"] = esperado
            detalle.append(detalle_alta)
            observaciones.append(_obs(categoria_alta, persona_dominio.login,
                nombre=persona_dominio.nombre, cargo=cargo_observado,
                sistema=sistema, esperado=esperado, encontrado="Sin registro en sistema",
                oficina_salar=persona_dominio.oficina,
                gerencia_salar=persona_dominio.gerencia,
                motivo=("La persona aparece en Dominio, no tiene coincidencia en SALAR ni cuenta en este sistema, "
                        "y la matriz espera acceso segÃºn su cargo."),
                origen_cargo=archivos.get("dominio", archivo_personas),
                cargo_en_matriz=True))

    # Relaciona las filas temporales con las cuentas del TXT del sistema.
    temporales_por_login: dict[str, list[dict[str, str]]] = defaultdict(list)
    for temporal in accesos_temporales or []:
        alias = normalizar_login_sistema(temporal["login"], config)
        keys = sistema_por_alias_temporal.get(alias, set())
        for key in keys:
            temporales_por_login[key].append(temporal)
    # Sustituye alertas ordinarias por el tratamiento específico de accesos temporales.
    if temporales_por_login:
        # Una fila temporal sustituye la observación ordinaria del mismo login,
        # para que quede agrupada en su categoría específica.
        aliases_sistema_temporales = {
            alias
            for key in temporales_por_login
            for cuenta in sistema_por_login.get(key, [])
            for alias in _aliases_temporales(cuenta, config)
        }
        observaciones[:] = [o for o in observaciones
            if not aliases_sistema_temporales.intersection(
                _variantes_alias_temporal(o["login"], config))]
        detalle[:] = [d for d in detalle
            if not aliases_sistema_temporales.intersection(
                _variantes_alias_temporal(d["login"], config))]
        # Evalúa vigencia, cargos esperados, roles/estado y oficina para cada acceso temporal.
        for key, temporal_rows in temporales_por_login.items():
            cuentas_rows = sistema_por_login.get(key, [])
            personas_rows = dominio_por_login.get(key, [])
            if len(cuentas_rows) != 1 or len(personas_rows) > 1:
                continue
            cuenta = cuentas_rows[0]
            persona = personas_rows[0] if personas_rows else None
            for temporal in temporal_rows:
                cargo_actual = temporal.get("cargo_actual", "")
                cargo_designado = temporal.get("cargo_designado", "")
                nombre = persona.nombre if persona else temporal["nombre"]
                fecha_fin = temporal.get("fecha_fin", "")
                fin_vigencia = date.fromisoformat(fecha_fin) if fecha_fin else None
                es_fin_temporal = bool(fin_vigencia and date.today() > fin_vigencia)
                conservar_roles_actuales = (temporal.get("cargo_actual_pintado") and
                                             temporal.get("cargo_designado_pintado"))
                cargo_actual_pintado = temporal.get("cargo_actual_pintado", False)
                if es_fin_temporal:
                    cargos_roles = [cargo_actual]
                elif conservar_roles_actuales:
                    cargos_roles = [cargo_actual, cargo_designado]
                elif cargo_actual_pintado:
                    cargos_roles = [cargo_actual]
                else:
                    # Si solo Cargo Designado está pintado, o ninguna celda
                    # tiene relleno, se conserva el comportamiento de usarlo.
                    cargos_roles = [cargo_designado]
                segmentos_esperados = (_segmentos_esperados_temporales(
                    cargos_roles, reglas_por_cargo, roles_matrix, problemas_matrix, config)
                    if conservar_roles_actuales and not es_fin_temporal else None)
                expected_temporal = ("".join(texto for texto, _ in segmentos_esperados)
                    if segmentos_esperados else _esperados_temporales(cargos_roles,
                        reglas_por_cargo, roles_matrix, problemas_matrix, config))
                if config.validar_solo_existencia:
                    valores_esperados_temporal = list(dict.fromkeys(
                        _mostrar_acceso_esperado(key_temporal, reglas_por_cargo,
                            acceso_esperado_matriz)
                        for key_temporal in (normalizar_cargo(cargo) for cargo in cargos_roles)
                        if key_temporal in reglas_por_cargo
                        and key_temporal not in problemas_matrix
                        and acceso_esperado_matriz.get(key_temporal, False)))
                    acceso_esperado = bool(valores_esperados_temporal)
                    resultado_roles = ("Conforme" if acceso_esperado and
                        cuenta.estado.strip().casefold() == config.estado_activo.strip().casefold()
                        else "Estado distinto de Activo" if acceso_esperado
                        else "Registro en sistema sin acceso esperado" if cuenta.estado.strip().casefold() == config.estado_activo.strip().casefold()
                        else "Conforme")
                    expected_temporal = "; ".join(valores_esperados_temporal)
                    encontrado_temporal = cuenta.estado or "(vacío)"
                else:
                    resultado_roles = _estado_roles_temporales(cargos_roles, cuenta,
                        reglas_por_cargo, roles_matrix, problemas_matrix, config)
                    encontrado_temporal = _access_text(cuenta.roles_texto,
                        cuenta.organigrama_texto, config)
                resultado_cargo = (_estado_campo_temporal(cargo_designado, cuenta.cargo,
                    normalizar_cargo) if config.campo_cargo else "No configurado en el sistema")
                resultado_oficina = (_estado_campo_temporal(
                    temporal.get("oficina_designada", ""), cuenta.oficina,
                    _normalizar_texto) if config.campo_oficina else "No configurado en el sistema")
                categoria_temporal = ("Fin Acceso Temporal"
                                      if es_fin_temporal
                                      else "Con Acceso Temporal")
                motivo_temporal = f"Cargo original temporal: {temporal['cargo_actual']}."
                if es_fin_temporal:
                    motivo_temporal += " La vigencia terminó; se esperan únicamente los accesos del cargo de origen."
                elif conservar_roles_actuales:
                    motivo_temporal += " Ambas celdas de cargo están pintadas; se conservan los roles actuales y se agregan los del cargo designado."
                elif cargo_actual_pintado:
                    motivo_temporal += " Solo Cargo Actual está pintado; se esperan los roles del cargo actual."
                elif temporal.get("cargo_designado_pintado"):
                    motivo_temporal += " Solo Cargo Designado está pintado; se esperan los roles del cargo designado."
                if temporal.get("fecha_inicio") or fecha_fin:
                    motivo_temporal += (f" Vigencia informada: {temporal.get('fecha_inicio') or 'sin inicio'}"
                                        f" a {fecha_fin or 'sin fin'}.")
                motivo_temporal += (f" Roles: {resultado_roles}; cargo del sistema: {resultado_cargo};"
                                    f" oficina del sistema: {resultado_oficina}.")
                cargo_temporal_visible = " + ".join(
                    cargo for cargo in cargos_roles if cargo)
                detalle_temporal = _detalle(cuenta.login, nombre,
                    cargo_temporal_visible, sistema, encontrado_temporal,
                    resultado_roles)
                detalle_temporal["esperados"] = expected_temporal
                detalle.append(detalle_temporal)
                observaciones.append(_obs(categoria_temporal, cuenta.login, nombre=nombre,
                    nombre_sistema=cuenta.nombre,
                    cargo=cargo_designado, sistema=sistema, esperado=expected_temporal,
                    encontrado=encontrado_temporal,
                    cargo_actual=temporal["cargo_actual"],
                    cargo_original=persona.cargo if persona else "",
                    motivo=motivo_temporal, fecha_inicio=temporal.get("fecha_inicio", ""),
                    fecha_fin=fecha_fin,
                    cargo_designado=cargo_designado,
                    cargo_actual_seleccionado=bool(
                        es_fin_temporal or cargo_actual_pintado),
                    cargo_designado_seleccionado=bool(
                        not es_fin_temporal and (temporal.get("cargo_designado_pintado")
                            or not cargo_actual_pintado)),
                    oficina_actual=temporal.get("oficina_actual", ""),
                    oficina_designada=temporal.get("oficina_designada", ""),
                    cargo_sistema=cuenta.cargo, oficina_sistema=cuenta.oficina,
                    resultado_roles=resultado_roles,
                    resultado_cargo=resultado_cargo, resultado_oficina=resultado_oficina,
                    esperado_temporal_segmentos=segmentos_esperados,
                    origen_cargo=archivo_personas, origen_acceso=archivo_sistema))

    # Completa el nombre de la cuenta tal como aparece en el sistema.
    for observacion in observaciones:
        key = normalizar_login_sistema(observacion["login"], config)
        cuentas_usuario = sistema_por_login.get(key, [])
        observacion["nombre_sistema"] = " | ".join(dict.fromkeys(
            cuenta.nombre.strip() for cuenta in cuentas_usuario if cuenta.nombre.strip()))

    # Devuelve resultados y contadores de claves coincidentes o solo presentes en el sistema.
    return {"observaciones": observaciones, "errores": errores, "detalle": detalle,
            "coincidentes": sum(1 for key in sistema_por_login if key in dominio_por_login),
            "sin_coincidencia": sum(1 for key in sistema_por_login if key not in dominio_por_login)}


# Copia metadatos de procedencia y fuente a una observacion.
def _origen(base):
    # Copia a la observación únicamente los datos de procedencia y contexto.
    return {k: base[k] for k in (
        "origen_cargo", "origen_acceso", "nombre_sistema", "cargo_sistema", "login_salar",
        "oficina_sistema", "oficina_salar", "gerencia_sistema", "gerencia_salar")}


# Obtiene la clave de identidad configurada para una persona.
def _clave_persona(persona, config):
    # Si el sistema cruza por correo, permite una clave alternativa por login de respaldo.
    if config.clave_cruce == "correo":
        correo = normalizar_login(persona.correo)
        if correo:
            return f"correo:{correo}"
        if config.campo_login_respaldo:
            login = normalizar_login(persona.login)
            return f"login:{login}" if login else ""
        return ""
    return normalizar_login(persona.login)


# Obtiene la clave de cruce de una cuenta y considera logins de respaldo si aplica.
def _clave_cuenta(cuenta, config, dominio_por_login):
    # Prioriza el correo de la cuenta y usa login de respaldo cuando corresponde.
    if config.clave_cruce == "correo":
        correo = normalizar_login(cuenta.login)
        clave_correo = f"correo:{correo}" if correo else ""
        if clave_correo and clave_correo in dominio_por_login:
            return clave_correo
        login = normalizar_login(cuenta.login_respaldo)
        prefijo = (config.prefijo_login_respaldo or "").upper()
        if prefijo and login.startswith(prefijo):
            login = login[len(prefijo):]
        clave_login = f"login:{login}" if login else ""
        if clave_login and clave_login in dominio_por_login:
            return clave_login
        return clave_correo or clave_login
    return normalizar_login_sistema(cuenta.login, config)


# Busca una cuenta en Dominio usando la clave de cruce configurada para el sistema.
def _buscar_en_dominio(cuenta, personas_dominio, config):
    # Busca una identidad secundaria en Dominio con la clave configurada para el sistema.
    if not personas_dominio:
        return []
    if config.clave_cruce == "correo":
        correo = normalizar_login(cuenta.login)
        coincidencias = [p for p in personas_dominio
                         if correo and normalizar_login(p.correo) == correo]
        if coincidencias:
            return coincidencias
        login = normalizar_login(cuenta.login_respaldo)
        prefijo = (config.prefijo_login_respaldo or "").upper()
        if prefijo and login.startswith(prefijo):
            login = login[len(prefijo):]
    else:
        login = normalizar_login_sistema(cuenta.login, config)
    return [p for p in personas_dominio if login and normalizar_login(p.login) == login]


# Genera las identidades posibles de una cuenta para enlazarla con accesos temporales.
def _aliases_temporales(cuenta, config):
    # Devuelve identidades posibles de la cuenta para asociarla al Excel temporal.
    if config.clave_cruce != "correo":
        login = normalizar_login_sistema(cuenta.login, config)
        return {login} if login else set()
    aliases = set()
    correo = normalizar_login(cuenta.login)
    if correo:
        aliases.add(correo)
    login = normalizar_login(cuenta.login_respaldo)
    prefijo = (config.prefijo_login_respaldo or "").upper()
    if prefijo and login.startswith(prefijo):
        login = login[len(prefijo):]
    if login:
        aliases.add(login)
    return aliases


# Genera aliases de una persona de Dominio para reconciliarla con SALAR.
def _aliases_dominio(persona, config):
    # Genera claves equivalentes de Dominio para evitar alertas duplicadas de altas.
    if config.clave_cruce != "correo":
        login = normalizar_login(persona.login)
        return {login} if login else set()
    aliases = set()
    correo = normalizar_login(persona.correo)
    if correo:
        aliases.add(correo)
    login = normalizar_login(persona.login)
    prefijo = (config.prefijo_login_respaldo or "").upper()
    if prefijo and login.startswith(prefijo):
        login = login[len(prefijo):]
    if login:
        aliases.add(login)
    return aliases


def _variantes_alias_temporal(valor, config):
    """Genera variantes de login para asociar la planilla temporal a una persona."""
    raw = normalizar_login(valor)
    normalizado_sistema = normalizar_login_sistema(valor, config)
    aliases = {alias for alias in (raw, normalizado_sistema) if alias}
    if config.clave_cruce == "correo":
        prefijo = (config.prefijo_login_respaldo or "").upper()
        for alias in tuple(aliases):
            if prefijo and alias.startswith(prefijo):
                aliases.add(alias[len(prefijo):])
    return aliases


def _temporales_de_persona(persona, accesos_temporales, config):
    """Busca filas temporales por correo o login, aunque la cuenta del sistema falte."""
    if config.clave_cruce == "correo":
        aliases_persona = set(_aliases_dominio(persona, config))
        aliases_persona.update(_variantes_alias_temporal(persona.login, config))
        aliases_persona.update(_variantes_alias_temporal(persona.correo, config))
    else:
        aliases_persona = _variantes_alias_temporal(persona.login, config)
    if not aliases_persona:
        return []
    resultado = []
    for temporal in accesos_temporales or []:
        if aliases_persona.intersection(
                _variantes_alias_temporal(temporal.get("login", ""), config)):
            resultado.append(temporal)
    return resultado


def _expectativa_temporal_para_alta(persona, accesos_temporales,
                                    reglas_por_cargo, roles_matrix,
                                    problemas_matrix, acceso_esperado_matriz,
                                    config):
    """Calcula la expectativa de una persona sin cuenta usando sus cargos temporales."""
    temporales = _temporales_de_persona(persona, accesos_temporales, config)
    if not temporales:
        return None

    cargos_roles = []
    cargos_visibles = []
    for temporal in temporales:
        cargo_actual = (temporal.get("cargo_actual") or persona.cargo or "").strip()
        cargo_designado = (temporal.get("cargo_designado") or "").strip()
        fecha_fin = temporal.get("fecha_fin", "")
        fin_vigencia = date.fromisoformat(fecha_fin) if fecha_fin else None
        vencido = bool(fin_vigencia and date.today() > fin_vigencia)
        ambos_pintados = (temporal.get("cargo_actual_pintado")
                          and temporal.get("cargo_designado_pintado"))
        if vencido:
            seleccionados = [cargo_actual] if cargo_actual else []
            visible = cargo_actual
        elif ambos_pintados:
            seleccionados = [cargo for cargo in (cargo_actual, cargo_designado) if cargo]
            visible = "\n".join(part for part in (
                f"(Actual) {cargo_actual}" if cargo_actual else "",
                f"(Designado) {cargo_designado}" if cargo_designado else "") if part)
        elif temporal.get("cargo_actual_pintado"):
            seleccionados = [cargo_actual] if cargo_actual else []
            visible = cargo_actual
        else:
            seleccionados = [cargo_designado] if cargo_designado else []
            visible = cargo_designado
        cargos_roles.extend(seleccionados)
        if visible and visible not in cargos_visibles:
            cargos_visibles.append(visible)

    cargos_roles = list(dict.fromkeys(cargos_roles))
    cargo_visible = "\n".join(cargos_visibles)
    if not cargos_roles:
        return {"cargo": cargo_visible, "esperado": "",
                "tiene_acceso": False, "problema": "La fila temporal no contiene cargos evaluables."}

    claves = list(dict.fromkeys(normalizar_cargo(cargo) for cargo in cargos_roles))
    if any(key not in reglas_por_cargo for key in claves):
        return {"cargo": cargo_visible, "esperado": "",
                "tiene_acceso": False,
                "problema": "Uno de los cargos temporales no aparece en la matriz."}
    if any(key in problemas_matrix for key in claves):
        return {"cargo": cargo_visible, "esperado": "",
                "tiene_acceso": False,
                "problema": "Uno de los cargos temporales tiene una regla ambigua o inválida en la matriz."}

    if config.validar_solo_existencia:
        valores_esperados = list(dict.fromkeys(
            _mostrar_acceso_esperado(key, reglas_por_cargo, acceso_esperado_matriz)
            for key in claves if acceso_esperado_matriz.get(key, False)))
        return {"cargo": cargo_visible,
                "esperado": "; ".join(valores_esperados),
                "tiene_acceso": bool(valores_esperados), "problema": ""}

    esperados = set().union(*(roles_matrix[key] for key in claves))
    esperados_org = set().union(*(
        roles_matrix.get((key, "organigrama"), set()) for key in claves))
    esperado = _expected_access_text(_fmt(esperados), _fmt(esperados_org), config)
    return {"cargo": cargo_visible, "esperado": esperado,
            "tiene_acceso": bool(esperados or esperados_org), "problema": ""}


def _persona_dominio_tiene_cuenta(persona, cuentas, config):
    """Comprueba si Dominio tiene una cuenta correspondiente entre las cuentas cargadas."""
    login_persona = normalizar_login(persona.login)
    if config.clave_cruce == "correo":
        correo_persona = normalizar_login(persona.correo)
        prefijo = (config.prefijo_login_respaldo or "").upper()
        for cuenta in cuentas:
            if correo_persona and normalizar_login(cuenta.login) == correo_persona:
                return True
            login_respaldo = normalizar_login(cuenta.login_respaldo)
            if prefijo and login_respaldo.startswith(prefijo):
                login_respaldo = login_respaldo[len(prefijo):]
            if login_persona and login_respaldo == login_persona:
                return True
        return False
    return any(normalizar_login_sistema(cuenta.login, config) == login_persona
               for cuenta in cuentas if login_persona)


def _dominio_sin_persona_salar(personas_salar, personas_dominio):
    # Concilia por login/correo y usa nombre exacto solo para SALAR sin correo.
    """Encuentra personas solo en Dominio con el criterio de conciliaciÃ³n SALAR-Dominio."""
    usadas_salar, usadas_dominio = set(), set()

    # Empareja SALAR y Dominio con claves unicas y marca los casos ambiguos.
    def emparejar(clave_salar, clave_dominio):
        # Solo acepta una pareja cuando la clave es única a ambos lados.
        salar_por_clave, dominio_por_clave = defaultdict(list), defaultdict(list)
        for indice, persona in enumerate(personas_salar):
            clave = clave_salar(persona)
            if clave and indice not in usadas_salar:
                salar_por_clave[clave].append(indice)
        for indice, persona in enumerate(personas_dominio):
            clave = clave_dominio(persona)
            if clave and indice not in usadas_dominio:
                dominio_por_clave[clave].append(indice)
        for clave in salar_por_clave.keys() & dominio_por_clave.keys():
            filas_salar = salar_por_clave[clave]
            filas_dominio = dominio_por_clave[clave]
            if len(filas_salar) == len(filas_dominio) == 1:
                usadas_salar.add(filas_salar[0])
                usadas_dominio.add(filas_dominio[0])

    emparejar(lambda p: normalizar_login(p.login),
              lambda p: normalizar_login(p.login))
    emparejar(lambda p: normalizar_login(p.correo),
              lambda p: normalizar_login(p.correo))

    salar_por_nombre, dominio_por_nombre = defaultdict(list), defaultdict(list)
    for indice, persona in enumerate(personas_salar):
        nombre = normalizar_cargo(persona.nombre)
        if nombre and indice not in usadas_salar:
            salar_por_nombre[nombre].append(indice)
    for indice, persona in enumerate(personas_dominio):
        nombre = normalizar_cargo(persona.nombre)
        if nombre and indice not in usadas_dominio:
            dominio_por_nombre[nombre].append(indice)
    nombres_ambiguos_sin_correo = set()
    for nombre in salar_por_nombre.keys() & dominio_por_nombre.keys():
        filas_salar = salar_por_nombre[nombre]
        filas_dominio = dominio_por_nombre[nombre]
        if len(filas_salar) == len(filas_dominio) == 1:
            indice_salar, indice_dominio = filas_salar[0], filas_dominio[0]
            if not personas_salar[indice_salar].correo.strip():
                usadas_salar.add(indice_salar)
                usadas_dominio.add(indice_dominio)
        elif any(not personas_salar[indice].correo.strip() for indice in filas_salar):
            nombres_ambiguos_sin_correo.add(nombre)

    cuenta_login_dominio, cuenta_correo_dominio = defaultdict(int), defaultdict(int)
    for persona in personas_dominio:
        login, correo = normalizar_login(persona.login), normalizar_login(persona.correo)
        if login:
            cuenta_login_dominio[login] += 1
        if correo:
            cuenta_correo_dominio[correo] += 1
    resultado = []
    for indice, persona in enumerate(personas_dominio):
        if indice in usadas_dominio:
            continue
        nombre = normalizar_cargo(persona.nombre)
        if nombre in nombres_ambiguos_sin_correo:
            continue
        login, correo = normalizar_login(persona.login), normalizar_login(persona.correo)
        if ((login and cuenta_login_dominio[login] > 1)
                or (correo and cuenta_correo_dominio[correo] > 1)):
            continue
        resultado.append(persona)
    return resultado


# Determina si un valor de matriz representa un acceso esperado.
def _valor_indica_acceso(valor):
    # Interpreta si la celda de matriz expresa acceso o está vacía, Null o con guiones.
    if valor is None:
        return False
    texto = str(valor).strip().casefold()
    if texto in {"", "null"}:
        return False
    guiones = {"-", "‐", "‑", "‒", "–", "—", "−"}
    compactado = "".join(texto.split())
    return not compactado or not all(caracter in guiones for caracter in compactado)


def _mostrar_acceso_esperado(key, reglas_por_cargo, acceso_esperado_matriz):
    """Conserva el valor de matriz cuando expresa acceso; si no, devuelve celda vacía."""
    if not acceso_esperado_matriz.get(key, False):
        return ""
    # En sistemas validados por estado, se conserva la expectativa (por ejemplo,
    # "22"), aunque el TXT solo permita confirmar si la cuenta está Activa.
    for regla in reglas_por_cargo.get(key, []):
        valor = str(regla.roles_texto or "").strip()
        if _valor_indica_acceso(valor):
            return valor
    return "Con Acceso"


def _evaluar_cuenta_por_cargo(cargo, cuenta, reglas_por_cargo, roles_matrix,
                              problemas_matrix, acceso_esperado_matriz, config):
    # Evalúa una cuenta sin persona principal usando el cargo encontrado en Dominio.
    """Evalua una cuenta sin persona en la fuente principal usando su cargo de Dominio."""
    key = normalizar_cargo(cargo)
    encontrado = _access_text(cuenta.roles_texto, cuenta.organigrama_texto, config)
    if key not in reglas_por_cargo:
        return "", "Cargo no encontrado en matriz", encontrado
    if key in problemas_matrix:
        mensaje = f"No evaluable: {problemas_matrix[key]}"
        return mensaje, mensaje, encontrado
    if config.validar_solo_existencia:
        esperado_existe = acceso_esperado_matriz[key]
        activo = cuenta.estado.strip().casefold() == config.estado_activo.strip().casefold()
        esperado = _mostrar_acceso_esperado(key, reglas_por_cargo,
            acceso_esperado_matriz)
        if esperado_existe and not activo:
            resultado = "Estado distinto de Activo"
        elif not esperado_existe and activo:
            resultado = "Registro en sistema sin acceso esperado"
        else:
            resultado = "Conforme"
        return esperado, resultado, cuenta.estado or "(vacío)"
    encontrados, problema = _roles(cuenta.roles_texto, config.separador_roles,
        config.null_sin_roles, config.normalizar_codigos_rol)
    esperados = roles_matrix[key]
    encontrados_org, problema_org = (_roles(cuenta.organigrama_texto,
        config.separador_organigrama_txt) if config.campo_organigrama else (set(), None))
    esperados_org = roles_matrix.get((key, "organigrama"), set())
    esperado = _expected_access_text(_fmt(esperados), _fmt(esperados_org), config)
    if problema or problema_org:
        return esperado, "No evaluable: formato o significado de roles", encontrado
    if esperados == encontrados and esperados_org == encontrados_org:
        return esperado, "Conforme", encontrado
    return esperado, "Diferencia de roles", encontrado


# Compara nombre, cargo, oficina y gerencia que el sistema tenga configurados.
def _comparar_datos_persona(persona, cuenta, sistema, archivo_personas,
                            archivo_sistema, observaciones, config, fuente_personas):
    # Compara campos personales configurados y agrupa diferencias en una observación.
    campos = (
        ("cargo", persona.cargo, cuenta.cargo, normalizar_cargo),
        ("oficina", persona.oficina, cuenta.oficina, _normalizar_texto),
        ("gerencia", persona.gerencia, cuenta.gerencia, _normalizar_texto),
    )
    diferencias = [etiqueta for etiqueta, esperado, encontrado, normalizar in campos
                   if getattr(config, f"campo_{etiqueta}", None)
                   and (esperado or encontrado)
                   and normalizar(esperado) != normalizar(encontrado)]
    if diferencias:
        detalles = "; ".join(
            f"{etiqueta}: {fuente_personas}='{esperado.strip()}', sistema='{encontrado.strip()}'"
            for etiqueta, esperado, encontrado, _ in campos if etiqueta in diferencias)
        observaciones.append(_obs(
            "Diferencia de Oficina/Gerencia" if any(x in diferencias for x in ("oficina", "gerencia"))
            else "Diferencia de cargo",
            persona.login, nombre=persona.nombre, cargo=persona.cargo,
            nombre_sistema=cuenta.nombre, cargo_sistema=cuenta.cargo,
            oficina_sistema=cuenta.oficina, oficina_salar=persona.oficina,
            gerencia_sistema=cuenta.gerencia, gerencia_salar=persona.gerencia,
            sistema=sistema,
            motivo=f"Campos distintos respecto a {fuente_personas}: {detalles}.",
            origen_cargo=archivo_personas, origen_acceso=archivo_sistema))


# Normaliza texto organizacional para comparar ignorando acentos y formato.
def _normalizar_texto(valor):
    # Normaliza campos organizacionales para ignorar acentos y signos de puntuación.
    texto = (valor or "").casefold()
    texto = "".join(c for c in unicodedata.normalize("NFKD", texto)
                   if not unicodedata.combining(c))
    texto = "".join(" " if unicodedata.category(c).startswith("P") else c
                    for c in texto)
    partes = " ".join(texto.split()).split()
    if partes and partes[0] == "age":
        partes[0] = "agencia"
    return " ".join(partes)


# Estandariza una fila del detalle del analisis para el reporte.
def _detalle(login, nombre, cargo, sistema, encontrados, resultado):
    # Crea una fila estandarizada para la hoja de detalle del reporte.
    return {"login": login, "nombre": nombre, "cargo": cargo, "sistema": sistema,
            "esperados": "", "encontrados": encontrados, "coincidentes": "",
            "faltantes": "", "adicionales": "", "resultado": resultado}


# Ordena y convierte conjuntos de roles en texto.
def _fmt(valores: set[str]) -> str:
    # Ordena y une códigos para mostrarlos de forma estable en el Excel.
    return ", ".join(sorted(valores, key=lambda v: (len(v), v)))


# Arma el texto de accesos encontrados segun los campos que usa el sistema.
def _access_text(roles: str, organigrama: str, config: SistemaConfig) -> str:
    # Da formato conjunto a roles y organigrama, omitiendo campos que el sistema no usa.
    if not config.campo_organigrama:
        return roles or "Sin acceso"
    return f"Roles: {roles or 'Sin acceso'}; Organigrama: {organigrama or 'Sin acceso'}"


def _expected_access_text(roles: str, organigrama: str, config: SistemaConfig) -> str:
    """Deja vacía la expectativa cuando la matriz no asigna ningún acceso."""
    if not roles and not organigrama:
        return ""
    return _access_text(roles, organigrama, config)


# Consulta los accesos esperados de un cargo y devuelve errores de matriz asociados.
def _expected_for_cargo(cargo, reglas_por_cargo, roles_matrix, problemas_matrix, config):
    # Devuelve los accesos esperados o el motivo por el que el cargo no se puede evaluar.
    key = normalizar_cargo(cargo)
    if key not in reglas_por_cargo:
        return "Cargo no encontrado en matriz"
    if key in problemas_matrix:
        return f"No evaluable: {problemas_matrix[key]}"
    return _expected_access_text(_fmt(roles_matrix[key]),
                                 _fmt(roles_matrix.get((key, "organigrama"), set())), config)


# Evalua roles de una cuenta para un cargo temporal frente a la matriz.
def _estado_roles_temporal(cargo, cuenta, reglas_por_cargo, roles_matrix,
                           problemas_matrix, config):
    # Compara los roles de una cuenta temporal con la expectativa de un cargo.
    key = normalizar_cargo(cargo)
    if not cargo.strip():
        return "Cargo designado vacío"
    if key not in reglas_por_cargo:
        return "Cargo sin correspondencia en matriz"
    if key in problemas_matrix:
        return "No evaluable: matriz ambigua o inválida"
    encontrados, problema = _roles(cuenta.roles_texto, config.separador_roles,
                                   config.null_sin_roles, config.normalizar_codigos_rol)
    encontrados_org, problema_org = _roles(cuenta.organigrama_texto,
        config.separador_organigrama_txt) if config.campo_organigrama else (set(), None)
    if problema or problema_org:
        return "No evaluable: formato o significado de roles"
    esperados = roles_matrix[key]
    esperados_org = roles_matrix.get((key, "organigrama"), set())
    if (esperados == encontrados and esperados_org == encontrados_org):
        return "Conforme"
    return "Diferencia de roles"


# Combina accesos esperados de los cargos que cubre el temporal.
def _esperados_temporales(cargos, reglas_por_cargo, roles_matrix, problemas_matrix, config):
    # Calcula accesos esperados para uno o varios cargos de una asignación temporal.
    if len(cargos) == 1:
        return _expected_for_cargo(cargos[0], reglas_por_cargo, roles_matrix,
                                   problemas_matrix, config)
    partes = []
    for cargo in cargos:
        esperado = _expected_for_cargo(cargo, reglas_por_cargo,
            roles_matrix, problemas_matrix, config)
        if esperado:
            partes.append(f"{cargo}: {esperado}")
    return " + ".join(partes)


def _segmentos_esperados_temporales(cargos, reglas_por_cargo, roles_matrix,
                                    problemas_matrix, config):
    # Separa expectativas por cargo para colorearlas como segmentos en el reporte.
    """Texto y colores por origen para mostrar roles temporales en Excel."""
    segmentos = []
    for cargo in cargos:
        key = normalizar_cargo(cargo)
        if key not in reglas_por_cargo or key in problemas_matrix:
            if segmentos:
                segmentos.append((" + ", None))
            segmentos.append((_expected_for_cargo(cargo, reglas_por_cargo,
                roles_matrix, problemas_matrix, config), None))
            continue
        roles = _fmt(roles_matrix[key])
        if config.campo_organigrama:
            organigrama = _fmt(roles_matrix.get((key, "organigrama"), set()))
        else:
            organigrama = ""
        if not roles and not organigrama:
            continue
        if segmentos:
            segmentos.append((" + ", None))
        color = "1F4E79" if not any(color for _, color in segmentos) else "006100"
        if roles:
            segmentos.append((roles, color))
        if organigrama:
            texto_org = f" (Organigrama: {organigrama})" if roles else f"Organigrama: {organigrama}"
            segmentos.append((texto_org, None))
    return segmentos


# Compara los roles actuales con las expectativas de varios cargos temporales.
def _estado_roles_temporales(cargos, cuenta, reglas_por_cargo, roles_matrix,
                             problemas_matrix, config):
    # Un cargo designado pintado sustituye la expectativa normal. Si también
    # está pintado el actual, se compara contra la unión de ambos conjuntos.
    claves = [normalizar_cargo(cargo) for cargo in cargos]
    if any(not cargo.strip() for cargo in cargos):
        return "Cargo temporal vacío"
    for key in claves:
        if key not in reglas_por_cargo:
            return "Cargo sin correspondencia en matriz"
        if key in problemas_matrix:
            return "No evaluable: matriz ambigua o inválida"
    encontrados, problema = _roles(cuenta.roles_texto, config.separador_roles,
                                   config.null_sin_roles, config.normalizar_codigos_rol)
    encontrados_org, problema_org = _roles(cuenta.organigrama_texto,
        config.separador_organigrama_txt) if config.campo_organigrama else (set(), None)
    if problema or problema_org:
        return "No evaluable: formato o significado de roles"
    esperados = set().union(*(roles_matrix[key] for key in claves))
    esperados_org = set().union(*(roles_matrix.get((key, "organigrama"), set()) for key in claves))
    return ("Conforme" if esperados == encontrados and esperados_org == encontrados_org
            else "Diferencia de roles")


# Compara un campo temporal cuando su contenido puede requerir normalizacion.
def _estado_campo_temporal(esperado, encontrado, normalizar):
    # Compara un campo temporal aplicando el normalizador apropiado.
    if not esperado:
        return "Dato designado vacío"
    if not encontrado:
        return "Dato del sistema vacío"
    return "Conforme" if normalizar(esperado) == normalizar(encontrado) else "Diferencia"


# Decide si una cuenta tiene un acceso activo segun su configuracion.
def _tiene_acceso(cuenta: CuentaSistema, config: SistemaConfig) -> bool:
    # Determina si el TXT muestra algún acceso según las reglas particulares del sistema.
    valores = [cuenta.roles_texto]
    if config.campo_organigrama:
        valores.append(cuenta.organigrama_texto)
    return any(valor.strip() and valor.strip() != "-" and
               not (config.null_sin_roles and valor.strip().casefold() == "null")
               for valor in valores)


# Estandariza un error de carga con archivo, fila, campo y descripcion.
def _error(archivo, fila, campo, descripcion):
    # Da forma común a errores para reunirlos en una sola hoja del reporte.
    return {"archivo": archivo, "fila": fila, "campo": campo, "descripcion": descripcion}


# Estandariza una observacion de acceso con su contexto y valores comparados.
def _obs(categoria, login, nombre="", cargo="", sistema="", esperado="", encontrado="",
         faltantes="", adicionales="", motivo="", origen_cargo="", origen_acceso="",
         cargo_actual="", cargo_original="", nombre_sistema="", fecha_inicio="",
         fecha_fin="", cargo_designado="", oficina_actual="", oficina_designada="",
         cargo_sistema="", oficina_sistema="", resultado_roles="", resultado_cargo="",
         resultado_oficina="", oficina_salar="", gerencia_sistema="", gerencia_salar="",
         esperado_temporal_segmentos=None, login_salar="", cargo_en_matriz=None,
         cargo_actual_seleccionado=False, cargo_designado_seleccionado=False):
    # Estandariza observaciones con campos comunes y datos opcionales por categoría.
    return {"categoria": categoria, "login": login, "nombre": nombre, "cargo": cargo,
            "nombre_sistema": nombre_sistema, "login_salar": login_salar,
            "sistema": sistema, "esperado": esperado, "encontrado": encontrado,
            "faltantes": faltantes, "adicionales": adicionales, "motivo": motivo,
            "origen_cargo": origen_cargo, "origen_acceso": origen_acceso,
            "cargo_actual": cargo_actual,
            "cargo_original": cargo_original, "fecha_inicio": fecha_inicio,
            "fecha_fin": fecha_fin, "cargo_designado": cargo_designado,
            "oficina_actual": oficina_actual, "oficina_designada": oficina_designada,
            "cargo_sistema": cargo_sistema, "oficina_sistema": oficina_sistema,
            "resultado_roles": resultado_roles, "resultado_cargo": resultado_cargo,
            "resultado_oficina": resultado_oficina,
            "esperado_temporal_segmentos": esperado_temporal_segmentos,
            "cargo_actual_seleccionado": cargo_actual_seleccionado,
            "cargo_designado_seleccionado": cargo_designado_seleccionado,
            "oficina_salar": oficina_salar, "gerencia_sistema": gerencia_sistema,
            "gerencia_salar": gerencia_salar, "cargo_en_matriz": cargo_en_matriz}
