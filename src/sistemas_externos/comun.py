"""Lectura y comparación compartidas por los sistemas externos con roles."""
from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import ClassVar

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

from src.sistemas_internos.lectores import (DOMINIO_CONFIG, SALAR_CONFIG,
    PersonasConfig, TemporalesConfig, ErrorCarga, _leer_txt, leer_accesos_temporales,
    leer_personas_txt, leer_salar, normalizar_cargo, normalizar_login)


SHEET_MATRIX = "Matriz de Accesos"
UNMATCHED_ROLE_CATEGORY = "Rol - código no coincidentes"


# Define las posiciones y reglas de lectura de un sistema externo basado en roles.
@dataclass(frozen=True)
class SistemaExternoConfig:
    campos_txt: int = 3
    campo_correo: int = 1
    campo_nombre: int | None = 2
    campo_rol: int = 3
    hoja_codigos: str = "Codigos"
    seccion_codigos: str = "SISTEMAS ASFI"
    columna_seccion_codigos: str = "A"
    columna_rol_matriz: str = "C"
    columna_codigo_txt: str = "D"
    nombre_sistema: str = "ASFI"
    campo_estado: int | None = None
    estado_inactivo: str = ""
    columna_matriz: str = "H"
    equivalencias_roles: ClassVar[dict[str, str]] = {}
    nombres_roles: ClassVar[dict[str, str]] = {}
    aliases_seccion_codigos: ClassVar[tuple[str, ...]] = ()


# Agrupa identidad, roles TXT y datos personales consolidados de una cuenta.
@dataclass
class UsuarioExterno:
    correo: str
    nombre: str = ""
    roles: list[str] = field(default_factory=list)
    filas: list[int] = field(default_factory=list)


# Limpia espacios y convierte texto a una forma uniforme para comparar.
def _texto_normalizado(value) -> str:
    text = " ".join(str(value or "").split()).casefold()
    return "".join(char for char in unicodedata.normalize("NFKD", text)
                   if not unicodedata.combining(char))


# Normaliza un codigo de rol preservando los separadores con significado.
def _clave_rol(value) -> str:
    # Quita espacios para equiparar, por ejemplo, SRF_OperadorEIF y
    # SRF_Operador EIF; conserva guiones y guiones bajos que distinguen códigos.
    return re.sub(r"\s+", "", _texto_normalizado(value))


def _clave_equivalencia(value) -> str:
    """Normaliza diferencias de separadores al buscar un alias TXT en Codigos."""
    return re.sub(r"[^a-z0-9]+", "", _texto_normalizado(value))


# Forma una clave tolerante a espacios para comparar roles TXT y matriz.
def _clave_comparacion_rol(value, config: SistemaExternoConfig) -> str:
    key = _clave_rol(value)
    return config.equivalencias_roles.get(key, key)


# Ordena roles para mostrarlos de forma estable en el reporte.
def _roles_mostrables(roles, config: SistemaExternoConfig) -> list[str]:
    return sorted({config.nombres_roles.get(_clave_rol(role), role) for role in roles})


def _diferencias_roles(expected: set[str], role_candidates: list[tuple[str, set[str]]],
                       config: SistemaExternoConfig):
    """Compara por código canónico y conserva los textos originales al reportar."""
    expected_by_key: dict[str, set[str]] = defaultdict(set)
    for role in expected:
        expected_by_key[_clave_comparacion_rol(role, config)].add(role)
    satisfied_keys = {
        _clave_comparacion_rol(candidate, config)
        for _, candidates in role_candidates
        for candidate in candidates
        if _clave_comparacion_rol(candidate, config) in expected_by_key
    }
    missing = set()
    for key in set(expected_by_key) - satisfied_keys:
        missing.update(expected_by_key[key])
    additional = {
        next(iter(candidates))
        for _, candidates in role_candidates
        if len(candidates) == 1
        and _clave_comparacion_rol(next(iter(candidates)), config) not in expected_by_key
    }
    return missing, additional


# Lee filas del TXT externo, filtra estados inactivos y agrupa roles por correo.
def _leer_usuarios_sistema_externo(path: str | Path,
                        config: SistemaExternoConfig = SistemaExternoConfig()) -> tuple[list[UsuarioExterno], list[dict]]:
    # Consolida las filas por correo después de excluir el estado configurado como inactivo.
    path = Path(path)
    grouped: dict[str, UsuarioExterno] = {}
    errores = []
    positions = [config.campo_correo, config.campo_rol]
    positions.extend(position for position in (config.campo_nombre, config.campo_estado)
                     if position is not None)
    if (config.campos_txt < 1 or any(position < 1 or position > config.campos_txt
                                     for position in positions)
            or (config.campo_estado is not None and not config.estado_inactivo.strip())):
        raise ErrorCarga(f"Configuración de {config.nombre_sistema} inválida: los campos "
                         "configurados deben estar dentro del TXT y el estado inactivo debe indicarse.")
    for row_number, row in _leer_txt(path, config.campos_txt):
        estado = (row[config.campo_estado - 1].strip()
                  if config.campo_estado is not None else "")
        if (config.campo_estado is not None
                and estado.casefold() == config.estado_inactivo.strip().casefold()):
            continue
        correo = row[config.campo_correo - 1].strip()
        nombre = (row[config.campo_nombre - 1].strip()
                  if config.campo_nombre is not None else "")
        rol = row[config.campo_rol - 1].strip()
        email_key = correo.casefold()
        if not email_key:
            errores.append({"archivo": path.name, "fila": row_number,
                "campo": "correo", "descripcion": "Correo vacío; no se puede consolidar el usuario."})
            continue
        user = grouped.setdefault(email_key, UsuarioExterno(correo=correo, nombre=nombre))
        if not user.nombre and nombre:
            user.nombre = nombre
        if rol and rol not in user.roles:
            user.roles.append(rol)
        user.filas.append(row_number)
    return list(grouped.values()), errores


# Carga equivalencias que traducen los roles externos del TXT a roles de matriz.
def _leer_codigos_sistema_externo(workbook, config: SistemaExternoConfig = SistemaExternoConfig()):
    # Lee la hoja de códigos para traducir los roles del TXT a roles de matriz.
    sheet = next((ws for ws in workbook.worksheets
                  if ws.title.strip().casefold() == config.hoja_codigos.strip().casefold()), None)
    if sheet is None:
        raise ErrorCarga(f"La matriz no contiene la hoja {config.hoja_codigos!r}.")

    try:
        section_column = column_index_from_string(config.columna_seccion_codigos.strip().upper())
        matrix_role_column = column_index_from_string(config.columna_rol_matriz.strip().upper())
        txt_code_column = column_index_from_string(config.columna_codigo_txt.strip().upper())
    except ValueError as exc:
        raise ErrorCarga("Configuración de columnas de Codigos inválida; usa letras como A, C o AA.") from exc

    current_group = ""
    matrix_roles: dict[str, set[str]] = defaultdict(set)
    # Un rol TXT puede corresponder a varios roles matriz.
    txt_aliases: dict[str, list[str]] = defaultdict(list)
    max_column = max(section_column, matrix_role_column, txt_code_column)
    section_key = _texto_normalizado(config.seccion_codigos)
    section_keys = {section_key}
    # Los alias corrigen grafías conocidas del bloque predeterminado; si el
    # usuario configuró otra sección, se respeta únicamente esa elección.
    if section_key == _texto_normalizado(type(config).seccion_codigos):
        section_keys.update(_texto_normalizado(alias)
                            for alias in config.aliases_seccion_codigos)
    for row in sheet.iter_rows(min_col=1, max_col=max_column, values_only=True):
        section_value = row[section_column - 1]
        if section_value is not None and str(section_value).strip():
            current_group = _texto_normalizado(section_value)
        if current_group not in section_keys:
            continue
        matrix_role = str(row[matrix_role_column - 1] or "").strip()
        txt_code = str(row[txt_code_column - 1] or "").strip()
        if matrix_role:
            matrix_roles[_clave_rol(matrix_role)].add(matrix_role)
            alias_key = _clave_equivalencia(txt_code)
            if txt_code and matrix_role not in txt_aliases[alias_key]:
                txt_aliases[alias_key].append(matrix_role)

    if not matrix_roles:
        raise ErrorCarga(f"No se encontraron roles en el bloque {config.seccion_codigos} "
                         f"de la hoja {config.hoja_codigos}.")
    return matrix_roles, txt_aliases


# Resuelve un rol TXT a uno o varios codigos de la matriz.
def _traducir_rol(role, matrix_roles, txt_aliases):
    key = _clave_rol(role)
    direct = matrix_roles.get(key, set())
    if len(direct) == 1:
        return set(direct), ""
    if len(direct) > 1:
        return set(direct), ""

    matches = txt_aliases.get(_clave_equivalencia(role), [])
    if matches:
        return set(matches), ""
    return set(), f"No existe equivalencia entre los roles TXT y de matriz de {role!r}."


# Construye expectativas por cargo y detecta cargos repetidos o reglas inválidas.
def _leer_matriz_sistema_externo(workbook, config: SistemaExternoConfig = SistemaExternoConfig()):
    # Construye las expectativas de acceso por cargo para el sistema configurado.
    sheet = next((ws for ws in workbook.worksheets
                  if ws.title.strip().casefold() == SHEET_MATRIX.casefold()), None)
    if sheet is None:
        raise ErrorCarga(f"La matriz no contiene la hoja {SHEET_MATRIX!r}.")
    expected_by_cargo: dict[str, set[str]] = {}
    cargo_labels: dict[str, str] = {}
    duplicate_cargos = set()
    try:
        matrix_column = column_index_from_string(config.columna_matriz.strip().upper())
    except ValueError as exc:
        raise ErrorCarga(f"Columna de matriz inválida para {config.nombre_sistema}: "
                         f"{config.columna_matriz!r}.") from exc
    for row in sheet.iter_rows(min_row=7, min_col=2, max_col=matrix_column,
                               values_only=True):
        cargo = str(row[0] or "").strip()
        if not cargo:
            continue
        cargo_key = normalizar_cargo(cargo)
        expected_text = str(row[matrix_column - 2] or "").strip()
        roles = {part.strip() for part in expected_text.split(",")
                 if part.strip() and part.strip() not in {"-", "Null"}}
        if cargo_key in expected_by_cargo:
            duplicate_cargos.add(cargo_key)
            expected_by_cargo[cargo_key].update(roles)
        else:
            expected_by_cargo[cargo_key] = roles
            cargo_labels[cargo_key] = cargo
    if not expected_by_cargo:
        raise ErrorCarga(f"No se encontraron cargos y accesos de {config.nombre_sistema} "
                         "en la hoja Matriz de Accesos.")
    return expected_by_cargo, cargo_labels, duplicate_cargos


def _observacion_acceso_externo(categoria, user, person, esperado, encontrado,
                             grupo_informe, cargo_en_matriz, motivo,
                             config: SistemaExternoConfig, **extra):
    """Adapta un caso externo al formato común de observaciones de acceso."""
    expected_values = (_roles_mostrables(esperado, config)
                       if isinstance(esperado, (set, list, tuple)) else esperado)
    found_values = (_roles_mostrables(encontrado, config)
                    if isinstance(encontrado, (set, list, tuple)) else encontrado)
    return {
        "categoria": categoria,
        "correo": user.correo if user else "",
        "nombre": (person.nombre if person and person.nombre else
                   user.nombre if user else ""),
        "cargo": person.cargo if person else "",
        "esperado": expected_values or "",
        "encontrado": found_values or "",
        "grupo_informe": grupo_informe,
        "cargo_en_matriz": cargo_en_matriz,
        "motivo": motivo,
        **extra,
    }


def _expectativa_temporal_externa(temporal, cargo_base, expected_by_cargo,
                               duplicate_cargos, nombre_sistema: str = "ASFI"):
    """Obtiene los roles que corresponden al cargo temporal y el cargo visible."""
    cargo_actual = (temporal.get("cargo_actual") or cargo_base or "").strip()
    cargo_designado = (temporal.get("cargo_designado") or "").strip()
    fecha_fin = temporal.get("fecha_fin", "")
    vigencia_finalizada = bool(fecha_fin and date.today() > date.fromisoformat(fecha_fin))

    if vigencia_finalizada:
        cargos = [cargo_actual] if cargo_actual else []
        cargo_visible = cargo_actual
    elif (temporal.get("cargo_actual_pintado")
          and temporal.get("cargo_designado_pintado") and cargo_actual):
        cargos = [cargo_actual, cargo_designado]
        cargo_visible = f"(Actual) {cargo_actual}\n(Designado) {cargo_designado}"
    elif temporal.get("cargo_actual_pintado"):
        cargos = [cargo_actual] if cargo_actual else []
        cargo_visible = cargo_actual
    else:
        cargos = [cargo_designado] if cargo_designado else []
        cargo_visible = cargo_designado

    cargo_keys = list(dict.fromkeys(normalizar_cargo(cargo) for cargo in cargos if cargo))
    if not cargo_keys:
        return set(), cargo_visible, "La fila temporal no contiene un cargo evaluable."
    if any(key not in expected_by_cargo for key in cargo_keys):
        return set(), cargo_visible, f"El cargo temporal no tiene correspondencia en la matriz {nombre_sistema}."
    if any(key in duplicate_cargos for key in cargo_keys):
        return set(), cargo_visible, f"El cargo temporal aparece duplicado en la matriz {nombre_sistema}."
    expected = set().union(*(expected_by_cargo[key] for key in cargo_keys))
    return expected, cargo_visible, ""


def _seleccion_cargos_temporales(temporal, cargo_base):
    """Indica qué celdas de cargo se usaron, sin eliminar sus valores visibles."""
    actual = (temporal.get("cargo_actual") or cargo_base or "").strip()
    designado = (temporal.get("cargo_designado") or "").strip()
    fecha_fin = temporal.get("fecha_fin", "")
    vencido = bool(fecha_fin and date.today() > date.fromisoformat(fecha_fin))
    actual_pintado = temporal.get("cargo_actual_pintado")
    designado_pintado = temporal.get("cargo_designado_pintado")
    if vencido:
        return True, False
    elif actual_pintado and designado_pintado and actual:
        return True, True
    elif actual_pintado:
        return True, False
    return False, True


def analizar_sistema_externo(matriz: str | Path, salar: str | Path,
                  archivo_sistema: str | Path, dominio: str | Path | None = None,
                  archivo_temporales: str | Path | None = None,
                  config_temporales: TemporalesConfig = TemporalesConfig(),
                  config_dominio: PersonasConfig = DOMINIO_CONFIG,
                  config_sistema: SistemaExternoConfig = SistemaExternoConfig()) -> dict:
    """Compara un sistema externo por correo con SALAR, Dominio y su matriz.

    Dominio es opcional y permite reconocer usuarios externos que
    no aparecen en SALAR. La planilla temporal también es opcional; su columna
    Usuario debe contener el correo usado por el sistema.
    """
    nombre_sistema = config_sistema.nombre_sistema
    # Carga fuentes, configura el cruce por correo y prepara expectativas de la matriz.
    matriz, salar, archivo_sistema = Path(matriz), Path(salar), Path(archivo_sistema)
    usuarios_sistema, errores_sistema = _leer_usuarios_sistema_externo(archivo_sistema, config_sistema)
    people, errores_salar = leer_salar(salar, SALAR_CONFIG)
    domain_people, errores_domain = (leer_personas_txt(dominio, config_dominio)
                                     if dominio else ([], []))
    temporales = (leer_accesos_temporales(archivo_temporales, config_temporales)
                  if archivo_temporales else [])

    try:
        workbook = load_workbook(matriz, data_only=True, read_only=True)
    except Exception as exc:
        raise ErrorCarga(f"No se pudo leer la matriz {matriz}: {exc}") from exc
    try:
        matrix_roles, txt_aliases = _leer_codigos_sistema_externo(workbook, config_sistema)
        expected_by_cargo, cargo_labels, duplicate_cargos = _leer_matriz_sistema_externo(workbook, config_sistema)
    finally:
        workbook.close()

    # Crea índices por correo para enlazar las identidades de SALAR, Dominio y el sistema.
    people_by_email: dict[str, list] = defaultdict(list)
    for person in people:
        email_key = normalizar_login(person.correo)
        if email_key:
            people_by_email[email_key].append(person)

    domain_by_email: dict[str, list] = defaultdict(list)
    domain_by_name: dict[str, list] = defaultdict(list)
    for person in domain_people:
        if person.correo.strip():
            domain_by_email[normalizar_login(person.correo)].append(person)
        if person.nombre.strip():
            domain_by_name[normalizar_cargo(person.nombre)].append(person)

    # Al igual que el flujo SALAR interno, se puede reconocer una fila SALAR
    # sin correo si su nombre coincide de forma única con Dominio. El cruce
    # sigue usando correo; Dominio aporta el correo solo a esa fila SALAR.
    # Completa únicamente el correo SALAR vacío con una coincidencia única por nombre en Dominio.
    salar_sin_correo_por_nombre = defaultdict(list)
    for person in people:
        if not person.correo.strip() and person.nombre.strip():
            salar_sin_correo_por_nombre[normalizar_cargo(person.nombre)].append(person)
    for name_key, salar_rows in salar_sin_correo_por_nombre.items():
        domain_rows = domain_by_name.get(name_key, [])
        if len(salar_rows) == len(domain_rows) == 1 and domain_rows[0].correo.strip():
            inferred_email = domain_rows[0].correo.strip()
            salar_rows[0].correo = inferred_email
            people_by_email[normalizar_login(inferred_email)].append(salar_rows[0])
        elif salar_rows and not domain_rows:
            for person in salar_rows:
                errores_salar.append({"archivo": salar.name, "fila": person.fila,
                    "campo": "correo (19)",
                    "usuario_salar": person.login, "nombre_salar": person.nombre,
                    "cargo_salar": person.cargo, "correo_salar": person.correo,
                    "descripcion": "Correo vacío y sin coincidencia única por nombre en Dominio."})
        elif salar_rows and (len(domain_rows) != 1 or not domain_rows[0].correo.strip()):
            for person in salar_rows:
                errores_salar.append({"archivo": salar.name, "fila": person.fila,
                    "campo": "correo (19)",
                    "usuario_salar": person.login, "nombre_salar": person.nombre,
                    "cargo_salar": person.cargo, "correo_salar": person.correo,
                    "descripcion": ("La coincidencia por nombre en Dominio no es única o no tiene correo; "
                                    "no se asignó un correo automáticamente.")})

    errores = list(errores_sistema) + list(errores_salar) + list(errores_domain)
    email_aliases: dict[str, set[str]] = defaultdict(set)
    for person in people:
        email_key = normalizar_login(person.correo)
        if not email_key:
            continue
        email_aliases[email_key].add(email_key)
        if person.login.strip():
            email_aliases[normalizar_login(person.login)].add(email_key)
    for person in domain_people:
        email_key = normalizar_login(person.correo)
        if email_key and person.login.strip():
            email_aliases[email_key].add(email_key)
            email_aliases[normalizar_login(person.login)].add(email_key)
    for user in usuarios_sistema:
        email_key = normalizar_login(user.correo)
        email_aliases[email_key].add(email_key)

    # Resuelve el login de la hoja temporal a un correo; registra los enlaces ambiguos o ausentes.
    temporales_by_email: dict[str, list] = defaultdict(list)
    for temporal in temporales:
        alias = normalizar_login(temporal["login"])
        matches = email_aliases.get(alias, set())
        if not matches and "@" in alias:
            matches = {alias}
        if len(matches) == 1:
            temporales_by_email[next(iter(matches))].append(temporal)
        elif len(matches) > 1:
            errores.append({"archivo": Path(archivo_temporales).name,
                "fila": temporal["fila"], "campo": "Usuario",
                "descripcion": f"El usuario temporal {temporal['login']!r} coincide con varios correos; no se asignó a {nombre_sistema}."})
        else:
            errores.append({"archivo": Path(archivo_temporales).name,
                "fila": temporal["fila"], "campo": "Usuario",
                "descripcion": f"No se encontró correo en SALAR o Dominio para {temporal['login']!r}."})

    # Evalúa cada cuenta con la ruta SALAR → Dominio → matriz, usando sus equivalencias.
    salar_names_without_email = {normalizar_cargo(person.nombre) for person in people
        if not person.correo.strip() and normalizar_cargo(person.nombre)}
    cuentas_por_correo = {normalizar_login(user.correo): user for user in usuarios_sistema}
    observaciones, detalle = [], []
    observaciones_temporales = []
    for user in usuarios_sistema:
        email_key = normalizar_login(user.correo)
        role_candidates, unmapped = [], []
        for role in user.roles:
            candidates, reason = _traducir_rol(role, matrix_roles, txt_aliases)
            if reason:
                unmapped.append({"rol_txt": role, "motivo": reason})
            else:
                role_candidates.append((role, candidates))
        translated = sorted({candidate for _, candidates in role_candidates
                             for candidate in candidates})
        found_with_unmapped = translated + [item["rol_txt"] for item in unmapped]

        person_rows = people_by_email.get(email_key, [])
        domain_rows = domain_by_email.get(email_key, [])
        if len(person_rows) > 1:
            errores.append({"archivo": salar.name, "fila": ", ".join(
                str(person.fila) for person in person_rows), "campo": "correo (19)",
                "descripcion": f"El correo {user.correo!r} aparece en varias filas SALAR; no se comparó."})
            continue
        if not person_rows and len(domain_rows) > 1:
            errores.append({"archivo": Path(dominio).name if dominio else "Dominio",
                "fila": ", ".join(str(p.fila) for p in domain_rows),
                "campo": "correo", "descripcion":
                f"El correo {user.correo!r} aparece en varias filas Dominio; no se clasificó."})
            continue

        if (not person_rows and len(domain_rows) == 1
                and normalizar_cargo(domain_rows[0].nombre) in salar_names_without_email):
            errores.append({"archivo": salar.name, "fila": "", "campo": "correo (19)",
                "descripcion": (
                    f"El correo de {nombre_sistema} {user.correo!r} coincide con una persona de Dominio "
                    "cuyo nombre también aparece en SALAR sin correo; no se puede "
                    "determinar de forma segura si es la misma persona.")})
            continue

        # SALAR manda cuando hay coincidencia por correo. Si no, una coincidencia
        # única en Dominio permite consultar la matriz con el cargo de Dominio.
        person = person_rows[0] if person_rows else (
            domain_rows[0] if len(domain_rows) == 1 else None)
        grupo_informe = "salar" if person_rows else (
            "domain" if person is not None else "external")
        temporal_rows = temporales_by_email.get(email_key, [])
        if person is None:
            # No existe una identidad única ni en SALAR ni en Dominio.
            observation = _observacion_acceso_externo(
                "Usuario fuera de SALAR con acceso pendiente de retirar",
                user, None, "", found_with_unmapped,
                "both", False,
                f"El correo de {nombre_sistema} no se encontró de forma única en SALAR ni Dominio.",
                config=config_sistema)
            observaciones.append(observation)
            detalle.append({**observation, "resultado": "Usuario externo; identidad no encontrada"})
            continue

        cargo_key = normalizar_cargo(person.cargo)
        cargo_en_matriz = cargo_key in expected_by_cargo
        if not cargo_en_matriz and not temporal_rows:
            if grupo_informe == "salar":
                category = "Cargo sin correspondencia en matriz"
                reason = f"El cargo SALAR no aparece en la matriz de {nombre_sistema}."
            else:
                category = "Usuario Externo (Dominio)"
                reason = f"La persona no aparece en SALAR y su cargo no tiene regla en la matriz de {nombre_sistema}."
            observation = _observacion_acceso_externo(
                category, user, person, "", found_with_unmapped,
                "salar" if grupo_informe == "salar" else "external", False, reason,
                config=config_sistema)
            observaciones.append(observation)
            detalle.append({**observation, "resultado": "Cargo sin correspondencia en matriz"})
            continue

        if cargo_key in duplicate_cargos and not temporal_rows:
            expected_duplicate = expected_by_cargo[cargo_key]
            observation = _observacion_acceso_externo(
                "Duplicados o casos especiales", user, person, expected_duplicate,
                found_with_unmapped, grupo_informe, True,
                f"El cargo tiene varias filas en la matriz de {nombre_sistema}; no se comparó.",
                config=config_sistema)
            observaciones.append(observation)
            detalle.append({**observation,
                "resultado": "No evaluable: cargo duplicado en matriz"})
            continue

        expected = expected_by_cargo.get(cargo_key, set())
        if unmapped and not temporal_rows:
            mismatch = _observacion_acceso_externo(
                UNMATCHED_ROLE_CATEGORY, user, person, expected,
                found_with_unmapped, grupo_informe, True,
                "No se compararon los roles porque al menos uno no tiene código en Codigos.",
                config=config_sistema,
                roles_no_coincidentes=unmapped)
            observaciones.append(mismatch)
            detalle.append({**mismatch, "resultado": "No evaluable: rol sin código"})
            if temporal_rows:
                if len(temporal_rows) > 1:
                    errores.append({"archivo": Path(archivo_temporales).name,
                        "fila": ", ".join(t["fila"] for t in temporal_rows),
                        "campo": "Usuario (correo)",
                        "descripcion": f"El correo {user.correo!r} aparece varias veces en accesos temporales de {nombre_sistema}."})
                else:
                    temporal = temporal_rows[0]
                    actual_seleccionado, designado_seleccionado = (
                        _seleccion_cargos_temporales(temporal, person.cargo))
                    temporal_mismatch = {**mismatch,
                        "cargo_actual": temporal["cargo_actual"],
                        "cargo_designado": temporal["cargo_designado"],
                        "cargo_actual_seleccionado": actual_seleccionado,
                        "cargo_designado_seleccionado": designado_seleccionado,
                        "fecha_inicio": temporal.get("fecha_inicio", ""),
                        "fecha_fin": temporal.get("fecha_fin", ""),
                        "resultado": "No evaluable: rol sin código",
                        "motivo": mismatch["motivo"] + " El caso temporal requiere revisión manual."}
                    observaciones_temporales.append(temporal_mismatch)
            continue

        if temporal_rows:
            if len(temporal_rows) > 1:
                errores.append({"archivo": Path(archivo_temporales).name,
                    "fila": ", ".join(t["fila"] for t in temporal_rows),
                    "campo": "Usuario (correo)",
                    "descripcion": f"El correo {user.correo!r} aparece varias veces en accesos temporales de {nombre_sistema}."})
                detalle.append({"sistema": nombre_sistema, "correo": user.correo,
                    "nombre": person.nombre or user.nombre, "cargo": person.cargo,
                    "esperado": _roles_mostrables(expected, config_sistema), "encontrado": translated,
                    "resultado": "No evaluable: usuario temporal duplicado"})
                continue
            temporal = temporal_rows[0]
            vigencia_finalizada = bool(temporal.get("fecha_fin") and
                date.today() > date.fromisoformat(temporal["fecha_fin"]))
            categoria_temporal = ("Fin Acceso Temporal" if vigencia_finalizada
                                  else "Con Acceso Temporal")
            temporal_expected, cargo_temporal_visible, temporal_problem = (
                _expectativa_temporal_externa(
                    temporal, person.cargo, expected_by_cargo,
                    duplicate_cargos, nombre_sistema))
            actual_seleccionado, designado_seleccionado = (
                _seleccion_cargos_temporales(temporal, person.cargo))
            if unmapped:
                temporal_missing, temporal_additional = [], []
            else:
                temporal_missing, temporal_additional = _diferencias_roles(
                    temporal_expected, role_candidates, config_sistema)
            is_fin = categoria_temporal == "Fin Acceso Temporal"
            resultado_temporal = temporal_problem or (
                "No evaluable: rol sin cÃ³digo" if unmapped else
                "Diferencia de roles" if temporal_missing or temporal_additional else "Conforme")
            obs_temporal = {
                "categoria": categoria_temporal, "correo": user.correo,
                "nombre": person.nombre or user.nombre, "cargo": cargo_temporal_visible,
                "cargo_actual": temporal["cargo_actual"],
                "cargo_designado": temporal["cargo_designado"],
                "cargo_actual_seleccionado": actual_seleccionado,
                "cargo_designado_seleccionado": designado_seleccionado,
                "esperado": _roles_mostrables(temporal_expected, config_sistema),
                "encontrado": _roles_mostrables(found_with_unmapped, config_sistema),
                "faltantes": sorted(temporal_missing),
                "adicionales": sorted(temporal_additional),
                "fecha_inicio": temporal.get("fecha_inicio", ""),
                "fecha_fin": temporal.get("fecha_fin", ""),
                "resultado": resultado_temporal,
                "grupo_informe": grupo_informe,
                "cargo_en_matriz": True,
                "motivo": temporal_problem or (
                    "Vigencia finalizada; se esperan solo los roles del cargo de origen."
                    if is_fin else
                    "Acceso temporal evaluado con el cargo designado y las reglas de pintado de la planilla."),
            }
            # Conserva todos los temporales en el análisis; el informe final
            # omite los que el detalle confirma como conformes.
            observaciones_temporales.append(obs_temporal)
            detalle.append({**obs_temporal, "resultado": resultado_temporal})
            continue

        # Una cuenta con varios alias TXT puede satisfacer cualquier código
        # esperado, pero solo los roles de equivalencia única pueden ser adicionales.
        missing, additional = _diferencias_roles(expected, role_candidates, config_sistema)
        result = {
            "sistema": nombre_sistema, "correo": user.correo,
            "nombre": person.nombre or user.nombre, "cargo": person.cargo,
            "esperado": _roles_mostrables(expected, config_sistema),
            "encontrado": _roles_mostrables(translated, config_sistema),
            "faltantes": sorted(missing), "adicionales": sorted(additional),
            "resultado": "Diferencia de roles" if missing or additional else "Conforme",
        }
        detalle.append(result)
        if missing or additional:
            category = ("Diferencia de roles" if grupo_informe == "salar"
                        else "Usuario Externo (Dominio)")
            observaciones.append(_observacion_acceso_externo(
                category, user, person, expected, translated, grupo_informe, True,
                f"Los roles de {nombre_sistema} encontrados difieren de los que espera la matriz.",
                config=config_sistema,
                faltantes=sorted(missing), adicionales=sorted(additional)))

        # Busca altas SALAR, usando correo como clave y la matriz por cargo.
    emails_salar_con_error = set()
    for person in people:
        email_key = normalizar_login(person.correo)
        if not email_key or email_key in cuentas_por_correo:
            continue
        person_rows = people_by_email.get(email_key, [])
        if len(person_rows) > 1:
            if email_key not in emails_salar_con_error:
                errores.append({"archivo": salar.name,
                    "fila": ", ".join(str(p.fila) for p in person_rows),
                    "campo": "correo (19)",
                    "descripcion": f"El correo {person.correo!r} aparece en varias filas SALAR; no se generó alta."})
                emails_salar_con_error.add(email_key)
            continue
        temporal_rows = temporales_by_email.get(email_key, [])
        cargo_visible = person.cargo
        motivo_temporal = ""
        if len(temporal_rows) > 1:
            errores.append({"archivo": Path(archivo_temporales).name,
                "fila": ", ".join(item["fila"] for item in temporal_rows),
                "campo": "Usuario",
                "descripcion": (f"El correo {person.correo!r} tiene varias filas temporales; "
                                "no se generó alta automática con un cargo elegido al azar.")})
            detalle.append({"sistema": nombre_sistema, "correo": person.correo,
                "nombre": person.nombre, "cargo": person.cargo,
                "esperado": "", "encontrado": "",
                "resultado": "No evaluable: usuario temporal duplicado"})
            continue
        if temporal_rows:
            expected, cargo_visible, temporal_problem = _expectativa_temporal_externa(
                temporal_rows[0], person.cargo, expected_by_cargo, duplicate_cargos,
                nombre_sistema)
            if temporal_problem:
                errores.append({"archivo": Path(archivo_temporales).name,
                    "fila": temporal_rows[0]["fila"], "campo": "Cargo Designado",
                    "descripcion": (f"No se pudo evaluar el acceso temporal de "
                                    f"{person.correo!r}: {temporal_problem}")})
                detalle.append({"sistema": nombre_sistema, "correo": person.correo,
                    "nombre": person.nombre, "cargo": cargo_visible,
                    "esperado": "", "encontrado": "",
                    "resultado": f"No evaluable: {temporal_problem}"})
                continue
            motivo_temporal = (
                "Cargo temporal aplicado al alta; se consultó la matriz con la "
                "asignación vigente y las reglas de celdas pintadas.")
        else:
            cargo_key = normalizar_cargo(person.cargo)
            if cargo_key not in expected_by_cargo:
                observation = {
                    "categoria": "Cargo sin correspondencia en matriz",
                    "correo": person.correo, "nombre": person.nombre,
                    "cargo": person.cargo, "esperado": "", "encontrado": "",
                    "grupo_informe": "salar", "cargo_en_matriz": False,
                    "motivo": f"El cargo SALAR no aparece en la matriz de {nombre_sistema}.",
                }
                observaciones.append(observation)
                detalle.append({**observation,
                    "resultado": f"Sin cuenta {nombre_sistema}; cargo sin correspondencia en matriz"})
                continue
            if cargo_key in duplicate_cargos:
                errores.append({"archivo": matriz.name, "fila": "", "campo": f"cargo (matriz {nombre_sistema})",
                    "descripcion": f"No se generó alta para {person.correo!r}: el cargo {person.cargo!r} está duplicado en la matriz."})
                detalle.append({"sistema": nombre_sistema, "correo": person.correo,
                    "nombre": person.nombre, "cargo": person.cargo,
                    "esperado": _roles_mostrables(expected_by_cargo[cargo_key], config_sistema),
                    "encontrado": "", "resultado": "No evaluable: cargo duplicado en matriz"})
                continue
            expected = expected_by_cargo[cargo_key]
        if expected:
            observation = {
                "categoria": "Usuarios en SALAR sin registro en sistema",
                "correo": person.correo, "nombre": person.nombre,
                "cargo": cargo_visible,
                "esperado": _roles_mostrables(expected, config_sistema),
                "encontrado": "Sin registro en sistema",
                "grupo_informe": "salar", "cargo_en_matriz": True,
                "motivo": (motivo_temporal or
                    f"La matriz espera roles de {nombre_sistema}, pero el correo no aparece en el TXT."),
            }
            observaciones.append(observation)
            detalle.append({**observation, "resultado": "Alta pendiente"})
        else:
            detalle.append({"sistema": nombre_sistema, "correo": person.correo,
                "nombre": person.nombre, "cargo": cargo_visible,
                "esperado": "", "encontrado": "",
                "resultado": "No se espera acceso según matriz"})

    # Busca personas solo en Dominio sin cuenta del sistema y con acceso esperado.
    salar_logins = {normalizar_login(person.login) for person in people
                    if normalizar_login(person.login)}
    domain_emails_with_error = set()
    for domain_person in domain_people:
        email_key = normalizar_login(domain_person.correo)
        if not email_key:
            cargo_key = normalizar_cargo(domain_person.cargo)
            if cargo_key in expected_by_cargo and expected_by_cargo[cargo_key]:
                errores.append({"archivo": Path(dominio).name if dominio else "Dominio",
                    "fila": domain_person.fila, "campo": "correo",
                    "descripcion": (f"El cargo {domain_person.cargo!r} espera acceso en {nombre_sistema}, "
                                    "pero Dominio no proporciona correo para buscar la cuenta.")})
            continue
        if email_key in people_by_email or normalizar_login(domain_person.login) in salar_logins:
            continue
        if normalizar_cargo(domain_person.nombre) in salar_names_without_email:
            # Nombre candidato a SALAR sin correo: no crear una alta externa ambigua.
            continue
        if email_key in cuentas_por_correo:
            continue
        domain_rows = domain_by_email.get(email_key, [])
        if len(domain_rows) > 1:
            if email_key not in domain_emails_with_error:
                errores.append({"archivo": Path(dominio).name if dominio else "Dominio",
                    "fila": ", ".join(str(p.fila) for p in domain_rows),
                    "campo": "correo",
                    "descripcion": f"El correo {domain_person.correo!r} se repite en Dominio; no se generó alta."})
                domain_emails_with_error.add(email_key)
            continue
        temporal_rows = temporales_by_email.get(email_key, [])
        cargo_visible = domain_person.cargo
        motivo_temporal = ""
        if len(temporal_rows) > 1:
            errores.append({"archivo": Path(archivo_temporales).name,
                "fila": ", ".join(item["fila"] for item in temporal_rows),
                "campo": "Usuario",
                "descripcion": (f"El correo {domain_person.correo!r} tiene varias filas temporales; "
                                "no se generó alta automática con un cargo elegido al azar.")})
            continue
        if temporal_rows:
            expected, cargo_visible, temporal_problem = _expectativa_temporal_externa(
                temporal_rows[0], domain_person.cargo, expected_by_cargo, duplicate_cargos,
                nombre_sistema)
            if temporal_problem:
                errores.append({"archivo": Path(archivo_temporales).name,
                    "fila": temporal_rows[0]["fila"], "campo": "Cargo Designado",
                    "descripcion": (f"No se pudo evaluar el acceso temporal de "
                                    f"{domain_person.correo!r}: {temporal_problem}")})
                continue
            motivo_temporal = (
                "Cargo temporal aplicado al alta; se consultó la matriz con la "
                "asignación vigente y las reglas de celdas pintadas.")
        else:
            cargo_key = normalizar_cargo(domain_person.cargo)
            if cargo_key not in expected_by_cargo:
                continue
            if cargo_key in duplicate_cargos:
                errores.append({"archivo": matriz.name, "fila": "", "campo": f"cargo (matriz {nombre_sistema})",
                    "descripcion": f"No se generó alta para {domain_person.correo!r}: el cargo {domain_person.cargo!r} está duplicado en la matriz."})
                continue
            expected = expected_by_cargo[cargo_key]
        if not expected:
            continue
        observation = {
            "categoria": "Usuario en Dominio sin registro en sistema",
            "correo": domain_person.correo, "nombre": domain_person.nombre,
            "cargo": cargo_visible,
            "esperado": _roles_mostrables(expected, config_sistema),
            "encontrado": "Sin registro en sistema",
            "grupo_informe": "domain", "cargo_en_matriz": True,
            "motivo": (motivo_temporal or
                       "La persona aparece en Dominio, no está asociada a SALAR y "
                       f"la matriz espera roles de {nombre_sistema} para su cargo."),
        }
        observaciones.append(observation)
        detalle.append({**observation, "resultado": "Alta pendiente"})

    resultado = {"sistema": nombre_sistema, "observaciones": observaciones,
        "observaciones_temporales": observaciones_temporales, "detalle": detalle,
        "errores": errores, "usuarios": len(usuarios_sistema),
        "usuarios_con_detalle": len(detalle),
        "usuarios_salar": len(people), "usuarios_temporales": len(observaciones_temporales)}
    resultado[f"usuarios_{nombre_sistema.casefold()}"] = len(usuarios_sistema)
    return resultado

