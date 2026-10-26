"""Genera el formulario de informe a partir del Excel de revisión de accesos.

Este módulo es independiente de ``src/sistemas_internos``. Lee el reporte ya
generado y una plantilla editable, y escribe una copia con un bloque por sistema.
"""
from __future__ import annotations

import argparse
import re
import sys
import textwrap
import unicodedata
from zipfile import ZipFile
from collections import Counter
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE = ROOT / "documentos base" / "plantilla informe.xlsx"
DEFAULT_OUTPUT = ROOT / "data" / "salida" / "informe_accesos.xlsx"

# Hoja de observaciones del análisis -> tipo de observación de la plantilla.
# Los textos quedan centralizados aquí para ampliar o ajustar el mapeo sin tocar
# el analizador.
CATEGORY_TYPES = {
    "usuarios en salar sin registro en sistema": "ALTA",
    "usuario en salar sin registro en sistema": "ALTA",
    "usuario en dominio sin registro en sistema": "ALTA ESPECIAL",
    "acceso temporal": "ACCESO TEMPORAL",
    "con acceso temporal": "ACCESO TEMPORAL",
    "fin acceso temporal": "FIN ACCESO TEMPORAL",
    "usuario fuera de salar con acceso pendiente de retirar": "BAJA",
    "usuario fuera de dominio con acceso pendiente de retirar": "BAJA",
    "diferencia de roles": "ROL",
    "diferencia de oficina/gerencia": "OTRO",
    "diferencia por oficina/gerencia": "OTRO",
    "diferencia por oficina": "OTRO",
    "diferencia por gerencia": "OTRO",
    "diferencia de oficina": "OTRO",
    "diferencia de gerencia": "OTRO",
    "diferencia de cargo": "OTRO",
    "diferencia por cargo": "OTRO",
    "diferencia por nombre": "OTRO",
    "rol - codigo no coincidentes": "OTRO",
}
EXCLUDED_CATEGORIES = set()
GENERIC_CATEGORY_SHEETS = {
    "usuario externo": "Usuarios Externos",
    # Los usuarios externos se filtran por la marca de cargo en matriz.
    "usuario sin dominio pendiente de creacion": "Salar sin dominio",
    "cargo sin correspondencia en matriz": "Cargos Salar sin Matriz",
}

# Filas de la plantilla que ya tienen el color/estilo de cada tipo.
TYPE_TEMPLATE_ROW = {"ALTA": 10, "ALTA ESPECIAL": 10, "ACCESO TEMPORAL": 11, "BAJA": 12,
                     "FIN ACCESO TEMPORAL": 13, "ROL": 14, "OTRO": 15}
TYPE_LABEL = {"ALTA": "Altas", "ACCESO TEMPORAL": "Accesos Temporales",
              "ALTA ESPECIAL": "Altas",
              "BAJA": "Bajas", "FIN ACCESO TEMPORAL": "Fin Acceso Temporal",
              "OTRO": "Otros", "ROL": "Roles"}
TYPE_ORDER = {"ALTA": 0, "ALTA ESPECIAL": 0, "ACCESO TEMPORAL": 1, "BAJA": 2,
              "FIN ACCESO TEMPORAL": 3, "ROL": 4, "OTRO": 5}
SYSTEM_DISPLAY_NAMES = {
    "BDPSAM": "ArgosForm 2doPiso (BDPSAM)",
    "FG": "Fondo de Garantia (FG)",
}


# Normaliza encabezados y valores para comparar sin tildes ni diferencias de mayusculas.
def _norm(value) -> str:
    text = " ".join(str("" if value is None else value).split()).casefold()
    text = "".join(char for char in unicodedata.normalize("NFKD", text)
                   if not unicodedata.combining(char))
    return text


def _valor_acceso_visible(value, conservar_roles_sin_acceso=False,
                          conservar_null=False) -> str:
    """Muestra ausencia de accesos como celda vacía, incluso en campos compuestos."""
    text = str(value or "").strip()
    sin_acceso = {"", "sin acceso", "sin accesos", "null", "-"}
    if conservar_null and _norm(text) == "null":
        return text
    if _norm(text) in sin_acceso:
        return ""
    parts = []
    for part in text.split(";"):
        fragment = part.strip()
        label, separator, content = fragment.partition(":")
        if (conservar_roles_sin_acceso and separator
                and _norm(label) == "roles"
                and _norm(content) in {"sin acceso", "sin accesos"}):
            parts.append(fragment)
            continue
        if separator and _norm(content) in sin_acceso:
            continue
        if _norm(fragment) in sin_acceso:
            continue
        parts.append(fragment)
    return "; ".join(parts)


def _access_values_equal(expected, found) -> bool:
    """Compare access codes as sets, ignoring duplicates and order."""
    expected_text = str(expected or "").strip()
    found_text = str(found or "").strip()
    expected_codes = set(re.findall(r"\d+", expected_text))
    found_codes = set(re.findall(r"\d+", found_text))
    if expected_codes or found_codes:
        return expected_codes == found_codes

    # Tokeniza un texto de usuario para emparejar excepciones con mas tolerancia.
    def canonical_tokens(value):
        tokens = re.split(r"[+,;\n]+", _norm(value))
        normalized = set()
        for token in tokens:
            token = " ".join(token.split())
            if token in {"", "-", "null"}:
                normalized.add("sin acceso")
            elif token in {"con acceso", "activo", "activa"}:
                normalized.add("con acceso")
            elif token in {"sin acceso", "inactivo", "inactiva"}:
                normalized.add("sin acceso")
            else:
                normalized.add(token)
        return normalized

    return canonical_tokens(expected_text) == canonical_tokens(found_text)


def _punto_reclamo_activo_cubre_rol_esperado(expected, found) -> bool:
    """Punto de Reclamo expone estado; Activo representa el acceso esperado."""
    expected_codes = set(re.findall(r"\d+", str(expected or "")))
    return bool(expected_codes) and _norm(found) == "activo"


def _tipos_acceso_externo(expected, found) -> list[str]:
    """Deriva Altas/Roles de las diferencias entre matriz y cuenta de Dominio."""
    expected_text = str(expected or "").strip()
    found_text = str(found or "").strip()
    expected_normalized = _norm(expected_text)
    found_normalized = _norm(found_text)

    # Una matriz ambigua no permite recomendar acciones; cargo ausente o celda
    # vacía significa que no hay acceso esperado y se revisa si el sistema sí da acceso.
    if expected_normalized.startswith("no evaluable"):
        return []

    expected_codes = set(re.findall(r"\d+", expected_text))
    found_codes = set(re.findall(r"\d+", found_text))
    estado_activo = ("activo" in found_normalized
        and "inactivo" not in found_normalized)
    estado_inactivo = "inactivo" in found_normalized

    # Punto de Reclamo informa estado y no códigos de rol: Activo confirma la
    # existencia esperada, aunque la matriz conserve un código numérico.
    if estado_activo or estado_inactivo:
        espera_acceso = "con acceso" in expected_normalized
        espera_sin_acceso = "sin acceso" in expected_normalized
        if expected_codes or espera_acceso:
            return [] if estado_activo else ["ALTA"]
        if not expected_normalized or espera_sin_acceso or "cargo no encontrado" in expected_normalized:
            return ["ROL"] if estado_activo else []
        return []

    if expected_codes or found_codes:
        if expected_codes and not found_codes:
            return ["ALTA"]
        if expected_codes != found_codes:
            return ["ROL"]
        return []

    # Sistemas que validan existencia expresan la expectativa como Con/Sin Acceso
    # y el resultado como Activo/Inactivo, en lugar de listar códigos.
    espera_acceso = "con acceso" in expected_normalized
    espera_sin_acceso = ("sin acceso" in expected_normalized
                         or expected_normalized in {"-", "null"})
    tiene_acceso = ("con acceso" in found_normalized
                    or estado_activo)
    sin_acceso = (not found_normalized or "sin acceso" in found_normalized
                  or "inactivo" in found_normalized or "vacio" in found_normalized)
    if espera_acceso and sin_acceso:
        return ["ALTA"]
    if espera_sin_acceso and tiene_acceso:
        return ["ROL"]
    if not expected_normalized and found_normalized and not sin_acceso:
        return ["ROL"]
    return []


# Devuelve el valor de una celda solo si la columna existe en la hoja.
def _cell_value(row, index):
    return row[index] if index is not None and index < len(row) else None


# Obtiene la fecha del analisis desde la hoja de resumen.
def _date_from_summary(workbook) -> date:
    # Reutiliza la fecha del análisis para que el informe refleje cuándo se generó.
    if "Resumen" in workbook.sheetnames:
        sheet = workbook["Resumen"]
        for row in sheet.iter_rows(min_row=1, max_col=2, values_only=True):
            if _norm(row[0]) == "fecha y hora de ejecucion":
                value = row[1]
                if isinstance(value, datetime):
                    return value.date()
                if isinstance(value, date):
                    return value
                if value:
                    match = re.search(r"\d{4}-\d{2}-\d{2}", str(value))
                    if match:
                        return date.fromisoformat(match.group(0))
    return date.today()


def _verify_template_images_loaded(template_path: Path, workbook) -> None:
    """Avoid silently saving a template after its embedded images were dropped."""
    with ZipFile(template_path) as archive:
        template_has_images = any(name.startswith("xl/media/")
                                  for name in archive.namelist())
    images_loaded = any(getattr(sheet, "_images", []) for sheet in workbook.worksheets)
    if template_has_images and not images_loaded:
        workbook.close()
        if getattr(sys, "frozen", False):
            instruction = (
                "Vuelve a crear el ejecutable en Windows con 'python ejecutable.py' "
                "para incluir Pillow."
            )
        elif sys.platform == "win32":
            instruction = (
                "Instala las dependencias con "
                "'.venv\\Scripts\\python.exe -m pip install -r requirements.txt'."
            )
        else:
            instruction = (
                "Instala las dependencias con "
                "'.venv/bin/pip install -r requirements.txt'."
            )
        raise RuntimeError(
            "La plantilla contiene imágenes, pero no se pudieron cargar. "
            f"{instruction} Luego vuelve a generar el informe."
        )


# Convierte distintas hojas de analisis a filas comunes para el informe visual.
def _observation_rows(sheet, default_report_group="salar"):
    # Convierte hojas de observaciones a una estructura común, aunque cambien encabezados.
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, ())
    columns = {_norm(value): index for index, value in enumerate(header)
               if value is not None}

    # Encuentra el indice de la primera columna cuyo encabezado coincide con los alias dados.
    def index(*names):
        return next((columns[name] for name in names if name in columns), None)

    cat_i = index("categoria")
    if cat_i is None:
        return []
    login_i = index("login usado para el cruce", "login", "login (salar)",
                    "login (dominio)", "login salar", "correo", "correo asfi")
    email_i = index("correo", "correo electronico", "correo electronico (salar)",
                    "email", "e-mail", "correo asfi")
    name_i = index("nombre (salar)", "nombre (dominio)", "nombre", "nombre asfi")
    system_name_i = index("nombre en sistema")
    cargo_i = index("cargo (salar)", "cargo salar/dominio", "cargo salar",
                    "cargo (dominio)", "cargo")
    system_cargo_i = index("cargo (sistema)")
    office_source_i = index("oficina (salar)", "oficina (dominio)")
    office_system_i = index("oficina (sistema)")
    management_source_i = index("gerencia (salar)", "gerencia (dominio)")
    management_system_i = index("gerencia (sistema)")
    temp_actual_i = index("cargo actual temporal", "cargo (temporal, actual)",
                          "cargo de origen", "cargo actual")
    temp_designated_i = index("cargo designado temporal", "cargo (temporal, designado)",
                              "cargo designado")
    expected_i = index("valor esperado (matriz)", "valor esperado (matriz o salar)",
                       "valor esperado")
    found_i = index("valor encontrado (sistema)", "valor encontrado (asfi)",
                    "valor encontrado (infocred)",
                    "valor encontrado (dominio)",
                    "valor encontrado")
    matrix_cargo_i = index("cargo en matriz")
    report_group_i = index("grupo informe", "grupo del informe")
    result_i = index("resultado", "resultado de la comparacion")
    is_domain_comparison = "valor encontrado (dominio)" in columns
    system_name = sheet.title.removeprefix("Observaciones ").removeprefix("Temporal ")
    es_cia = _norm(system_name) == "cia"
    es_punto_reclamo = _norm(system_name) in {
        "punto reclamo", "punto de reclamo", "punto de reclamo (pr) bdp",
    }
    observations = []
    for row in rows:
        category = _cell_value(row, cat_i)
        if not category:
            continue
        category_text = str(category).strip()
        normalized_category = _norm(category_text)
        if normalized_category in EXCLUDED_CATEGORIES:
            continue
        expected = _cell_value(row, expected_i)
        found = _cell_value(row, found_i)
        expected = "" if expected is None else expected
        found = "" if found is None else found
        if (not is_domain_comparison
                and normalized_category in {"con acceso temporal", "acceso temporal",
                                            "fin acceso temporal"}):
            result = _norm(_cell_value(row, result_i))
            conforme = result == "conforme" if result_i is not None else (
                _access_values_equal(expected, found))
            if es_punto_reclamo:
                conforme = conforme or _punto_reclamo_activo_cubre_rol_esperado(
                    expected, found)
            if conforme:
                continue
        # Los usuarios externos conservan esa categoría en el análisis, pero sus
        # diferencias de acceso se muestran en Altas y/o Roles del informe general.
        if normalized_category == "usuario externo (dominio)":
            type_names = _tipos_acceso_externo(expected, found)
        else:
            type_name = CATEGORY_TYPES.get(normalized_category)
            type_names = [type_name] if type_name else []
        if not type_names:
            continue
        # En el informe general, el tipo de observación ya indica el alta pendiente;
        # deja vacía la celda de acceso encontrado si no existe una cuenta en el sistema.
        if _norm(found) == "sin registro en sistema":
            found = ""
        source_cargo = _cell_value(row, cargo_i) or ""
        system_cargo = _cell_value(row, system_cargo_i) or ""
        if normalized_category == "diferencia de cargo":
            expected, found = source_cargo, system_cargo
        elif normalized_category in {"diferencia de oficina/gerencia",
                                      "diferencia de oficina", "diferencia de gerencia"}:
            expected_parts, found_parts = [], []
            office_source = _cell_value(row, office_source_i) or ""
            office_system = _cell_value(row, office_system_i) or ""
            management_source = _cell_value(row, management_source_i) or ""
            management_system = _cell_value(row, management_system_i) or ""
            if _norm(office_source) != _norm(office_system):
                expected_parts.append(f"Oficina: {office_source}")
                found_parts.append(f"Oficina: {office_system}")
            if _norm(management_source) != _norm(management_system):
                expected_parts.append(f"Gerencia: {management_source}")
                found_parts.append(f"Gerencia: {management_system}")
            expected = "\n".join(expected_parts)
            found = "\n".join(found_parts)
        login = _cell_value(row, login_i)
        email = _cell_value(row, email_i)
        name = _cell_value(row, name_i) or _cell_value(row, system_name_i)
        cargo = source_cargo or system_cargo
        if normalized_category in {"con acceso temporal", "acceso temporal",
                                   "fin acceso temporal"}:
            actual_cargo = (_cell_value(row, temp_actual_i)
                            or system_cargo or source_cargo)
            designated_cargo = _cell_value(row, temp_designated_i)
            labeled_cargos = []
            if actual_cargo:
                labeled_cargos.append(f"(Actual) {actual_cargo}")
            if designated_cargo:
                labeled_cargos.append(f"(Designado) {designated_cargo}")
            if labeled_cargos:
                cargo = "\n".join(labeled_cargos)
        # Clasificar con los textos originales y luego vaciar las celdas que
        # representan ausencia de accesos.
        expected = _valor_acceso_visible(
            expected, conservar_roles_sin_acceso=es_cia)
        found = _valor_acceso_visible(found, conservar_roles_sin_acceso=es_cia,
                                      conservar_null=True)
        matrix_cargo = _cell_value(row, matrix_cargo_i)
        if isinstance(matrix_cargo, str):
            matrix_cargo = _norm(matrix_cargo) in {"si", "true", "1", "yes"}
        elif matrix_cargo is not None:
            matrix_cargo = bool(matrix_cargo)
        else:
            # Los archivos de análisis anteriores no incluían esta columna.
            matrix_cargo = bool(_valor_acceso_visible(
                _cell_value(row, expected_i)))
        explicit_group = _norm(_cell_value(row, report_group_i))
        valid_groups = {"salar", "domain", "dominio", "external", "externo",
                        "both", "ambos"}
        if explicit_group in valid_groups:
            report_group = ("domain" if explicit_group in {"domain", "dominio"}
                            else "both" if explicit_group in {"both", "ambos"}
                            else "external" if explicit_group in {"external", "externo"}
                            else "salar")
        else:
            report_group = default_report_group
            if normalized_category == "usuario en dominio sin registro en sistema":
                report_group = "domain"
            elif normalized_category == "usuario externo (dominio)":
                report_group = "domain" if matrix_cargo else "external"
            elif normalized_category == "usuario fuera de salar con acceso pendiente de retirar":
                # Una cuenta sin identidad en SALAR ni Dominio es relevante en ambas revisiones.
                report_group = "both"
            elif normalized_category == "usuario fuera de dominio con acceso pendiente de retirar":
                report_group = "external"
        for type_name in type_names:
            observations.append({
                "type": type_name, "category": category_text,
                "report_group": report_group,
                "login": login or "", "cargo": cargo or "",
                "expected": expected,
                "found": found,
                "name": name or "",
                "email": email or (login if "@" in str(login or "") else ""),
            })
    return observations


def _false_positive_key(value) -> str:
    """Normalize a system name or user identifier for exception matching."""
    return re.sub(r"[^a-z0-9@._+-]+", "", _norm(value))


# Unifica variantes de nombres para tipos de observacion de excepciones.
def _false_positive_type(value) -> str:
    normalized = _norm(value)
    aliases = {
        "alta": "alta", "altas": "alta",
        "alta especial": "alta", "altas especiales": "alta",
        "acceso temporal": "acceso temporal", "accesos temporales": "acceso temporal",
        "baja": "baja", "bajas": "baja",
        "fin acceso temporal": "fin acceso temporal",
        "fin accesos temporales": "fin acceso temporal",
        "rol": "rol", "roles": "rol",
        "otro": "otro", "otros": "otro",
    }
    return aliases.get(normalized, normalized)


# Forma la clave completa para identificar una excepcion aprobada.
def _false_positive_match_key(system, type_name, user, cargo, expected, found):
    return (
        _false_positive_key(system),
        _false_positive_type(type_name),
        _false_positive_key(user),
        _norm(cargo), _norm(expected), _norm(found),
    )


def _load_false_positives(path: str | Path | None) -> list[dict]:
    """Read detailed false-positive records from the optional workbook."""
    # Carga excepciones opcionales por sistema, tipo, usuario, cargo y valores comparados.
    if path is None:
        return []
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"No existe el archivo de falsos positivos: {path}")
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, ())
    columns = {_norm(value): index for index, value in enumerate(header)
               if value is not None}

    # Encuentra el indice de la primera columna cuyo encabezado coincide con los alias dados.
    def index(*names):
        return next((columns[name] for name in names if name in columns), None)

    positions = (
        index("sistema"), index("tipo"), index("usuario", "usuario/login", "login"),
        index("cargo/s", "cargos", "cargo"), index("valor esperado"),
        index("valor encontrado"), index("descripcion", "descripción"),
    )
    if any(position is None for position in positions[:6]):
        workbook.close()
        raise ValueError(
            "El Excel de falsos positivos debe incluir las columnas Sistema, Tipo, "
            "Usuario, Cargo/s, Valor esperado y Valor encontrado."
        )

    records_by_key = {}
    for row in rows:
        values = tuple(_cell_value(row, position) if position is not None else ""
                       for position in positions)
        system, type_name, user, cargo, expected, found, description = values
        if not _false_positive_key(system) or not _false_positive_type(type_name) \
                or not _false_positive_key(user):
            continue
        key = _false_positive_match_key(system, type_name, user, cargo, expected, found)
        records_by_key.setdefault(key, {
            "key": key,
            "values": (system, type_name, user, cargo, expected, found, description),
        })
    workbook.close()
    return list(records_by_key.values())


# Construye la clave de excepcion a partir de una observacion analizada.
def _observation_false_positive_key(system: str, item: dict):
    return _false_positive_match_key(
        system, TYPE_LABEL[item["type"]], item.get("login"), item.get("cargo"),
        item.get("expected"), item.get("found"),
    )


def _filter_false_positives(systems, false_positives):
    """Remove matching report observations and return only applied exceptions."""
    # Retira únicamente observaciones cuya clave coincide exactamente con una excepción.
    by_key = {record["key"]: record for record in false_positives}
    matched_keys = set()
    filtered_systems = []
    for system, observations in systems:
        remaining = []
        for item in observations:
            key = _observation_false_positive_key(system, item)
            if key in by_key:
                matched_keys.add(key)
            else:
                remaining.append(item)
        filtered_systems.append((system, remaining))
    applied = [record for record in false_positives if record["key"] in matched_keys]
    return filtered_systems, applied


# Actualiza la hoja que documenta falsos positivos/excepciones.
def _fill_false_positive_sheet(workbook, records):
    title = "Falsos Positivos"
    sheet = workbook[title] if title in workbook.sheetnames else workbook.create_sheet(title)
    headers = ("SISTEMA", "TIPO", "USUARIO", "CARGO/S", "VALOR ESPERADO",
               "VALOR ENCONTRADO", "DESCRIPCIÓN")
    widths = (22, 24, 28, 42, 30, 30, 44)
    for row in sheet.iter_rows(min_row=1, max_row=max(1, sheet.max_row),
                               min_col=1, max_col=len(headers)):
        for cell in row:
            cell.value = None
    for column, (header, width) in enumerate(zip(headers, widths), start=1):
        cell = sheet.cell(1, column, header)
        cell.fill = PatternFill(fill_type="solid", fgColor="305F91")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        sheet.column_dimensions[cell.column_letter].width = width
    sheet.row_dimensions[1].height = 30

    for row_number, record in enumerate(records, start=2):
        line_count = 1
        for column, value in enumerate(record["values"], start=1):
            cell = sheet.cell(row_number, column, value if value is not None else "")
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)
            line_count = max(line_count, _wrapped_line_count(
                value, sheet.column_dimensions[cell.column_letter].width))
        sheet.row_dimensions[row_number].height = max(22, line_count * 15 + 6)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:G{max(1, len(records) + 1)}"


def _generic_category_rows(sheet):
    """Read users and cargo categories destined for the generic list sheets."""
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, ())
    columns = {_norm(value): index for index, value in enumerate(header)
               if value is not None}

    # Encuentra el indice de la primera columna cuyo encabezado coincide con los alias dados.
    def index(*names):
        return next((columns[name] for name in names if name in columns), None)

    category_i = index("categoria")
    if category_i is None:
        return []
    matrix_cargo_i = index("cargo en matriz")
    login_i = index("login usado para el cruce", "login", "login (salar)",
                    "login (dominio)", "login salar")
    name_i = index("nombre (salar)", "nombre (dominio)", "nombre")
    email_i = index("correo", "correo (salar)", "correo (dominio)", "email")
    system_name_i = index("nombre en sistema")
    salar_cargo_i = index("cargo (salar)", "cargo (dominio)")
    system_cargo_i = index("cargo (sistema)")
    cargo_i = index("cargo")
    salar_office_i = index("oficina (salar)", "oficina (dominio)")
    system_office_i = index("oficina (sistema)")
    office_i = index("oficina")
    salar_management_i = index("gerencia (salar)", "gerencia (dominio)")
    system_management_i = index("gerencia (sistema)")
    management_i = index("gerencia")

    output = []
    for row in rows:
        category = _norm(_cell_value(row, category_i))
        if category not in GENERIC_CATEGORY_SHEETS:
            continue
        if category == "usuario externo":
            # Solo listar aqui personas de Dominio que no tengan cargo en matriz.
            if matrix_cargo_i is None or _norm(_cell_value(row, matrix_cargo_i)) != "no":
                continue
        sheet_name = GENERIC_CATEGORY_SHEETS[category]
        login = _cell_value(row, login_i)
        name = _cell_value(row, name_i) or _cell_value(row, system_name_i)
        email = _cell_value(row, email_i)
        cargo = (_cell_value(row, salar_cargo_i) or _cell_value(row, system_cargo_i)
                 or _cell_value(row, cargo_i))
        office = (_cell_value(row, salar_office_i) or _cell_value(row, system_office_i)
                  or _cell_value(row, office_i))
        management = (_cell_value(row, salar_management_i)
                      or _cell_value(row, system_management_i)
                      or _cell_value(row, management_i))
        if sheet_name == "Usuarios Externos":
            values = (login, name, email, cargo, office, management)
        else:
            values = (login, name, cargo, office, management)
        output.append((sheet_name, values))
    return output


def _error_rows(sheet):
    """Read internal and external loading errors using their shared headers."""
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, ())
    columns = {_norm(value): index for index, value in enumerate(header)
               if value is not None}

    # Encuentra el indice de la primera columna cuyo encabezado coincide con los alias dados.
    def index(*names):
        return next((columns[name] for name in names if name in columns), None)

    positions = (
        index("archivo"), index("numero de fila", "num fila", "fila"),
        index("campo afectado", "campo"), index("descripcion", "descripción"),
    )
    if any(position is None for position in positions):
        return []
    return [tuple(_cell_value(row, position) or "" for position in positions)
            for row in rows
            if any(_cell_value(row, position) not in (None, "")
                   for position in positions)]


def _fill_error_sheet(workbook, records):
    """Fill the template's loading-error tab with internal and external errors."""
    title = "Errores de carga"
    if title in workbook.sheetnames:
        sheet = workbook[title]
    else:
        sheet = workbook.create_sheet(title)
        sheet.append(["N.º", "ARCHIVO", "NUM FILA", "CAMPO AFECTADO", "DESCRIPCION"])

    prototype_styles = [copy(sheet.cell(2, col)._style) for col in range(1, 6)]
    for row in range(2, sheet.max_row + 1):
        for col in range(1, 6):
            sheet.cell(row, col).value = None

    for row_number, record in enumerate(records, start=2):
        values = (row_number - 1, *record)
        line_count = 1
        for col, value in enumerate(values, start=1):
            cell = sheet.cell(row_number, col)
            cell._style = copy(prototype_styles[col - 1])
            cell.value = value
            alignment = copy(cell.alignment)
            alignment.wrap_text = True
            alignment.vertical = "center"
            if col == 1:
                alignment.horizontal = "center"
            cell.alignment = alignment
            width = sheet.column_dimensions[cell.column_letter].width
            line_count = max(line_count, _wrapped_line_count(value, width))
        sheet.row_dimensions[row_number].height = max(24, line_count * 15 + 6)


def _fill_generic_sheets(workbook, records):
    # Rellena las secciones genéricas de la plantilla con pendientes y errores consolidados.
    """Fill each generic tab once per distinct record."""
    by_sheet = {sheet_name: [] for sheet_name in GENERIC_CATEGORY_SHEETS.values()}
    positions = {sheet_name: {} for sheet_name in by_sheet}
    for sheet_name, values in records:
        username_key = _norm(values[0])
        name_key = _norm(values[1])
        name_alias = ("nombre", name_key) if name_key else None
        user_alias = ("usuario", username_key) if username_key else None

        if sheet_name == "Salar sin dominio" and name_key:
            # Une pendientes repetidos por nombre y conserva el login SALAR disponible.
            key = name_alias
        elif sheet_name == "Cargos Salar sin Matriz":
            # Usa el login como identidad principal y el nombre como respaldo si falta login.
            key = user_alias or name_alias or (
                "registro", tuple(_norm(value) for value in values))
        elif username_key:
            key = user_alias
        else:
            # Sin usuario ni nombre, compara el registro completo.
            key = ("registro", tuple(_norm(value) for value in values))

        existing_index = positions[sheet_name].get(key)
        if (sheet_name == "Cargos Salar sin Matriz" and existing_index is None
                and name_alias is not None and name_alias in positions[sheet_name]):
            # Una fila sin login puede unirse por nombre si no hay otro login ambiguo.
            candidate = positions[sheet_name][name_alias]
            if candidate is not None:
                candidate_login = _norm(by_sheet[sheet_name][candidate][0])
                if not username_key or not candidate_login or candidate_login == username_key:
                    existing_index = candidate

        if existing_index is None:
            positions[sheet_name][key] = len(by_sheet[sheet_name])
            by_sheet[sheet_name].append(tuple(values))
            existing_index = len(by_sheet[sheet_name]) - 1
        else:
            existing = list(by_sheet[sheet_name][existing_index])
            # Completa campos vacios con datos de las filas repetidas.
            for index, value in enumerate(values):
                if existing[index] in (None, "") and value not in (None, ""):
                    existing[index] = value
            by_sheet[sheet_name][existing_index] = tuple(existing)

        if sheet_name == "Cargos Salar sin Matriz":
            # Registra aliases para unir las filas posteriores del mismo usuario.
            if user_alias is not None:
                positions[sheet_name][user_alias] = existing_index
            if name_alias is not None:
                if name_alias not in positions[sheet_name]:
                    positions[sheet_name][name_alias] = existing_index
                elif positions[sheet_name][name_alias] != existing_index:
                    # Si el nombre pertenece a varios logins, ya no es una clave segura.
                    positions[sheet_name][name_alias] = None

    for sheet_name, values_list in by_sheet.items():
        if sheet_name not in workbook.sheetnames and sheet_name == "Usuarios Fuera de Planilla":
            # Aceptar y corregir el nombre con errata que usó una versión de la plantilla.
            misspelled = next((name for name in workbook.sheetnames
                               if _norm(name) == "usuaros fuera de planilla"), None)
            if misspelled:
                workbook[misspelled].title = sheet_name
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        column_count = 7 if sheet_name == "Usuarios Externos" else 6
        prototype_styles = [copy(sheet.cell(2, col)._style)
                            for col in range(1, column_count + 1)]
        default_widths = {"A": 8, "B": 20, "C": 38, "D": 42,
                          "E": 30, "F": 38}
        if sheet_name == "Usuarios Externos":
            default_widths["G"] = 38
        for column, width in default_widths.items():
            if sheet.column_dimensions[column].width in (None, 13.0):
                sheet.column_dimensions[column].width = width
        for row in range(2, max(2, sheet.max_row) + 1):
            for col in range(1, column_count + 1):
                sheet.cell(row, col).value = None
        for row_offset, values in enumerate(values_list, start=2):
            line_count = 1
            numbered_values = (row_offset - 1, *values)
            for col, value in enumerate(numbered_values, start=1):
                cell = sheet.cell(row_offset, col)
                cell._style = copy(prototype_styles[col - 1])
                alignment = copy(cell.alignment)
                alignment.horizontal = "center"
                alignment.wrap_text = True
                cell.alignment = alignment
                cell.value = value
                width = sheet.column_dimensions[cell.column_letter].width
                line_count = max(line_count, _wrapped_line_count(value, width))
            sheet.row_dimensions[row_offset].height = max(24, line_count * 15 + 6)


# Copia el estilo de una celda para conservar el formato de la plantilla.
def _copy_style(source, destination):
    if source.has_style:
        destination._style = copy(source._style)
    if source.number_format:
        destination.number_format = source.number_format
    if source.alignment:
        destination.alignment = copy(source.alignment)


def _wrapped_line_count(value, width):
    """Estimate wrapped lines using the Excel column width in characters."""
    text = str(value or "")
    if not text:
        return 1
    line_width = max(1, int(width or 13))
    return sum(max(1, len(textwrap.wrap(
        line, width=line_width, break_long_words=True, break_on_hyphens=True
    ))) for line in text.splitlines())


# Calcula altura suficiente para que el contenido envuelto sea visible.
def _fit_observation_row_height(sheet, row):
    lines = 1
    for col in range(9, 14):
        cell = sheet.cell(row, col)
        if cell.alignment.wrap_text:
            width = sheet.column_dimensions[cell.column_letter].width
            lines = max(lines, _wrapped_line_count(cell.value, width))
    # A little padding keeps wrapped text away from the borders.
    sheet.row_dimensions[row].height = max(24, lines * 15 + 6)


# Organiza las observaciones en bloques por sistema y tipo dentro del reporte general.
def _make_system_blocks(sheet, systems):
    # Distribuye las observaciones en los bloques de sistema existentes en la plantilla.
    systems = [(SYSTEM_DISPLAY_NAMES.get(name, name), observations)
               for name, observations in systems]
    system_name_width = max((len(str(name).strip()) + 2 for name, _ in systems),
                            default=13)
    current_width = sheet.column_dimensions["A"].width or 13
    sheet.column_dimensions["A"].width = max(current_width,
                                               min(system_name_width, 24))

    # Guardar las seis filas de ejemplo (10 a 15) como prototipos de estilo.
    prototype_styles = {
        row: [copy(sheet.cell(row, col)._style) for col in range(1, 14)]
        for row in range(10, 16)
    }
    prototype_heights = {row: sheet.row_dimensions[row].height for row in range(10, 16)}

    if "A8:A9" in {str(merged) for merged in sheet.merged_cells.ranges}:
        sheet.unmerge_cells("A8:A9")
    for merged in list(sheet.merged_cells.ranges):
        if merged.min_row >= 10:
            sheet.unmerge_cells(str(merged))
    total_rows = sum(max(1, len(observations)) for _, observations in systems)
    if total_rows > 6:
        sheet.insert_rows(16, total_rows - 6)
    elif total_rows < 6:
        sheet.delete_rows(10 + total_rows, 6 - total_rows)

    # Limpiar las filas de ejemplo y preparar una fila por observación.
    for row in sheet.iter_rows(min_row=10, max_row=9 + total_rows,
                               min_col=1, max_col=13):
        for cell in row:
            cell.value = None

    current_row = 10
    for system, observations in systems:
        observations.sort(key=lambda item: TYPE_ORDER[item["type"]])
        block_size = max(1, len(observations))
        end_row = current_row + block_size - 1
        counts = Counter(item["type"] for item in observations)
        for offset in range(block_size):
            target_row = current_row + offset
            style_row = 10 + (offset % 6)
            if target_row > 15:
                sheet.row_dimensions[target_row].height = prototype_heights[style_row]
            for col in range(1, 14):
                sheet.cell(target_row, col)._style = copy(prototype_styles[style_row][col - 1])

            # El sistema se repite en cada fila para que el filtro de Excel
            # conserve todas las observaciones visibles del sistema elegido.
            sheet.cell(target_row, 1).value = system

            item = observations[offset] if offset < len(observations) else None
            if offset == 0:
                for col, type_name in ((2, "ALTA"), (3, "ACCESO TEMPORAL"),
                                       (4, "BAJA"), (5, "FIN ACCESO TEMPORAL"),
                                       (6, "ROL"), (7, "OTRO")):
                    if type_name == "ALTA":
                        count = counts.get("ALTA", 0) + counts.get("ALTA ESPECIAL", 0)
                    else:
                        count = counts.get(type_name, 0)
                    sheet.cell(target_row, col).value = count or "-"
                sheet.cell(target_row, 8).value = len(observations)
            if item:
                style_row = TYPE_TEMPLATE_ROW[item["type"]]
                for col in range(9, 14):
                    cell = sheet.cell(target_row, col)
                    cell._style = copy(prototype_styles[style_row][col - 1])
                sheet.cell(target_row, 9).value = TYPE_LABEL[item["type"]]
                sheet.cell(target_row, 10).value = item["login"]
                sheet.cell(target_row, 11).value = item["cargo"]
                sheet.cell(target_row, 12).value = item["expected"]
                sheet.cell(target_row, 13).value = item["found"]
                for col in range(9, 14):
                    cell = sheet.cell(target_row, col)
                    alignment = copy(cell.alignment)
                    alignment.wrap_text = True
                    alignment.shrink_to_fit = False
                    alignment.vertical = "center"
                    cell.alignment = alignment
                login_cell = sheet.cell(target_row, 10)
                login_alignment = copy(login_cell.alignment)
                login_alignment.horizontal = "center"
                login_alignment.vertical = "center"
                login_alignment.wrap_text = True
                login_cell.alignment = login_alignment
                _fit_observation_row_height(sheet, target_row)
            else:
                for col in range(9, 14):
                    sheet.cell(target_row, col).fill = PatternFill(fill_type=None)

        # El resumen se mantiene combinado por sistema; la columna SISTEMA
        # queda sin combinar para que el filtro tenga un valor en cada fila.
        if not observations:
            sheet.row_dimensions[current_row].height = max(
                sheet.row_dimensions[current_row].height or 0, 26)
        if end_row > current_row:
            for col in range(2, 9):
                sheet.merge_cells(start_row=current_row, start_column=col,
                                  end_row=end_row, end_column=col)
        current_row = end_row + 1
    sheet["A8"] = None
    sheet["A9"] = "SISTEMA"
    _copy_style(sheet["A8"], sheet["A9"])
    sheet["A9"].alignment = Alignment(horizontal="center", vertical="center",
                                      wrap_text=True)
    # Aplicar el filtro solo a SISTEMA: los demás encabezados de la plantilla
    # son agrupaciones y conteos, no columnas independientes de la observación.
    sheet.auto_filter.ref = f"A9:A{max(9, current_row - 1)}"
    # Mantener el filtro de SISTEMA sin inmovilizar filas ni columnas.
    sheet.freeze_panes = None
    sheet.print_area = f"A1:M{max(9, current_row - 1)}"


# Lee el analisis, filtra excepciones y llena las secciones de la plantilla final.
def generar_informe(analisis: str | Path, plantilla: str | Path = DEFAULT_TEMPLATE,
                    salida: str | Path = DEFAULT_OUTPUT,
                    falsos_positivos: str | Path | None = None) -> Path:
    # Lee el reporte de análisis, aplica excepciones y vuelca resultados en la plantilla final.
    analisis, plantilla, salida = Path(analisis), Path(plantilla), Path(salida)
    if not analisis.is_file():
        raise FileNotFoundError(f"No existe el reporte de análisis: {analisis}")
    if not plantilla.is_file():
        raise FileNotFoundError(f"No existe la plantilla: {plantilla}")

    workbook = load_workbook(analisis, data_only=True, read_only=True)
    systems = []
    generic_records = []
    error_records = []
    # Unifica hojas normales, temporales, conciliación de personas e incidencias de carga.
    for sheet_name in workbook.sheetnames:
        if sheet_name.startswith("Observaciones "):
            sheet = workbook[sheet_name]
            systems.append((sheet_name.removeprefix("Observaciones "),
                            _observation_rows(sheet)))
            generic_records.extend(_generic_category_rows(sheet))
        elif sheet_name == "Dominio Usuarios Temporales":
            sheet = workbook[sheet_name]
            observations = next((items for name, items in systems if name == "Dominio"), None)
            if observations is None:
                observations = []
                systems.append(("Dominio", observations))
            observations.extend(_observation_rows(sheet))
        elif sheet_name == "Historial SALAR-Dominio":
            sheet = workbook[sheet_name]
            rows = sheet.iter_rows(values_only=True)
            next(rows, ())
            for row in rows:
                if not str(_cell_value(row, 0) or "").startswith("Posible baja"):
                    continue
                domain_observations = next(
                    (items for name, items in systems if name == "Dominio"), None)
                if domain_observations is None:
                    domain_observations = []
                    systems.append(("Dominio", domain_observations))
                carnet = _cell_value(row, 1) or ""
                domain_observations.append({
                    "type": "BAJA",
                    "category": "Posible baja de SALAR; revisar cuenta habilitada",
                    "report_group": "external",
                    "login": _cell_value(row, 6) or _cell_value(row, 3) or "",
                    "cargo": _cell_value(row, 8) or _cell_value(row, 5) or "",
                    "expected": f"Carnet {carnet} ausente del SALAR actual",
                    "found": "Cuenta habilitada en Dominio",
                    "name": _cell_value(row, 2) or "",
                    "email": _cell_value(row, 7) or _cell_value(row, 4) or "",
                })
        elif sheet_name == "Comparación SALAR-Dominio":
            sheet = workbook[sheet_name]
            # Estos casos pertenecen al análisis SALAR-Dominio, que alimentará
            # La marca de matriz separa el listado externo del reporte Dominio-Matriz.
            systems.append(("Dominio", _observation_rows(
                sheet, default_report_group="external")))
            generic_records.extend(_generic_category_rows(sheet))
        elif sheet_name.startswith("Temporal "):
            system = sheet_name.removeprefix("Temporal ")
            observations = next((items for name, items in systems if name == system), None)
            if observations is None:
                observations = []
                systems.append((system, observations))
            observations.extend(_observation_rows(workbook[sheet_name]))
        elif sheet_name in {"Errores de carga", "Externos Errores de carga",
                            "INFOCRED Errores",
                            "Errores SALAR-Dominio"}:
            error_records.extend(_error_rows(workbook[sheet_name]))
    error_records = list(dict.fromkeys(tuple(str(value or "") for value in record)
                                       for record in error_records))
    report_date = _date_from_summary(workbook)
    workbook.close()

    # Filtra excepciones antes de crear el informe y conserva las aplicadas para auditoría.
    false_positive_records = _load_false_positives(falsos_positivos)
    applied_false_positives = []
    if false_positive_records:
        systems, applied_false_positives = _filter_false_positives(
            systems, false_positive_records)

    # Separa poblaciones: registros SALAR y registros solo de Dominio
    # cuyo cargo tiene una regla válida en la matriz.
    # Los casos `both` se incluyen en ambos informes para no perder Bajas sin identidad.
    systems_salar = [
        (name, [item for item in observations
                if item.get("report_group", "salar") in {"salar", "both"}])
        for name, observations in systems
    ]
    systems_domain = [
        (name, [item for item in observations
                if item.get("report_group") in {"domain", "both"}])
        for name, observations in systems
    ]

    # Copia los resultados a la plantilla, conserva sus estilos y guarda el archivo final.
    target = load_workbook(plantilla)
    _verify_template_images_loaded(plantilla, target)
    if "Observaciones por sistema" in target.sheetnames:
        del target["Observaciones por sistema"]
    _fill_generic_sheets(target, generic_records)
    _fill_error_sheet(target, error_records)
    if falsos_positivos is not None:
        _fill_false_positive_sheet(target, applied_false_positives)
    if {"Reporte Salar-Matriz", "Reporte Dominio-Matriz"}.issubset(target.sheetnames):
        report_sheets = [target["Reporte Salar-Matriz"],
                         target["Reporte Dominio-Matriz"]]
    else:
        # Compatibilidad con plantillas antiguas de una sola hoja.
        report_sheet = (target["Informe Revision Accesos"]
                        if "Informe Revision Accesos" in target.sheetnames else
                        target["Plantilla"] if "Plantilla" in target.sheetnames else
                        target.active)
        report_sheets = [report_sheet]
    # Nombres de meses en español sin depender del locale de Windows.
    months = ("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
              "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    for report_sheet in report_sheets:
        report_sheet["J5"] = (f"{report_date.day} DE {months[report_date.month - 1]} "
                              f"DE {report_date.year}")
    if len(report_sheets) == 2:
        _make_system_blocks(report_sheets[0], systems_salar)
        _make_system_blocks(report_sheets[1], systems_domain)
    else:
        _make_system_blocks(report_sheets[0], systems)
    # Las pestañas auxiliares son tablas con encabezados en la primera fila;
    # aplicar el filtro a todo el rango usado incluye cada una de sus columnas.
    # La hoja principal se conserva aparte porque contiene celdas combinadas y
    # un formato de impresión, no una tabla rectangular filtrable.
    for table_sheet in target.worksheets:
        if table_sheet in report_sheets or table_sheet.max_row < 1 or table_sheet.max_column < 1:
            continue
        last_column = get_column_letter(table_sheet.max_column)
        table_sheet.auto_filter.ref = f"A1:{last_column}{table_sheet.max_row}"
    active_sheet = report_sheets[0]
    for worksheet in target.worksheets:
        worksheet.sheet_view.tabSelected = worksheet is active_sheet
    target.active = target.worksheets.index(active_sheet)
    # La plantilla conserva el desplazamiento de pestañas del último guardado;
    # mostrar desde la primera hoja al abrir el informe generado.
    if target.views:
        target.views[0].firstSheet = 0
        target.views[0].activeTab = target.worksheets.index(active_sheet)
    salida.parent.mkdir(parents=True, exist_ok=True)
    target.save(salida)
    return salida


# Permite ejecutar la generacion del informe desde la linea de comandos.
def main():
    parser = argparse.ArgumentParser(
        description="Llena la plantilla de informe usando el Excel de revisión de accesos.")
    parser.add_argument("analisis", type=Path,
                        help="Excel generado por el análisis de accesos.")
    parser.add_argument("--plantilla", type=Path, default=DEFAULT_TEMPLATE,
                        help=f"Plantilla Excel (por defecto: {DEFAULT_TEMPLATE}).")
    parser.add_argument("--salida", type=Path, default=DEFAULT_OUTPUT,
                        help=f"Archivo de informe generado (por defecto: {DEFAULT_OUTPUT}).")
    parser.add_argument("--falsos-positivos", type=Path,
                        help="Excel opcional con los campos para filtrar falsos positivos.")
    args = parser.parse_args()
    print(generar_informe(args.analisis, args.plantilla, args.salida,
                          args.falsos_positivos))


if __name__ == "__main__":
    main()
