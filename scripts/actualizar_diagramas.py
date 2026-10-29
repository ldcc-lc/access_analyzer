"""Mantiene los ocho flujos SVG alineados con las rutas actuales del programa.

Se conservan los paneles, colores y recorridos numerados originales. El script
ajusta sus explicaciones y añade una banda final con la salida de cada flujo.
Es idempotente: se puede ejecutar de nuevo después de cambiar un diagrama.
"""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "data" / "salida"
NS = "http://www.w3.org/2000/svg"
ET.register_namespace("", NS)


def tag(name: str) -> str:
    return f"{{{NS}}}{name}"


def update_text(root: ET.Element, changes: dict[tuple[int, int], str]) -> None:
    by_position = {}
    for element in root.iter(tag("text")):
        if element.get("x") and element.get("y"):
            by_position[(int(element.get("x")), int(element.get("y")))] = element
    for position, value in changes.items():
        if position not in by_position:
            raise ValueError(f"No se encontró el texto en {position}.")
        by_position[position].text = value


def remove_path(root: ET.Element, shape: str) -> None:
    for parent in root.iter():
        for child in list(parent):
            if child.tag == tag("path") and child.get("d") == shape:
                parent.remove(child)
                return
    raise ValueError(f"No se encontró la flecha {shape}.")


def update_rect_style(root: ET.Element, x: int, y: int, class_name: str) -> None:
    matches = [element for element in root.iter(tag("rect"))
               if element.get("x") == str(x) and element.get("y") == str(y)]
    if len(matches) != 1:
        raise ValueError(f"No se encontró el recuadro único en {(x, y)}.")
    matches[0].set("class", class_name)


def add_output(root: ET.Element, y: int, lines: tuple[str, str]) -> None:
    footer = next((group for group in root.iter(tag("g"))
                   if group.get("id") == "salida-del-flujo"), None)
    if footer is None:
        footer = ET.SubElement(root, tag("g"), {"id": "salida-del-flujo"})
    else:
        footer.clear()
        footer.set("id", "salida-del-flujo")

    ET.SubElement(footer, tag("rect"), {
        "x": "8", "y": str(y), "width": "1784", "height": "88", "rx": "12",
        "fill": "#e8e3f5", "stroke": "#b3a5d5", "stroke-width": "1.5",
    })
    ET.SubElement(footer, tag("rect"), {
        "x": "8", "y": str(y), "width": "228", "height": "88", "rx": "12",
        "fill": "#4662a0",
    })
    ET.SubElement(footer, tag("text"), {
        "x": "35", "y": str(y + 42), "class": "lane-title",
    }).text = "SALIDA"
    for offset, line in ((33, lines[0]), (61, lines[1])):
        ET.SubElement(footer, tag("text"), {
            "x": "270", "y": str(y + offset),
            "style": "font-size:14px;fill:#344e6e",
        }).text = line

    height = y + 105
    root.set("height", str(height))
    root.set("viewBox", f"0 0 1800 {height}")
    background = next(element for element in root.iter(tag("rect"))
                      if element.get("x") is None and element.get("width") == "1800")
    background.set("height", str(height))


UPDATES = {
    "flujo_sistemas_internos_general.svg": {
        "desc": "Por cada sistema, lee las cuentas filtradas, compara con SALAR y Dominio por la clave configurada, consulta la matriz por cargo, revisa altas pendientes y trata accesos temporales.",
        "footer_y": 1250,
        "output": (
            "Excel: resumen, observaciones, detalle, errores y hoja Temporal con Actual/Designado conservados; verde marca el cargo evaluado.",
            "Correo «Null» en SALAR/Dominio equivale a vacío; con cargo de Dominio en matriz se compara el acceso y sin identidad se revisa el retiro.",
        ),
        "text": {
            (900, 69): "Cada sistema usa sus propios campos, clave de cruce, columna de matriz y reglas de estado",
            (215, 162): "cuenta · roles/datos · estado configurado",
            (550, 162): "identidad, cargo, oficina y gerencia",
            (885, 162): "identidad adicional si no hay SALAR",
            (1220, 162): "cargo → acceso, organigrama o estado",
            (1570, 162): "usuario · cargos · fechas · marcas",
            (35, 360): "Filtra filas según",
            (35, 380): "configuración del TXT",
            (35, 410): "Cruza por login/correo",
            (365, 308): "cuenta incluida por el lector",
            (815, 287): "¿Una persona",
            (815, 306): "única en SALAR?",
            (830, 361): "Ninguna",
            (1252, 308): "roles · organigrama · estado",
            (1565, 285): "Conforme o diferencia",
            (1565, 308): "roles, datos o estado",
            (1027, 448): "misma clave del sistema",
            (1650, 425): "Externo",
            (1650, 448): "(Dominio)",
            (1540, 531): "Sin Dominio o identidad repetida",
            (1540, 553): "sin persona → baja; repetida → revisión",
            (700, 531): "Cargo SALAR sin matriz o regla inválida",
            (700, 553): "Con acceso → observación; sin acceso → detalle. Regla ambigua → no evaluable.",
            (1610, 740): "Salida del sistema",
            (1610, 763): "observación + detalle",
            (1480, 995): "Alta: Dominio sin cuenta",
            (655, 1100): "Sin cuenta: el temporal cambia el cargo esperado; vencido usa el cargo de origen.",
            (300, 1180): "Cuenta encontrada: la observación ordinaria pasa a «Con Acceso Temporal» o «Fin Acceso Temporal».",
            (300, 1207): "No vencido: cargo designado, o ambos si están marcados; vencido: solo cargo actual.",
            (35, 1220): "* El filtro de estado se aplica solo en los sistemas que lo configuran; las filas excluidas no cuentan como acceso.",
        },
    },
    "flujo_analisis_asfi.svg": {
        "desc": "ASFI agrupa roles por correo, traduce los valores del TXT con Codigos D a C, consulta los accesos esperados en matriz H y contrasta cuentas y altas con SALAR y Dominio.",
        "footer_y": 1210,
        "output": (
            "Excel ASFI: Temporal conserva Actual y Designado; verde marca el cargo evaluado y el informe muestra la etiqueta seleccionada.",
            "Correo «Null» equivale a vacío: SALAR puede obtenerlo por nombre único en Dominio; sin correo de Dominio no se genera alta por correo.",
        ),
        "text": {
            (900, 69): "Roles TXT por correo → Codigos D a C → matriz H por cargo → comparación y altas",
            (232, 162): "correo 1 · nombre 2 · rol 3",
            (1185, 139): "Matriz ASFI + Codigos",
            (1185, 162): "H: esperado · sección ASFI · TXT D → matriz C",
            (35, 360): "Agrupa por correo",
            (35, 380): "traduce TXT D → C",
            (35, 400): "consulta matriz H",
            (370, 308): "correo + roles TXT agrupados",
            (1605, 281): "Si existe temporal",
            (1605, 304): "comparar cargo aplicable",
            (1605, 322): "según fecha fin y marcas",
            (1665, 447): "con matriz o externo",
            (1330, 526): "Sin persona o correo ambiguo",
            (1330, 548): "sin identidad → externo; repetido → error",
            (580, 535): "Sin matriz, duplicados o rol sin código",
            (580, 555): "Sin regla → categoría aparte; ambigüedad → no evaluable.",
            (1285, 823): "Sin roles esperados → no generar alta",
            (1630, 721): "Esperado temporal",
            (1630, 746): "cargo según fecha fin",
            (1475, 996): "ALTA · Dominio sin cuenta ASFI",
            (655, 1110): "Temporal: no vencido usa designado/ambos; vencido usa actual. Esta ruta consulta la fecha fin.",
            (900, 1190): "Sin cargo en matriz se clasifica aparte; roles no traducibles y correos repetidos requieren revisión antes de comparar.",
        },
    },
    "flujo_analisis_infocred.svg": {
        "desc": "INFOCRED excluye solo BAJA, agrupa roles por correo, traduce el TXT con Codigos B a C, consulta la matriz P y compara con SALAR, Dominio y accesos temporales.",
        "footer_y": 1210,
        "output": (
            "Excel INFOCRED: Temporal conserva Actual y Designado; verde marca el cargo evaluado y el informe muestra la etiqueta seleccionada.",
            "Solo BAJA se excluye; «Null» en correo equivale a vacío. Sin correo de Dominio no se busca alta; un rol sin código requiere revisión.",
        ),
        "text": {
            (900, 69): "Excluir BAJA → agrupar por correo → Codigos B a C → matriz P por cargo",
            (232, 139): "INFOCRED.txt · 17 campos",
            (232, 162): "correo 7 · rol 9 · estado 13",
            (1185, 139): "Matriz INFOCRED + Codigos",
            (1185, 162): "P: esperado · A=Infocred · TXT B → matriz C",
            (35, 360): "Excluye solo BAJA",
            (35, 324): "hay en INFOCRED",
            (35, 380): "agrupa roles por correo",
            (35, 400): "traduce TXT B → C",
            (370, 308): "correo + roles TXT agrupados",
            (1605, 281): "Si existe temporal",
            (1605, 304): "comparar cargo aplicable",
            (1605, 322): "según fecha fin y marcas",
            (1665, 447): "con matriz o externo",
            (1330, 526): "Sin persona o correo ambiguo",
            (1330, 548): "sin identidad → externo; repetido → error",
            (580, 535): "Sin matriz, duplicados o rol sin código",
            (580, 555): "Sin regla → categoría aparte; ambigüedad → no evaluable.",
            (1285, 823): "Sin roles esperados → no generar alta",
            (1630, 721): "Esperado temporal",
            (1630, 746): "cargo según fecha fin",
            (1475, 996): "ALTA · Dominio sin cuenta INFOCRED",
            (655, 1110): "Temporal: no vencido usa designado/ambos; vencido usa actual. Esta ruta consulta la fecha fin.",
            (900, 1190): "Sin cargo en matriz se clasifica aparte; roles no traducibles y correos repetidos requieren revisión antes de comparar.",
        },
    },
    "flujo_analisis_salar_dominio.svg": {
        "desc": "Concilia SALAR con Dominio por login y correo únicos, usa nombre único solo cuando SALAR no tiene correo, compara datos laborales y revisa temporales e historial opcionales.",
        "footer_y": 1255,
        "output": (
            "Excel: compara cargo/oficina SALAR-Dominio y aplica asignaciones temporales; incluye detalle e historial opcional.",
            "Las temporales generan Conforme/Diferencia de cargo u oficina; no comparan roles de sistemas contra la matriz.",
        ),
        "text": {
            (900, 69): "Conciliar login → correo → nombre único si falta correo SALAR; luego comparar datos laborales",
            (355, 162): "login · nombre · correo · cargo",
            (1510, 162): "carnet · login/correo · cargo anteriores",
            (680, 443): "No se asigna una pareja por aproximación; revisar manualmente.",
            (460, 684): "Detalle: también conserva login y correo de ambas fuentes",
            (1062, 684): "Datos laborales: SALAR · login visible: Dominio",
            (1502, 802): "marcar si el cargo está en matriz, si se cargó",
            (760, 839): "Login/correo también se comparan, aunque no creen observación propia",
            (760, 863): "Las diferencias de nombre, cargo, oficina y gerencia sí aparecen como observaciones.",
            (480, 1053): "No iniciado → omitir; vencido → cargo original.",
            (480, 1076): "Vigente: ambos marcados → actual; solo designado → designado.",
            (930, 1053): "la temporal sustituye esas diferencias ordinarias",
            (930, 1076): "vencida: comparar cargo y oficina de origen",
            (995, 1157): "Carnet ausente del SALAR actual y cuenta única en Dominio → posible baja; carnet vacío, repetido o cruce ambiguo → revisión.",
            (900, 1235): "En esta conciliación sí se consulta la fecha inicio del temporal: los accesos que todavía no empiezan se omiten.",
        },
    },
    "flujo_analisis_telefonia.svg": {
        "desc": "Compara Dominio e Isabel por nombres y números; separa dobles internos de Dominio entre coincidencia exacta y revision manual.",
        "footer_y": 1240,
        "output": (
            "Excel Telefonia: inventario completo con altas de Dominio ausentes en Isabel y bajas sin candidato en Dominio.",
            "Las altas se detectan por número ausente en Isabel; las bajas por persona de Isabel sin nombre equivalente en Dominio.",
        ),
        "remove_paths": ["M1390 476 H1430"],
        "text": {
            (900, 69): "Dominio aporta personas y números; Isabel es el inventario de internos que se revisa",
            (1575, 307): "unir números de nombres repetidos",
            (1577, 458): "Nombre genérico en Isabel",
            (1577, 484): "revisar por número compartido",
            (610, 592): "unir abreviaciones solo si hay un nombre completo único",
            (815, 1017): "nombre normalizado y equivalencia controlada",
            (420, 1194): "Un campo puede tener varios números; las etiquetas genéricas se informan, sin tratarlas como personas.",
        },
    },
    "flujo_generacion_informe.svg": {
        "desc": "Lee el análisis de al menos un sistema interno, ASFI o INFOCRED, con las hojas de SALAR-Dominio y temporales cuando corresponden; clasifica tipos, filtra excepciones exactas y llena los reportes de la plantilla.",
        "footer_y": 1260,
        "output": (
            "Excel final: incluye Nombre entre Usuario y Cargo/s; las filas temporales etiquetan solo los cargos marcados en verde.",
            "Bajas: cuenta sin persona en SALAR ni Dominio o posible baja histórica; Dominio con matriz genera Altas/Roles según acceso esperado.",
        ),
        "remove_paths": ["M1385 750 H1430", "M1385 845 H1405 V830 H1430"],
        "rect_styles": [(1430, 710, "warning")],
        "temporal_no_branch": True,
        "text": {
            (350, 166): "Uno o varios: internos, ASFI o INFOCRED",
            (900, 314): "internos · ASFI/INFOCRED · Dominio · temporales",
            (850, 504): "Dominio con matriz: Alta o Roles por acceso",
            (850, 526): "Punto Reclamo: Activo cubre rol esperado",
            (467, 788): "coinciden sistema, tipo, usuario y valores",
            (1570, 773): "cuentas sin persona en",
            (1570, 743): "Grupo “ambos”",
            (1570, 796): "SALAR ni Dominio se",
            (1570, 819): "incluyen en ambos reportes",
            (997, 1130): "Usuarios Externos · Cargos Salar sin Matriz · Salar sin dominio · Errores de carga · Falsos Positivos aplicados, si corresponde.",
            (900, 1210): "La fecha sale de «Resumen»; la plantilla debe conservar sus hojas e imágenes institucionales.",
            (900, 1240): "Si la plantilla tiene una sola hoja principal, todas las observaciones se colocan allí.",
        },
    },
    "flujo_conversor_pdf_asfi_excel.svg": {
        "desc": "Extrae texto del PDF ASFI por coordenadas, identifica usuarios por correo, obtiene nombres y roles sin catálogo cerrado, y crea una fila Excel por rol o una fila vacía si no hay roles.",
        "footer_y": 1195,
        "output": (
            "Excel «Roles ASFI»: Correo, Nombre y Rol, filtro y encabezado azul; filas blancas con bordes negros.",
            "La interfaz primero genera el archivo temporal, después pide ubicación para guardarlo y limpia la selección al terminar.",
        ),
        "text": {
            (900, 69): "PDF con texto seleccionable → palabras y columnas → correo/nombre → roles → Excel",
            (860, 317): "agrupar palabras por su posición vertical",
            (1060, 527): "agrupar por correo sin distinguir mayúsculas",
            (1297, 595): "Sin un correo previo no se asignan roles; sin texto o encabezados legibles no se identifican usuarios.",
            (1080, 836): "Unir línea al rol anterior",
            (860, 961): "Añadir usuarios sin roles",
            (860, 987): "una fila con Rol vacío",
            (1572, 961): "Generar Excel",
            (1572, 987): "después elegir dónde guardarlo",
            (900, 1175): "El conversor conserva roles nuevos porque los lee de la columna ROLES, sin una lista predefinida.",
        },
    },
    "flujo_conversor_pdf_sintesis_excel.svg": {
        "desc": "Extrae texto del PDF de Síntesis, limpia encabezados, reconoce registros por patrón, separa campos y roles de una lista predefinida y escribe la hoja Operadores.",
        "footer_y": 1190,
        "output": (
            "Excel «Operadores»: CIU, entidad, agencia, código, teléfono, nombre, rol, cuenta única y estado.",
            "La interfaz genera el Excel, muestra cuántos registros contiene, pide ubicación de guardado y luego limpia la selección.",
        ),
        "text": {
            (900, 69): "PDF con texto → limpieza → patrón de operador → campos y rol → hoja Operadores",
            (1395, 317): "quitar encabezados y cortar en Totales/FIN DEL REPORTE",
            (1300, 643): "Fila fuera del patrón → no se extrae",
            (285, 1025): "Síntesis identifica campos por patrón y busca el rol en una lista predefinida.",
            (285, 1058): "Un cambio de orden, códigos o rol puede dejar datos en NOMBRE; conviene revisar el Excel y ajustar las reglas.",
            (285, 1090): "Interfaz: cargar PDF → Generar → ver cantidad de registros → elegir ubicación → limpiar selección.",
            (900, 1170): "Un PDF escaneado sin texto seleccionable o con otro formato puede quedar sin registros reconocidos.",
        },
    },
}


def main() -> None:
    for filename, data in UPDATES.items():
        path = OUTPUT / filename
        tree = ET.parse(path)
        root = tree.getroot()
        root.find(tag("desc")).text = data["desc"]
        update_text(root, data["text"])
        for x, y, class_name in data.get("rect_styles", []):
            update_rect_style(root, x, y, class_name)
        if filename == "flujo_analisis_salar_dominio.svg" and not any(
                element.get("id") == "campos-salar-extra" for element in root.iter(tag("text"))):
            ET.SubElement(root, tag("text"), {
                "id": "campos-salar-extra", "x": "355", "y": "179",
                "text-anchor": "middle", "class": "small",
            }).text = "oficina · gerencia · carnet"
        for shape in data.get("remove_paths", []):
            if any(el.get("d") == shape for el in root.iter(tag("path"))):
                remove_path(root, shape)
        if data.get("temporal_no_branch") and not any(
                el.get("id") == "temporal-no-conforme" for el in root.iter(tag("g"))):
            group = ET.SubElement(root, tag("g"), {"id": "temporal-no-conforme"})
            ET.SubElement(group, tag("path"), {
                "d": "M1145 542 V565", "class": "arrow-amber",
            })
            ET.SubElement(group, tag("text"), {
                "x": "1160", "y": "557", "class": "tag", "fill": "#c78313",
            }).text = "No"
        add_output(root, data["footer_y"], data["output"])
        ET.indent(tree, space="  ")
        tree.write(path, encoding="utf-8", xml_declaration=False)
        print(path.relative_to(ROOT))


if __name__ == "__main__":
    main()
