# Analizador de Accesos

Aplicación de escritorio para revisar cuentas, roles y datos de acceso de sistemas institucionales. Compara las cuentas que exporta cada sistema con las personas y cargos de SALAR, la información complementaria de Dominio y las reglas de la matriz de accesos. También genera un informe basado en plantilla, concilia SALAR con Dominio, revisa el directorio telefónico y convierte reportes PDF de ASFI y Síntesis a Excel.

El análisis señala diferencias y casos pendientes para revisión. Sus resultados son una ayuda de control: no ejecuta altas ni bajas en los sistemas institucionales.

## Funciones

La interfaz tiene cinco áreas principales:

| Área | Qué hace | Entradas principales |
| --- | --- | --- |
| **Analizar sistemas** | Compara uno o más sistemas seleccionados, internos o externos (ASFI/INFOCRED). Funciona también con un único sistema externo. Genera un Excel de análisis con resumen, observaciones, detalle y errores. | Matriz, SALAR, Dominio y TXT de al menos un sistema; temporales opcionales. |
| **Generar reporte** | Toma un Excel de análisis y rellena la plantilla institucional. Separa las revisiones SALAR-Matriz y Dominio-Matriz cuando la plantilla contiene ambas hojas. | Excel de análisis, plantilla incluida en `documentos base` y falsos positivos opcionales. |
| **Telefonía** | Compara nombres y números entre Dominio y el inventario de internos de Isabel. Sugiere altas, bajas y revisiones de titularidad. | TXT de Dominio y TXT de internos Isabel. |
| **Conversor PDF ASFI a Excel** | Extrae correo, nombre y roles de un PDF de ASFI. | PDF del reporte ASFI con texto seleccionable. |
| **Conversor PDF Síntesis a Excel** | Reconstruye registros de operadores de un PDF de Síntesis. | PDF con el formato de reporte esperado por el procesador. |

En las vistas que generan archivos, primero se procesa la información y después se elige dónde guardar el Excel. Tras guardarlo, la vista limpia los archivos seleccionados. Los conversores PDF también muestran cuántas filas o registros se generaron.

## Cómo se analiza el acceso

### Sistemas internos

La revisión parte de las personas de **SALAR**. Para cada sistema seleccionado, el programa lee su TXT y usa la clave configurada para relacionar cada cuenta con SALAR. Según el sistema, la clave puede ser login o correo. Las transformaciones configuradas del login —por ejemplo, quitar un prefijo— se usan para el cruce.

El cargo de la persona se busca en la matriz para obtener los accesos esperados. Luego se contrastan con lo encontrado en el sistema:

- Para sistemas con roles, compara los códigos esperados con los códigos encontrados. Los roles faltantes o adicionales producen una diferencia.
- **CIA** compara roles y organigrama.
- **Punto de Reclamo** informa estado en lugar de códigos de rol. Cuando la matriz espera acceso, un registro `Activo` se considera conforme; primero se filtran los registros `Inactivo`.
- **SRGNC** cruza principalmente por correo, puede usar el login de respaldo configurado y excluye el estado inactivo definido para ese sistema.
- Otros ajustes —como agrupar filas repetidas o normalizar códigos monetarios de PRESUPUESTO— se aplican solo a los sistemas que los necesitan.

Dominio complementa la revisión, no reemplaza SALAR como fuente principal:

En los TXT de personas, el correo literal `Null` se interpreta como un correo ausente. Esas filas no se cruzan ni generan altas automáticamente por ese valor.

1. Si una persona SALAR no tiene correo, se busca un nombre idéntico en Dominio. Solo una coincidencia única aporta el login y correo de Dominio; el cargo y los demás datos laborales siguen siendo de SALAR. Los nombres ambiguos quedan para revisión.
2. Si una cuenta del sistema no se encuentra en SALAR, se busca una identidad única en Dominio. Si aparece, el cargo de Dominio permite consultar la matriz y comparar el acceso. Si no aparece en ninguna fuente, se clasifica como cuenta externa pendiente de revisar o retirar.
3. Se recorren las personas de SALAR y las personas que solo están en Dominio para detectar accesos esperados sin cuenta. Si la matriz asigna acceso, se genera una observación de alta; si no lo asigna, no se crea esa alerta.

Los duplicados, cargos ambiguos en la matriz, formatos de rol desconocidos y coincidencias no únicas no se resuelven eligiendo una fila al azar: se registran como error o caso no evaluable.

### Accesos temporales

El Excel de accesos temporales es opcional. Sus columnas se pueden configurar. Si un acceso temporal aplica, ajusta el cargo (y oficina cuando corresponde) usado para consultar la matriz y reemplaza la observación ordinaria del campo evaluado.

- Si la asignación está vigente, se usa el cargo designado. Cuando las celdas de cargo actual y designado están marcadas, se consideran ambos.
- Si la fecha de inicio todavía no llegó, no se evalúa como temporal vigente.
- Si la fecha final ya venció, se vuelve a evaluar con el cargo de origen y puede aparecer como **Fin Acceso Temporal**.
- Para una persona sin cuenta, el cargo temporal también puede cambiar los accesos que se esperan al revisar un alta.

La configuración por sistema define cómo se reconocen sus columnas, roles, estado, claves y casos especiales. En la interfaz se puede ajustar desde los engranajes; esos valores se conservan en las preferencias de la aplicación.

### ASFI dentro del análisis combinado

ASFI se analiza con una lógica específica porque sus roles descriptivos deben traducirse a los códigos de la matriz:

1. Se agrupan las filas del TXT ASFI por correo y se reúnen todos los roles de cada cuenta.
2. La hoja `Codigos` de la matriz relaciona los roles ASFI del TXT con uno o más roles de matriz. La columna H de `Matriz de Accesos` contiene los accesos esperados por cargo.
3. El correo se cruza primero con SALAR y luego, ante una coincidencia única, con Dominio. Se comparan los roles traducidos con los esperados para el cargo.
4. También se buscan personas de SALAR o solo de Dominio que deberían tener roles ASFI pero no aparecen en el TXT, para señalar altas pendientes.

El archivo que consume el análisis ASFI es un **TXT** con correo, nombre y rol en los campos configurados. El conversor PDF ASFI, en cambio, produce un **Excel independiente** de correo, nombre y rol; ese Excel no se carga automáticamente como entrada del análisis.

### INFOCRED dentro del análisis combinado

INFOCRED utiliza el motor compartido de `src/sistemas_externos/comun.py`: agrupa por correo, obtiene el nombre y cargo desde SALAR o, si no encuentra a la persona allí, desde Dominio; traduce los roles TXT a roles de matriz y compara la expectativa por cargo. Su configuración está en `src/sistemas_externos/infocred.py`, separada de la de ASFI en `asfi.py`. Para el TXT actual se leen 17 campos: correo en el campo 7, rol en el campo 9 y estado en el campo 13. Se excluyen las filas cuyo estado es `BAJA`; luego se agrupan los roles restantes por correo.

La hoja `Codigos`, sección `Infocred` en la columna A, contiene la equivalencia de rol TXT en B a rol de matriz en C. También se acepta `lnfocred` con ele inicial, como aparece en la matriz vigente. La columna P de `Matriz de Accesos` contiene los roles esperados por cargo. INFOCRED puede cargarse junto con ASFI y los sistemas internos; genera sus propias hojas `Observaciones INFOCRED`, `Temporal INFOCRED`, `INFOCRED Detalle` e `INFOCRED Errores`.

### Conciliación SALAR-Dominio

Cuando se proporciona ambas fuentes, el análisis combinado agrega una comparación de sus registros. Busca parejas en este orden:

1. Login igual y único.
2. Correo igual y único para las filas aún sin pareja.
3. Nombre completo igual y único, únicamente cuando SALAR no tiene correo.

Para las parejas encontradas compara login, correo, nombre, cargo, oficina y gerencia. Las observaciones resumen diferencias de nombre o datos laborales; el detalle conserva también el método de coincidencia y los valores de ambas fuentes. En el reporte visible, los datos laborales proceden de SALAR y el login mostrado es el de Dominio cuando existe esa pareja.

Los registros solo en Dominio se clasifican aparte. Si se cargó la matriz, se marca si su cargo tiene correspondencia: los cargos sin matriz alimentan **Usuarios Externos**; las personas con cargo en matriz se revisan en **Reporte Dominio-Matriz**. También se pueden cargar SALAR anterior y accesos temporales para los análisis históricos y temporales.

### Hojas del Excel de análisis

El libro de revisión es la fuente para la opción **Generar reporte**. Sus hojas principales son:

- `Resumen`: archivos usados, cantidades por sistema, observaciones, casos no evaluables y errores.
- `Observaciones <Sistema>`: casos que requieren revisión para cada sistema. Si hay observaciones temporales, se agrega `Temporal <Sistema>` con cargos y fechas.
- `Detalle de comparación`: incluye también coincidencias conformes y casos que no se pudieron evaluar.
- `Errores de carga`: archivo, fila, campo afectado y descripción del problema.
- Si se cargó Dominio junto con SALAR: `Observaciones Dominio`, `Dominio Usuarios Temporales`, `Detalle SALAR-Dominio` y `Errores SALAR-Dominio`; `Historial SALAR-Dominio` aparece cuando se proporciona SALAR anterior.
- Si se cargó ASFI: `Observaciones ASFI`, `Temporal ASFI`, `Externos Detalle de comparación` y `Externos Errores de carga`.
- Si se cargó INFOCRED: `Observaciones INFOCRED`, `Temporal INFOCRED`, `INFOCRED Detalle` e `INFOCRED Errores`.

## Informe final

La opción **Generar reporte** no vuelve a analizar los TXT. Lee el Excel de análisis ya producido, adapta las distintas hojas a una estructura común y llena la plantilla `documentos base/plantilla informe.xlsx`.

- **Reporte Salar-Matriz** agrupa los casos asociados a SALAR.
- **Reporte Dominio-Matriz** agrupa los casos de personas solo en Dominio cuyo cargo tiene regla en la matriz.
- Una cuenta sin identidad en SALAR ni Dominio puede aparecer en ambos reportes para que no se pierda la baja pendiente.
- Las categorías se presentan como Altas, Accesos Temporales, Bajas, Fin Acceso Temporal, Roles u Otros, con filas agrupadas por sistema y conteos.
- Las hojas auxiliares organizan Usuarios Externos, Cargos Salar sin Matriz, Salar sin Dominio y errores de carga.
- Los usuarios externos de Dominio con diferencia de acceso pueden producir Altas o Roles en el reporte general según el valor esperado y encontrado.
- Los falsos positivos son opcionales. Solo se aplica una excepción si coinciden sistema, tipo, usuario, cargo y valores esperado/encontrado; las excepciones aplicadas quedan documentadas en la plantilla.

El Excel generado abre en **Reporte Salar-Matriz** y muestra las pestañas desde la primera hoja, aunque la plantilla se haya guardado mostrando hojas posteriores.

El informe deja vacías las celdas que representan ausencia de acceso o una cuenta aún no creada. El texto `Null` en el valor encontrado se conserva. En Punto de Reclamo, `Activo` puede satisfacer un acceso esperado representado por un código de la matriz. Los temporales conformes no se muestran como observaciones.

## Telefonía

El análisis de Telefonía es independiente de las matrices de roles. Lee nombres y teléfonos de Dominio y de Isabel/Internos; extrae las secuencias de dígitos y agrupa registros por nombre normalizado. Puede generar estas observaciones:

- **Dar de Alta en Isabel**: una persona de Dominio tiene un número que todavía no está en el inventario de internos.
- **Dar de Baja en Isabel**: una persona con interno no aparece en Dominio.
- **Teléfono con diferente usuario**: ambas fuentes tienen el mismo número, pero los titulares no son equivalentes.
- **Usuarios con doble interno Isabel**: una persona tiene más de un interno; las abreviaciones se agrupan cuando hay un único nombre completo candidato.
- **Doble interno Dominio - Coincidencia exacta por nombre**: el nombre normalizado de Dominio coincide exactamente con uno de Isabel.
- **Doble interno Dominio - Revision manual**: no hay una coincidencia exacta unica; se muestran posibles candidatos de Isabel cuando existen.
- **Nombres Genéricos Isabel**: una etiqueta como un área o una agencia está cargada como titular; el número se usa para mostrar posibles titulares en Dominio.

El Excel mantiene la hoja completa `Observaciones` y agrega las hojas `Coincidencia exacta` y `Revision manual`. La columna de categoria conserva motivos como doble interno Dominio, doble interno Isabel o dar de baja en Isabel. La hoja `Inventario completo` lista todos los grupos de ambas fuentes con su cruce, candidatos y filas de origen. En el inventario, los telefonos de Dominio ausentes en Isabel se anotan como `Dar de Alta en Isabel`; las identidades de Isabel sin candidato en Dominio como `Dar de Baja en Isabel`.
Las columnas de nombre y teléfono son configurables. Por defecto, Dominio usa campos 1 y 11; Isabel/Internos, campos 1 y 2. El resultado es un Excel con la hoja `Observaciones`.

## Conversores PDF a Excel

### ASFI

El procesador usa `pdfplumber` y la posición de las palabras para localizar las columnas `SISTEMA`, `ROLES` y `FECHA ALTA`. Agrupa las filas por correo, extrae el nombre entre las etiquetas `Nombre` y `Doc. Ident.`, y recupera el texto dentro de la columna de roles. Las líneas de rol con fecha se agregan como roles nuevos; fragmentos cercanos sin fecha pueden completar el rol anterior. No requiere un catálogo de roles conocido.

El Excel resultante contiene la hoja `Roles ASFI` con columnas `Correo`, `Nombre` y `Rol`. Cada rol ocupa una fila; usuarios sin roles se conservan con el rol vacío. Por eso, el contador que muestra la interfaz cuenta filas/roles y no personas únicas.

### Síntesis

El procesador extrae el texto de las páginas, elimina encabezados y separadores, identifica registros mediante un patrón de códigos y separa campos como CIU, entidad, agencia, código, teléfono, nombre, rol, cuenta única y estado. El Excel contiene la hoja `Operadores`.

Este procesador depende del orden y formato de los registros y de una lista fija de roles conocidos. Un rol nuevo o un cambio en el PDF puede requerir actualizar esa lista o el patrón. Los registros cuyo nombre quede vacío se descartan.

## Entradas y configuración predeterminada

Los TXT institucionales se leen delimitados por `|`. Los lectores reconocen UTF-8, Windows-1252 y Latin-1 y validan la cantidad de campos configurada. La interfaz permite cambiar las posiciones desde los engranajes; las columnas de los reportes pueden variar, así que confirma los campos antes de procesar archivos nuevos.

Valores iniciales de personas:

| Fuente | Campos del TXT | Nombre | Cargo | Login | Datos adicionales |
| --- | ---: | --- | ---: | ---: | --- |
| SALAR | 20 | campos 1–5 | 17 | 20 | oficina 15, gerencia 14, correo 19, carnet 7 |
| Dominio | 13 | campo 1 | 4 | 13 | oficina 10, gerencia 3, correo 5 |

Los sistemas internos configurados son:

| Sistema | Columna de acceso en matriz | Regla destacada |
| --- | --- | --- |
| ArgosAPX | R | Roles; `Null` se interpreta como ausencia de roles. |
| ArgosNet | S | Roles; `Null` se interpreta como ausencia de roles. |
| BDPSAM | W | Ajusta el prefijo configurado del login. |
| CIA | AA y AB | Compara roles y organigrama. |
| FG | X | Ajusta el prefijo y sufijo configurados del login. |
| SEGA | Z | Comparación por roles. |
| SRGNC | AO | Cruce por correo con login de respaldo; ignora el estado `INA`. |
| PRESUPUESTO | BD | Agrupa roles por login y normaliza códigos con formato monetario. |
| Punto de Reclamo | AD | Compara existencia y estado, no códigos de rol. |

La matriz interna usa por defecto la hoja `Matriz de Accesos` y el cargo en la columna B. ASFI consulta la columna H e INFOCRED la columna P para sus accesos. Ambos usan la hoja `Codigos` para traducir roles TXT a roles de matriz; las columnas y sección son configurables.

## Instalación y ejecución

Se recomienda Python 3.11 o posterior.

### Windows (PowerShell)

Desde la carpeta raíz del proyecto:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe interfaz.py
```

También puedes activar el entorno primero y luego ejecutar `python interfaz.py`.

