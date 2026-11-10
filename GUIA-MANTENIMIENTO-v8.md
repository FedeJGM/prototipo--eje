# Guía de mantenimiento del libro de neteo v8

Esta guía describe el Módulo 4 v8, su integración con el libro `.xlsm` y el botón **Exportar datos**. Las dos variantes comparten los flujos de validación, protección y gestión de hojas. La diferencia está en cómo conservan los formatos y el sistema de fechas en el `.xlsx` de salida. **Instale una sola variante.**

[Instalación y reversión](#instalación-y-reversión) · [Mapa del proyecto](#mapa-del-proyecto) · [Operación del libro](#operación-del-libro) · [Exportación de datos](#exportación-de-datos) · [Diagnóstico](#diagnóstico) · [Pruebas de aceptación](#pruebas-de-aceptación)

## Instalación y reversión

| Archivo del Módulo 4 | Cuerpo del `.xlsx` | Elegir cuando… |
|---|---|---|
| [Modulo-4-v8.txt](Modulo-4-v8.txt) | Valores con el `NumberFormat` de origen por celda: números, fechas, horas y formatos personalizados. Usa el sistema de fechas del libro origen. | Importa conservar la presentación numérica o de fechas. |
| [Modulo-4-v8-valores.txt](Modulo-4-v8-valores.txt) | Valores con formato `General`. Una fecha puede verse como número de serie. | Se necesita una salida sin formato numérico de origen. |

1. Guarde una copia íntegra del `.xlsm` anterior. Trabaje en otra copia, pause Autoguardado y cierre las autorizaciones vigentes con el código anterior completo. Conserve las fuentes del punto de partida.
2. Sustituya **todo el contenido del Módulo 4 existente** por el TXT elegido. Mantenga el nombre real del módulo. No anexe el TXT al código anterior ni importe las dos variantes. Si el libro tiene una clave administrada distinta de la base, coordine la integración sin publicar ni reponer una clave anterior.
3. Si la pareja Base1000 ya está instalada, conserve los módulos compatibles y `ThisWorkbook`. Si parte de una versión anterior, contraste el [mapa del proyecto](#mapa-del-proyecto) y complete el conjunto antes de ejecutar, guardar, cerrar o cambiar de hoja. No sustituya rutinas locales de los otros módulos indiscriminadamente.
4. Use **Depuración → Compilar proyecto VBA** en la copia. Asigne el botón **Exportar datos** a `ExportarDatos` y el botón de renombrado a `RenombrarPlantillaConCuif`, si corresponde. No reasigne los botones anteriores que ya apunten a sus macros. Complete las pruebas antes de instalar en uso habitual.

**Reversión:** descarte la copia de prueba si falla. Si solo cambió código y el libro no cambió de estado, reponga el conjunto compatible del punto de partida y compile. **Después de renombrar o eliminar hojas, o cambiar cuentas, fórmulas o enlaces en el libro, restaure el `.xlsm` completo anterior.** Reponer código no deshace nombres, referencias ni pérdidas de datos. Un `.xlsx` exportado no sustituye ese respaldo.

## Mapa del proyecto

| Componente | Responsabilidad y punto de entrada |
|---|---|
| Módulo 1: [BorrarDatos compatible](Modular-v7-Corte2-Base1000/Modulo-1-limpieza-zona-gris.txt) | Borra entradas A:G y U:V desde la fila 12. Llama a `LimpiarFormulasFinalesVacias` con la extensión anterior; no sustituya el Módulo 1 entero si contiene otras rutinas. |
| Módulo 2: [parche de Reclasificación](Modulo-2-v1-PARCHE.txt) | `copiar_reclasificacion` toma la hoja activa. Con el parche, `Hoja1` queda con sus encabezados si no hay entradas y las tablas de `Hoja2` permanecen en la continuación original. El TXT **no es un módulo completo**. |
| Módulo 3: [ALCON v3](Modulo-3-v3.txt) | `ALCON` toma la hoja activa y usa la hoja técnica ALCON. Con la versión compatible, habilita temporalmente la estructura para mostrar/ocultar ALCON y restaura el estado. El nombre operativo puede aparecer sugerido al guardar su salida. |
| Módulo 4: una variante v8 | Orquesta validación, importación, protección, MENÚ, plantillas, zona calculada y exportación. No duplique sus procedimientos. |
| [NeteoValidaciones](Modular-v7-Corte2-Base1000/NeteoValidaciones.txt) | `RevisarFilas`, reglas de entrada, listas, marcas rojas y cambios por bloques. |
| [NeteoDependencias](Modular-v7-Corte2-Base1000/NeteoDependencias.txt) | `VerificarDependenciasPlantilla` inspecciona referencias de creación/eliminación; `ExcluirZona` resta rectángulos conservados. |
| [NeteoRespaldos](Modular-v7-Corte2-Base1000/NeteoRespaldos.txt) | `CrearRespaldo`, `CapturarContenido`, `RestaurarContenido` y compatibilidad de rangos. Mantener el módulo; la eliminación de una hoja no invoca el auxiliar histórico de copia a disco. |
| [ThisWorkbook v4](ThisBook-v4.txt) | Conservar el módulo de objeto. `Workbook_Open` repone protección; `Workbook_SheetChange` coordina cambios; `Workbook_BeforeSave` valida; `Workbook_BeforeClose` finaliza permisos; eventos de hoja gestionan la salida de MENÚ. No pegar estos eventos en un módulo estándar. |

**Estado y fronteras.** Módulo 4 conserva en privado `hojaMantenimiento`, `respaldoMantenimiento` y `hojaCuentasMenu`. `CLAVE_HOJA` es la única constante compartida; no la copie ni exponga su valor. Los nombres definidos `_NeteoV3_Mantenimiento` y `_NeteoV4_Cuentas` ayudan a reconocer autorizaciones interrumpidas; los nombres de `PREFIJO_CONSERVAR` protegen zonas que se decidió conservar. No borre esos nombres indiscriminadamente. `ThisWorkbook` es dueño de `menuPendiente`. Los tres auxiliares son módulos internos (`Option Private Module`) y no son puntos de entrada para botones.

**Respaldo temporal no es copia del archivo.** Importación y mantenimiento usan respaldos temporales para revertir fallos manejados o cambios rechazados. `EliminarUltimaFilaMenu` también puede respaldar temporalmente su rango. `EliminarUltimaPlantilla` **no crea una copia en disco**; no hay deshacer automático tras eliminar. No use **End/Reset** para cancelar: puede perder referencias al respaldo y dejar Excel o la protección en otro estado.

## Operación del libro

### Protección, validación e importación

- **Operativas:** entradas A:G y U:V desde la fila 12; H:T es zona calculada y W en adelante permanece bloqueado. En mantenimiento se habilita A:V, incluidos encabezados. **MENÚ:** cerrado por defecto; el permiso de cuentas habilita A:B desde la fila 10 y el mantenimiento permite A:L, incluidos encabezados. DATOS, ALCON y el modelo PLANTILLA no se habilitan con `DesbloquearHoja`.
- `DesbloquearHoja` ofrece permiso para cuentas de MENÚ o **Otra acción** (mantenimiento). `VolverAProtegerHoja` finaliza y restablece el modo normal. Al salir de MENÚ hacia otra hoja del mismo libro, `AvisarSalidaMenu` puede preguntar si conserva el permiso. No es el mismo flujo que cambiar de libro.
- En mantenimiento, `GestionarCambioMantenimiento` conserva la edición aceptada o recupera el contenido inmediatamente anterior de la zona calculada y encabezados. Las entradas de un pegado mixto no se deshacen con esa decisión.
- `ValidarTodo_ReglasNegocio` revisa la operativa activa. `Workbook_BeforeSave` finaliza permisos y llama a `ValidarTodasLasHojasTrabajo`, también si se guarda desde MENÚ. Los errores bloquean el guardado; el permiso puede haber terminado aunque no se guarde. La validación puede actualizar datos según las reglas de Redes; **no equivale** a comprobar marcas rojas al exportar.
- `ImportarDatosDesdeArchivoExcel` pega valores en la zona editable sin reorganizar columnas. Comprueba todo el destino, confirma reemplazo y respalda el rango afectado. Una cancelación previa no cambia datos; ante un fallo manejado intenta restaurar. Los datos importados inválidos se marcan para corrección y bloquean el guardado. Un cierre forzado o un reinicio de VBA no está cubierto por esa recuperación.

### Base H12:T1000 y limpieza dinámica

`ULTIMA_FILA_BASE = 1000` protege las fórmulas preparadas de H12:T1000 frente a **limpieza automática**; no es un tope de datos. `ActualizarEntradaEnHoja` prepara las filas afectadas. `LimpiarFormulasFinalesVacias` solo actúa desde la 1001 en el tramo final afectado, después de la última entrada, y solo sobre fórmulas automáticas que coinciden con el modelo. No borra constantes, excepciones ni fórmulas distintas; conserva H de la siguiente fila de entrada. Por ejemplo, reducir la última fila de 1050 a 1004 puede limpiar las fórmulas elegibles de 1005:1050 y conservar H1005. No es una limpieza general de sobrantes históricos, estilos o filas físicas.

Para huecos heredados, active la operativa afectada → `DesbloquearHoja` → **Otra acción** → `RecuperarBaseZonaGris` → revise → `VolverAProtegerHoja`. La recuperación es manual y por hoja; rellena solo vacíos elegibles de H12:T1000. Usa la fórmula especial de H y el modelo I:T de PLANTILLA; no sustituye fórmulas existentes o vacíos conservados. Si faltan fórmulas modelo o hay estructuras incompatibles, deténgase y revise antes de modificar el rango. `BorrarDatos` no sustituye esta recuperación.

### MENÚ y hojas CUENTA

| Acción | Comportamiento |
|---|---|
| `AgregarFilaMenu` | Prepara la fila siguiente a la última cuenta en A o B, sin insertar ni desplazar. Si C:L está vacío, aplica el modelo C10:L10 y deja K:L de una fila nueva visualmente vacíos. Si las diez celdas ya contienen fórmulas sin `#REF!`, conserva su contenido y formato. Si la fila está incompleta, contiene constantes o `#REF!`, se detiene sin sobrescribir. La fila 10 se conserva como modelo. |
| `EliminarUltimaFilaMenu` | Pide clave y confirmación. Limpia A:L de la última fila preparada sin desplazar otras filas; no elimina la fila 10 ni una fila asociada a una hoja. Distinguir de eliminar una hoja operativa. |
| `CrearNuevaHojaDesdePlantilla` | Requiere BV en A, CUIF en B, fórmulas preparadas en C:L y secuencia válida. Copia el modelo reservado **PLANTILLA**, crea `CUENTA nn - CUIF` y vincula K:L de la fila correspondiente. Reutiliza autorización de MENÚ cuando procede; la clave de creación no concede ese permiso por sí misma. |
| `RenombrarPlantillaConCuif` | Pide clave y confirmación para renombrar **solo la operativa activa**. Toma B de su fila asociada y conserva el ordinal. Si B está vacío o devuelve texto vacío, deja `CUENTA nn`; un dato no vacío e inválido no se recorta ni se sustituye. No renombra por editar B ni hace migración masiva. Verifica K:L antes y después, deja a Excel actualizar las referencias directas y no reescribe fórmulas de otras hojas. |
| `EliminarUltimaPlantilla` | Pide clave y confirmación. Solo elimina la hoja del mayor ordinal, si este es mayor que 01 y las dependencias lo permiten. Conserva su fila, cuentas y C:J de MENÚ; K:L quedan visualmente vacías. Nunca elimina la 01, con ninguno de los prefijos, ni el modelo PLANTILLA. **No crea un respaldo a disco.** |

**Identidad y CUIF.** `PLANTILLA nn` y `CUENTA nn`, con o sin sufijo, comparten el mismo ordinal: **fila MENÚ = n + 9**. No deben coexistir dos hojas con el mismo ordinal ni dejarse huecos. Mantenga dos cifras para 01–99 y numeración natural a partir de 100. El sufijo usa ` - ` y dígitos ASCII. Ingrese CUIF como texto para conservar ceros iniciales o más de 15 dígitos. Si B ya es numérico, solo se acepta un entero no negativo de hasta 15 dígitos. El nombre completo admite 31 caracteres; `CUENTA 01 - ` deja hasta 19 dígitos para el sufijo, y el espacio baja si crece el ordinal. Un CUIF vacío solo se permite al renombrar una hoja existente, no al crear ni al exportar una hoja con registros.

**Modelo K10:L10.** Si existe la 01, ambas fórmulas modelo deben apuntar directamente a su nombre real. Antes de la primera creación se admite el nombre provisional `PLANTILLA 01` o `CUENTA 01`, pero las dos fórmulas deben usar el mismo. La creación toma el modelo antes de asignar el nombre nuevo y conserva las referencias relativas. No sustituya el nombre con búsqueda global de texto. La creación consulta hasta cuatro variantes de nombre en el guardián de dependencias; el botón de renombrado no realiza ese barrido. El guardián puede detenerse por referencias rotas, dinámicas, tridimensionales u otras dependencias no comprobables.

## Exportación de datos

`ExportarDatos` no pide contraseña. Requiere una hoja **CUENTA numerada activa de este `.xlsm`** y ningún mantenimiento activo o marcado como interrumpido. No desprotege ni termina el mantenimiento por su cuenta. Si no hay entradas, informa y no abre el selector ni crea un libro. Si hay entradas, el sufijo debe coincidir con el CUIF válido en B de su fila MENÚ; si no coincide, corrija B si corresponde y ejecute `RenombrarPlantillaConCuif`.

1. Determina la última entrada real en A:G/U:V, también si está oculta o filtrada. Copia A11:V11 a A1:V1 y conserva en el cuerpo las filas desde 12 hasta esa última entrada, incluidos huecos y filas ocultas/filtradas. La base hasta 1000 no limita la exportación. H:T no determina el último registro.
2. Si detecta **relleno directo rojo exacto `RGB(255, 0, 0)`** en A:G/U:V dentro de las filas exportadas, se detiene, aunque la celda marcada esté vacía u oculta. Lee el estado: no llama a la validación que corrige datos, no modifica marcas ni interpreta cualquier color o regla condicional como rojo de validación.
3. Si H:T contiene errores nativos de Excel, muestra **un aviso** y permite confirmar o cancelar. Si confirma, los errores se conservan como errores. Un texto literal como `#N/D` no activa el aviso. La copia conserva los textos con ceros iniciales o que parecen fórmulas como texto.
4. El selector sugiere **`CUIF-AAAAMMDD-HHMMSS.xlsx`**, sin sufijo «valores». Cancelar el selector no crea el libro temporal; si se rechaza la continuación, no se guarda el archivo. Excel mantiene el aviso de reemplazo si el archivo elegido ya existe. La macro guarda un `.xlsx` de una hoja y cierra solo el libro de salida; no guarda ni cierra el origen.

Las dos variantes conservan anchos y estilo visible del encabezado sin trasladar sus reglas. No exportan fórmulas, macros, vínculos, conexiones, objetos, validaciones ni formatos condicionales, ni copian rellenos o estilos del cuerpo. En la variante con formato, un color incorporado al **código numérico** puede seguir visible: no equivale a copiar rellenos o reglas. Los valores se escriben por bloques; la variante con formato copia `NumberFormat` después de escribir, por columna y tramo uniforme, y subdivide solo tramos mixtos. No se ha medido cuánto tarda en el libro real; una columna muy heterogénea puede aumentar el trabajo.

## Diagnóstico

| Síntoma | Dónde empezar |
|---|---|
| No aparece una macro o hay nombre ambiguo | Compile y compruebe que solo exista una variante v8 y que la figura apunte a `ExportarDatos` o `RenombrarPlantillaConCuif` en el `.xlsm` correcto. |
| La hoja CUENTA no exporta | En `ExportarDatos`, revise hoja activa, mantenimiento, última entrada, B/sufijo y marcas rojas. Para un último registro filtrado, siga `UltimaFilaParaExportarDatos`; para un error H:T, revise `HayErroresParaExportarDatos` sin confundirlo con texto. |
| La copia muestra seriales o formato inesperado | Confirme la variante instalada. En la de formato, revise `Date1904`, `EscribirValoresExportarDatos` y `CopiarFormatoNumericoExportarDatos`. Distinga el valor subyacente de lo mostrado en la celda. |
| No crea, renombra o elimina una hoja | Revise `VerificarSecuenciaPlantillas`, `NombrePlantillaConCuif`, `VerificarModeloMenu`, `VerificarVinculosFilaMenu` y, para creación/eliminación, `VerificarDependenciasPlantilla`. No cambie el ordinal ni desplace cuentas para forzar el flujo. |
| Permisos, eventos o fórmulas no se comportan como antes | Revise `Workbook_SheetChange`, `FinalizarMantenimiento`, `RecuperarMantenimientoInterrumpido` o `RestablecerProteccionAlAbrir`. Para crecimiento o huecos, revise las rutinas de zona gris y confirme qué versión de Módulo 1 está instalada. |

En una copia, busque el procedimiento por nombre. Use F9 para un punto de interrupción y F8 para avanzar. **Ver → Ventana Inmediato** permite consultar sin cambiar estado: `? Application.EnableEvents`, `? Application.ScreenUpdating`, `? Application.DisplayAlerts`, `? ThisWorkbook.ProtectStructure`, `? ActiveSheet.ProtectContents`. `ProtectContents=True` no prueba que una celda concreta esté bloqueada; revise también su propiedad `Locked`. Si ocurre un fallo, registre rutina, paso, mensaje y número; no registre claves ni cuentas. Si hubo una interrupción forzada, verifique estados y protección antes de guardar.

## Pruebas de aceptación

Ejecute en una copia del `.xlsm` y compare los resultados con el libro anterior. No basta con un mensaje de éxito.

- [ ] **Instalación y arranque:** compilar, ejecutar botones reales, guardar/reabrir y comprobar bloqueos, permisos y estados de Excel. Confirmar las versiones instaladas de Módulos 1–3.
- [ ] **Variantes:** exportar la misma muestra con cada variante, en instalaciones separadas. Comprobar importes negativos y positivos, fecha/hora, formatos personalizados y mixtos en una columna, texto `0012` y literal con `=`. En la de formato, probar un origen con sistema 1904. Revisar una hoja de salida, encabezados, anchos, valores y ausencia de fórmulas o vínculos.
- [ ] **Alcance y guardias:** último registro filtrado más allá de 1000, fila oculta y huecos. Marcar rojo exacto una celda vacía/oculta dentro de la zona editable y comprobar que se bloquea. Contrastar con un rojo distinto, una marca fuera del rango y un texto `#N/D`. Un error nativo en H:T debe advertir una vez; probar **No** y **Sí**.
- [ ] **Guardado:** comprobar nombre sugerido y extensión, cancelar el selector y rechazar el reemplazo de un archivo existente. Verificar que el origen sigue abierto, sin guardar cambios por la exportación, y que se restauran eventos, pantalla y alertas.
- [ ] **Regresiones:** reglas de entrada y guardado desde MENÚ; importación, cancelación y reversión; zona base 12/800/1000, crecimiento y reducción 1050→1004/900; recuperación con y sin permiso. Probar MENÚ vacío/completo/incompleto, creación con BV y CUIF, ordinal repetido, renombre a `CUENTA nn` si B está vacío y vínculos K:L. Eliminar solo en una copia desechable: primero No; luego verificar que se conserva la fila y que nunca se borra la 01.
- [ ] **Salidas y consumidores:** comparar Reclasificación y ALCON con resultados esperados. Comprobar catálogos, vínculos externos y cualquier consumidor Python con los nombres nuevos. No se verificaron esos consumidores ni se modificaron los vínculos por esta entrega.

**Alcance de verificación:** se revisaron las fuentes y modelos de comportamiento; no se compiló ni ejecutó el proyecto VBA v8 completo en Excel. Quedan por probar el libro real, su volumen y dependencias externas. El máximo de 100 caracteres que `EsNitValido` aplica al NIT sigue siendo un punto de negocio pendiente de decisión; no se confunde con el CUIF ni con un límite de filas.
