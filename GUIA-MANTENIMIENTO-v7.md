# Guía de mantenimiento — Macro Neteo Reclasificación v7

## Ruta rápida: leer esto primero

**Esta entrega solo organiza y documenta Módulo 4.** Las **102 rutinas y sus firmas/cuerpos se conservan byte por byte**; no se implementan módulos nuevos, clases ni cambios de reglas. **No se certifica producción**: faltan las regresiones completas en el libro real.

1. Trabajar en una copia del libro y conservar el código anterior. Sustituir **todo el contenido del Módulo 4 existente** por [Modulo-4-v7-organizado.txt](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modulo-4-v7-organizado.txt), sin crear otro módulo, duplicar rutinas ni pegarlo en ThisWorkbook.
2. Mantener ThisWorkbook y Módulo 3 actuales, incluidas sus correcciones ya hechas en el libro. Fuentes de referencia: [ThisBook-v4.txt](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/ThisBook-v4.txt) y [Modulo-3-v3.txt](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modulo-3-v3.txt); **no reinstalarlos por esta entrega**.
3. En el editor VBA, usar **Depuración → Compilar VBAProject**. Si falla, detenerse y compartir la línea. La comprobación de igualdad de bytes no sustituye compilar el proyecto instalado.
4. Probar los flujos de la lista de verificación en la copia. En la primera prueba de **Eliminar plantilla**, responder **No** a la confirmación: no hace falta eliminar para verificar su revisión previa.
5. Para localizar código, usar **Ctrl+F con el nombre del procedimiento**. Buscar `[G1]`…`[G9]` salta a las secciones; el catálogo al final muestra las líneas del archivo entregado, que pueden cambiar en futuras versiones.

**Tres distinciones esenciales:** agregar fila no crea hoja; eliminar fila no elimina plantilla; desbloquear funcionalmente una hoja no significa dejar `ProtectContents=False`.

La medición reportada por el usuario pasó de **48,727 s a 2,391 s** en una prueba **previa a la confirmación**. No es una prueba de exportación, eliminación efectiva, guardado/reapertura ni resultados contables completos. Esta reorganización no añade otra optimización.

---

## 1. Mapa de responsabilidades, no módulos separados

| Sección | Responsabilidad | Rutinas |
|---|---|---:|
| G1 | Acceso a hojas y utilidades de diseño | 8 |
| G2 | Reglas y validación de datos | 22 |
| G3 | Validaciones de Excel y catálogo de redes | 19 |
| G4 | Importación de valores | 1 |
| G5 | Entradas, última fila y fórmulas calculadas | 8 |
| G6 | Protección y autorizaciones | 11 |
| G7 | Respaldos, mantenimiento y excepciones | 12 |
| G8 | MENÚ: cuentas, filas y vínculos | 12 |
| G9 | Plantillas y comprobación de dependencias | 9 |

Se sigue utilizando **un único Módulo 4**. Las secciones ordenan la lectura, no son capas arquitectónicas ya implementadas. Las llamadas cruzadas existentes permanecen: por ejemplo, importar utiliza validación, protección y respaldos; MENÚ y plantillas comparten comprobaciones.

### Botones/macros: asignaciones esperadas

Este mapa muestra los procedimientos previstos, **no certifica las asignaciones reales de las figuras**: no se inspeccionó el libro instalado. Conservar nombres de módulos, macros, argumentos y visibilidad; no asignar una función auxiliar como sustituto de un botón. **No eliminar una rutina porque no se vea llamada desde el código**: puede estar asignada a una figura o invocada externamente; falta un inventario real de OnAction.

| Acción prevista | Procedimiento | Responsabilidad y hojas |
|---|---|---|
| Validar entradas | `ValidarTodo_ReglasNegocio` | Hoja operativa activa; valida, marca y aplica las correcciones existentes. |
| Importar valores | `ImportarDatosDesdeArchivoExcel` | Rango origen y destino operativo editable; sin reorganizar columnas. |
| Desbloquear/autorizar | `DesbloquearHoja` | MENÚ: cuentas A/B o mantenimiento A:L; operativa: mantenimiento A:V. |
| Volver a proteger | `VolverAProtegerHoja` | Cierra autorización/mantenimiento y restaura bloqueos funcionales. |
| Agregar fila MENÚ | `AgregarFilaMenu` | Prepara la siguiente fila según última información A/B; copia modelo C10:L10. No crea hoja. |
| Eliminar última fila MENÚ | `EliminarUltimaFilaMenu` | Limpia A:L, sin desplazamientos; nunca fila 10; bloquea una hoja asociada. |
| Crear plantilla | `CrearNuevaHojaDesdePlantilla` | Copia PLANTILLA, crea PLANTILLA nn y enlaza MENÚ. |
| Eliminar última plantilla | `EliminarUltimaPlantilla` | Solo última numerada permitida, con clave, confirmación y respaldo completo. Conserva su fila de MENÚ. |
| Configurar validaciones, uso técnico | `ConfigurarValidacionesEntradaHojaActiva` | Repone reglas de la operativa; no sustituye las pruebas de negocio. |
| Restablecer protección, automático/técnico | `RestablecerProteccionAlAbrir` | MENÚ/operativas y estructura; no migra por sí solo la protección del modelo. |
| Borrar datos, Módulo 1 | `BorrarDatos` | Limpieza de entradas según la fuente disponible; no se cambió aquí. |
| Reclasificación, Módulo 2 | `copiar_reclasificacion` | Libro de salida Hoja1/Hoja2 según su código actual; no se cambió aquí. |
| ALCON, Módulo 3 | `ALCON` | Operativa → hoja ALCON → libro de salida; restaura visibilidad y estructura según su versión actual. |

### Eventos de ThisWorkbook: llamadas directas verificadas

Los eventos permanecen en **ThisWorkbook**, no en Módulo 4: Excel los despacha por su ubicación/nombre. La tabla enumera llamadas directas presentes; no significa que todas se ejecuten incondicionalmente.

| Evento real | Llamadas directas a Módulo 4 | Efecto |
|---|---|---|
| `Workbook_Open` | `RestablecerProteccionAlAbrir` | Reponer protección y finalizar autorizaciones previas; si falla, no editar/guardar sin revisar. |
| `Workbook_SheetChange` | `RecuperarMantenimientoInterrumpido`, `EsHojaTecnica`, `EstaEnMantenimiento`, `GestionarCambioMantenimiento`, `ActualizarEntradaEnHoja`, `ActualizarRespaldoDespuesDeEntrada` | Decidir mantenimiento antes de preparar/validar las entradas afectadas; restaurar eventos y pantalla. |
| `Workbook_BeforeSave` | `FinalizarMantenimiento`, `RestablecerProteccionAlAbrir`, `ValidarTodasLasHojasTrabajo` | Finalizar autorización, reponer protección y bloquear guardado si hay entradas operativas inválidas o un error. |
| `Workbook_BeforeClose` | `FinalizarMantenimiento` | Finalizar mantenimiento; si falla, cancelar cierre. |
| `Workbook_SheetDeactivate` | — | Recordar si la hoja que se abandona es MENÚ; no llama directamente a Módulo 4. |
| `Workbook_SheetActivate` | `AvisarSalidaMenu` | Si hay una salida pendiente de MENÚ dentro del libro, preguntar si se conserva la autorización. |
| `Workbook_Deactivate` | — | Limpiar la salida pendiente al cambiar de libro/ventana; no equivale a un cambio de pestaña interno. |


## 2. Contrato de hojas y columnas

| Área | Entrada normal | Cálculo/protección |
|---|---|---|
| Operativas, desde fila 12 | A:G y U:V | H:T calculadas/bloqueadas; encabezados A1:V11 y W en adelante bloqueados fuera de mantenimiento. V es obligatoria cuando la fila tiene datos. |
| MENÚ | Ninguna celda editable por defecto | Autorización de cuentas: solo A/B desde fila 10. Mantenimiento: A:L, incluidos encabezados. M en adelante sigue bloqueada. |
| PLANTILLA | Modelo de copia, no una operativa numerada | No reemplazarlo para resolver un cambio autorizado de una celda de otra hoja. |
| DATOS/ALCON | Hojas técnicas fuera del desbloqueo funcional normal | DATOS se consulta como catálogo; ALCON interviene en su exportación propia. Su protección de hoja no se migra aquí. |

**H no se reconstruye igual que I:T.** `PrepararFilasEntrada` llama a `CompletarColumnaCalculada` con columna 8: H usa la fórmula especial de fecha de origen B contra corte $B$2; para fila 12: `=SI(B12=0,"",SUMA($B$2-B12))`. El separador visible puede variar por configuración de Excel. `CompletarFormulasZonaGris` recorre columnas 9…20 (**I:T**) y utiliza las fórmulas modelo de fila 12 de PLANTILLA. La recuperación rellena vacíos elegibles y respeta las excepciones conservadas; no reescribe indiscriminadamente todas las fórmulas.

**PLANTILLA y PLANTILLA 01 son distintas.** La primera se copia para crear nuevas hojas. K10/L10 de MENÚ sirven como modelo y deben conservar su referencia a PLANTILLA 01. La secuencia numerada y la relación número+9 con la fila de MENÚ se comprueban antes de crear/eliminar; los huecos intermedios bloquean y no se renumeran automáticamente. Una última hoja reconstruida desde el modelo no recupera sus datos borrados.

El validador global recorre hojas no clasificadas como técnicas por `EsHojaTecnica`. Guardar desde MENÚ también dispara el evento del libro y revisa esas operativas; no significa que MENÚ/DATOS/ALCON/PLANTILLA pasen por el mismo validador de entradas.

## 3. Seguir un flujo real sin perder el contexto

### Importar valores

`ImportarDatosDesdeArchivoExcel` solicita origen/rango y usa el destino editable elegido. Comprueba el bloque completo antes de pegar: debe caber en A:G o U:V y no ser el propio libro origen. Confirma reemplazo de datos existentes y respalda el rango que se reemplazará; no reorganiza columnas. Tras importar, valida/prepara las filas afectadas, marca errores y mantiene el bloqueo de guardado hasta corregirlos. Cancelación, fallo manejado y recuperación del rango son casos de prueba diferentes; no suponer éxito por aparecer valores.

### Editar una celda operativa

`Workbook_SheetChange` desactiva temporalmente eventos/pantalla para evitar recursión. Primero revisa mantenimiento y su decisión conservar/recuperar; después `ActualizarEntradaEnHoja` valida y prepara las filas de entrada afectadas, recupera H de forma especial e I:T desde sus modelos y prepara la siguiente fila. `ActualizarRespaldoDespuesDeEntrada` mantiene el respaldo coherente si corresponde. La salida del evento restaura los estados anteriores; no debe terminarse con End/Reset.

### MENÚ: cuentas y filas

Para editar cuentas se requiere la autorización de `DesbloquearHoja`; la hoja vuelve a estar protegida con las celdas permitidas desbloqueadas. `AgregarFilaMenu` reutiliza autorización vigente o solicita clave; se basa en la última información A/B, no en la última fórmula preparada C:L. Por eso repetir el botón sin llenar cuentas puede volver a la misma fila. No inserta/desplaza filas ni crea operativas. `EliminarUltimaFilaMenu` pide clave siempre, considera contenido/fórmulas A:L, protege fila 10 y rechaza una asociación existente; conserva el borde inferior de la fila superior.

Al salir manualmente de MENÚ, sus eventos pueden preguntar si se bloquea o mantiene el permiso; se repite mientras continúe autorizado. Las navegaciones internas con eventos desactivados no generan ese mismo aviso. Guardar/cerrar o volver a proteger finaliza autorización según el flujo actual.

### Crear y eliminar plantillas

`CrearNuevaHojaDesdePlantilla` valida autorización, modelo, cuentas/fórmulas, secuencia y dependencias; copia PLANTILLA, ajusta encabezados desde MENÚ y enlaza K/L. Restaura visibilidad del modelo y estructura. No necesita rehacer las cuentas ni copiar la última operativa como modelo.

`EliminarUltimaPlantilla` pide contraseña siempre y permite solo la última numerada adicional: nunca PLANTILLA 01 ni PLANTILLA. Revisa dependencias antes de Sí/No y guarda un archivo completo antes de borrar. Conserva A/B y C:J de MENÚ; deja K/L provisionales. Puede detenerse por referencias directas/rotas/dinámicas/3D u otros elementos no verificables. El recorrido de fórmulas por bloques y de validaciones por grupos **no es permiso para omitir reglas ilegibles**. Si la hoja ya se borró y hay un fallo posterior, no se promete reconstrucción automática: usar el respaldo completo de forma manual y controlada.

### Guardar/reabrir

`Workbook_BeforeSave` finaliza autorizaciones, restablece protección y llama a `ValidarTodasLasHojasTrabajo`. Si detecta errores o un fallo, cancela el guardado; la autorización puede haber finalizado incluso si no se guardó. `Workbook_Open` repone protección y UserInterfaceOnly al reabrir. Hay que probar ambos eventos y verificar entradas/editabilidad, no únicamente el mensaje de guardado.

## 4. Estado privado, respaldos y salidas seguras

| Elemento | Significado | Regla de mantenimiento |
|---|---|---|
| `hojaMantenimiento` | Hoja con mantenimiento autorizado en esta sesión | Privada: no volverla Public ni duplicarla en futuros módulos. |
| `respaldoMantenimiento` | Libro temporal con contenido previo de la zona autorizada | No cerrarlo/descartarlo a mitad ni confundirlo con una copia de archivo recuperable tras cerrar Excel. |
| `hojaCuentasMenu` | Referencia de autorización de cuentas de MENÚ | Privada; finalizarla por las rutinas existentes. |
| `CLAVE_HOJA` | Constante central que comparten autorizaciones, hojas gestionadas y estructura | Conservar una definición; no publicar su valor en guías, mensajes o trazas. |
| `PREFIJO_CONSERVAR` | Prefijo de nombres definidos que representan excepciones de vacíos conservados | No limpiar nombres definidos indiscriminadamente; pueden expresar cambios autorizados. |

`CrearRespaldo`, `CapturarContenido` y `RestaurarContenido` utilizan libros temporales para contenido/formato/validación según el flujo. Importar respalda el rango reemplazado; mantenimiento conserva el contenido inmediatamente anterior autorizado. No son lo mismo que `GuardarRespaldoPlantilla`, que crea un **archivo completo persistente** antes de eliminar una hoja. Un snapshot temporal no garantiza recuperación después de cerrar Excel, un fallo de proceso, End o Reset.

Las rutinas conservan y procuran restaurar `Application.EnableEvents`, `ScreenUpdating` y, donde corresponde, DisplayAlerts, AutomationSecurity y EnableCancelKey. No añadir salidas que salten sus etiquetas Salida/Fallo ni su protección. **No ejecutar End/Reset durante una operación**: puede perder referencias privadas y dejar estados globales o protección sin restaurar. `RecuperarMantenimientoInterrumpido` cierra una autorización cuyo respaldo se perdió; no recupera mágicamente el contenido anterior.

### Cambiar contraseña no es editar solamente la constante

Si se autoriza una migración real: respaldo y Autoguardado pausado; finalizar permisos con clave antigua; desproteger estructura y hojas gestionadas con la antigua; revisar también el modelo PLANTILLA si está protegido; cambiar la definición central; reproteger explícitamente ese modelo con la nueva y preservar opciones/visibilidad; restablecer la protección gestionada y probar guardar/reabrir/crear/ALCON. **No guardar/cerrar a mitad.** `RestablecerProteccionAlAbrir` omite el modelo, así que su contraseña no se migra automáticamente. No cambiar por error las claves independientes de apertura del archivo, proyecto VBA o catálogos externos; los respaldos antiguos conservan sus protecciones anteriores. Esta entrega no cambia ninguna contraseña ni crea una utilidad de migración.

## 5. Depuración mínima y segura

1. Usar una copia del libro; abrir el editor desde Programador → Visual Basic si Alt+F11 no responde.
2. Buscar la rutina por nombre. F9 alterna un punto de interrupción; F8 avanza una instrucción cuando la ejecución está detenida. Si el teclado del portátil intercepta las teclas, usar Fn cuando corresponda o las opciones equivalentes de punto de interrupción/paso a paso del menú **Depuración**.
3. Preferir un punto antes de la mutación y continuar hasta una cancelación normal; no usar End/Reset como sustituto del flujo de salida.
4. Abrir **Ver → Ventana Inmediato** (Ctrl+G dentro del editor). Se puede consultar, sin cambiar nada:

```vb
? Application.EnableEvents
? Application.ScreenUpdating
? ThisWorkbook.ProtectStructure
? ActiveSheet.ProtectContents
```

`ProtectContents=True` no demuestra que todas las celdas estén bloqueadas: las entradas autorizadas se gobiernan también por Locked. No imprimir la constante de clave ni datos de cuentas. Copiar una traza solo desde Inmediato; Ctrl+A dentro del módulo seleccionaría su código.

Si aparece un error, anotar procedimiento, paso/contexto y número exacto. No borrar validaciones ni activar eventos a ciegas para forzar que continúe. Si hubo una interrupción forzada, detener el trabajo y revisar protección/estados antes de guardar; no dar por hecho que simplemente cerrar/reabrir el libro restablecerá un estado global de Excel que quedó desactivado.

## 6. Lista mínima de regresión con el usuario

Registrar por prueba: copia usada, escenario, resultado esperado/observado, mensaje, duración si interesa y estado de protección/eventos al finalizar. Los casos siguientes son **pendientes**, no pruebas ya aprobadas.

| Flujo | Comprobaciones mínimas pendientes |
|---|---|
| Entrada operativa | Oficina/fecha/NIT/valor/tipo/red/V válidos e inválidos; corrección automática de Red; errores rojos; H especial e I:T en filas nuevas y excepciones conservadas. |
| Importación | Valor simple y pegado; rangos A:G/U:V; invasión bloqueada antes de pegar; sobrescritura confirmada; Cancelar sin pérdida; error/interrupción manejada y revisión del rango recuperado. |
| MENÚ | Bloqueo inicial completo; clave incorrecta/cancelada/correcta; A/B solo con permiso; mantenimiento A:L; salir de pestaña Sí/No; proteger y guardar finalizan autorización. |
| Filas MENÚ | Preparar después de A/B; no insertar; borrar última con/sin cuentas; fila 10 y asociación protegidas; conservar borde superior según regla. |
| Crear plantilla | Modelo y K10/L10 correctos; cuentas completas; secuencia sin huecos; cancelar; error sin copia incoherente; protección/visibilidad restauradas. |
| Eliminar plantilla | Primera pasada No; luego copia respaldada con/sin datos; última >01; dependencias bloqueadas; fila MENÚ conservada y K/L provisionales; abrir copia del respaldo y verificar datos. |
| Mantenimiento | Conservar y recuperar inmediatamente anterior; cambios autorizados persisten al guardar/reabrir; modelo no se altera; salidas de error seguras. |
| Guardar/reabrir | Guardar desde MENÚ y operativa; inválidos impiden guardar; corregir y guardar; reabrir con macros y comprobar protecciones/entradas/estructura. |
| Exportaciones | Reclasificación Hoja1/Hoja2 y ALCON con/sin datos; cancelar guardado de salida; verificar contenido y retorno/visibilidad/protección del libro origen. |

### Pendientes y siguiente fase

- Regresiones **end-to-end reales** y volumen del usuario; no hay certificación de producción ni compilación VBA realizada por esta entrega.
- Faltan los libros vinculados **Saldos Contables GSMAP.xlsx** y **Cuentas Reclasificadoras Cierre Trimestral.xlsx** y su acceso autorizado. No romper vínculos, descifrarlos, ocultar errores ni convertir fórmulas en valores para simular resultados correctos.
- `EsNitValido` conserva una restricción de **100 caracteres** que no está alineada con la regla funcional de dígitos sin ese límite. **No se corrigió** en esta reorganización; requiere decisión/cambio separado.
- Solo después de verificar esos flujos corresponde acordar interfaces y pruebas para una eventual separación en módulos. No trasladar estado privado ni cambiar Public/Private como atajo. **No hay seis módulos nuevos ni una arquitectura ya migrada.**

Para revertir esta entrega, restaurar el contenido completo de [Modulo-4-v7-optimizado.txt](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modulo-4-v7-optimizado.txt) en el mismo Módulo 4, sin mezclar fragmentos. No requiere cambiar Módulo 3/ThisWorkbook ni datos del libro. Esto revierte el orden/documentación, no recupera datos eliminados por operaciones posteriores.

---

## 7. Referencia completa: 102 procedimientos

Las líneas corresponden a **Modulo-4-v7-organizado.txt entregado**. Usar Ctrl+F por nombre si se añaden comentarios o se instala otra versión. “Pública implícita” conserva la visibilidad existente por ausencia de Private/Public explícito; no es una invitación a usar auxiliares como botones.

### G1 — Acceso a hojas y utilidades de diseño (8)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `ObtenerHojaMenu` | Pública implícita / Function | 32 |
| `ObtenerHojaPlantilla` | Pública implícita / Function | 53 |
| `EsHojaTecnica` | Pública implícita / Function | 72 |
| `EsMenu` | Privada / Function | 85 |
| `ExisteHoja` | Pública implícita / Function | 89 |
| `CopiarAnchosColumnas` | Pública implícita / Sub | 106 |
| `CopiarAlturasFilas` | Pública implícita / Sub | 119 |
| `CopiarFormasZonaSuperior` | Pública implícita / Sub | 132 |

### G2 — Reglas y validación de datos (22)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `ValidarTodo_ReglasNegocio` | Pública implícita / Sub | 174 |
| `ValidarHojaTrabajo` | Pública implícita / Function | 190 |
| `ValidarTodasLasHojasTrabajo` | Pública implícita / Function | 212 |
| `RevisarFilas` | Privada / Function | 221 |
| `EsCodigoOficinaValido` | Pública implícita / Function | 297 |
| `EsFechaOrigenPartidaValida` | Pública implícita / Function | 306 |
| `EsTipoClienteValido` | Pública implícita / Function | 310 |
| `EsNitValido` | Pública implícita / Function | 324 |
| `EsNombreClienteValido` | Pública implícita / Function | 355 |
| `EsOficinaReclasificacionValida` | Pública implícita / Function | 371 |
| `EsDescripcionPartidaValida` | Pública implícita / Function | 392 |
| `EsValorPartidaValido` | Pública implícita / Function | 408 |
| `EsPartidaReclasificarValida` | Pública implícita / Function | 413 |
| `FechaEntradaValida` | Privada / Function | 429 |
| `ValorEntradaValido` | Privada / Function | 442 |
| `HayDatosEnFila` | Privada / Function | 466 |
| `LimpiarTextoValidacion` | Pública implícita / Function | 478 |
| `EscribirCambios` | Privada / Sub | 505 |
| `MarcarErrores` | Privada / Sub | 528 |
| `LimpiarColoresZonaEditable` | Pública implícita / Sub | 549 |
| `QuitarRojoEnRango` | Privada / Sub | 559 |
| `ExistenErroresMarcadosEnRojo` | Pública implícita / Function | 573 |

### G3 — Validaciones de Excel y catálogo de redes (19)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `ConfigurarValidacionesEntradaHojaActiva` | Pública implícita / Sub | 586 |
| `LimpiarValidacionesEntradaHojaActiva` | Pública implícita / Sub | 596 |
| `ConfigurarValidacionesEntradaEnHoja` | Pública implícita / Sub | 618 |
| `RestaurarValidacionesBasicasEnHoja` | Pública implícita / Sub | 625 |
| `ConfigurarReglaEnRango` | Privada / Sub | 630 |
| `ConfigurarCodigoOficina` | Pública implícita / Function | 680 |
| `ConfigurarFechaOrigenPartida` | Pública implícita / Function | 686 |
| `ConfigurarTipoCliente` | Pública implícita / Function | 692 |
| `ConfigurarOficinaReclasificacion` | Pública implícita / Function | 698 |
| `ConfigurarValorPartida` | Pública implícita / Function | 704 |
| `ConfigurarPartidaReclasificar` | Pública implícita / Function | 710 |
| `ObtenerNitRed` | Pública implícita / Function | 716 |
| `ObtenerOficinaRed` | Pública implícita / Function | 752 |
| `EsNombreRedCatalogo` | Pública implícita / Function | 789 |
| `EsNitRedCatalogo` | Pública implícita / Function | 799 |
| `AplicarListaRedesEnCelda` | Pública implícita / Sub | 836 |
| `QuitarListaRedesEnCelda` | Pública implícita / Sub | 851 |
| `ActualizarListaRedSegunTipoCliente` | Pública implícita / Sub | 855 |
| `PrepararListasRed` | Privada / Sub | 859 |

### G4 — Importación de valores (1)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `ImportarDatosDesdeArchivoExcel` | Pública implícita / Sub | 888 |

### G5 — Entradas, última fila y fórmulas calculadas (8)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `FilaTieneDatosEntrada` | Pública implícita / Function | 1050 |
| `ObtenerUltimaFilaEntrada` | Pública implícita / Function | 1060 |
| `ObtenerUltimaFilaParaRevision` | Pública implícita / Function | 1072 |
| `ActualizarEntradaEnHoja` | Pública implícita / Sub | 1080 |
| `PrepararFilasEntrada` | Pública implícita / Sub | 1097 |
| `CompletarFormulasZonaGris` | Privada / Sub | 1127 |
| `CompletarColumnaCalculada` | Privada / Sub | 1136 |
| `ActualizarRespaldoDespuesDeEntrada` | Pública implícita / Sub | 1184 |

### G6 — Protección y autorizaciones (11)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `ProtegerTodasLasHojasTrabajo` | Pública implícita / Sub | 1208 |
| `ProtegerHojaTrabajo` | Pública implícita / Sub | 1221 |
| `ConfigurarBloqueoHoja` | Privada / Sub | 1226 |
| `RestaurarEstadoHoja` | Privada / Sub | 1248 |
| `DesbloquearHoja` | Pública implícita / Sub | 1272 |
| `EstaEnMantenimiento` | Pública implícita / Function | 1355 |
| `RestaurarBloqueoFuncional` | Privada / Sub | 1360 |
| `FinalizarMantenimiento` | Pública implícita / Sub | 1384 |
| `VolverAProtegerHoja` | Pública implícita / Sub | 1408 |
| `RestablecerProteccionAlAbrir` | Pública implícita / Sub | 1427 |
| `ProtegerEstructuraLibro` | Pública implícita / Sub | 1441 |

### G7 — Respaldos, mantenimiento y excepciones (12)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `ZonaCalculada` | Privada / Function | 1450 |
| `VerificarContenidoRestaurable` | Privada / Sub | 1458 |
| `CrearRespaldo` | Privada / Function | 1481 |
| `TextosLiterales` | Privada / Function | 1502 |
| `CapturarContenido` | Privada / Sub | 1517 |
| `RestaurarContenido` | Privada / Sub | 1560 |
| `GestionarCambioMantenimiento` | Pública implícita / Sub | 1610 |
| `ExcluirZona` | Privada / Function | 1646 |
| `AgregarExcepcion` | Privada / Sub | 1679 |
| `RegistrarVaciosConservados` | Privada / Sub | 1696 |
| `VerificarZonaMantenimiento` | Privada / Sub | 1729 |
| `RecuperarMantenimientoInterrumpido` | Pública implícita / Function | 1748 |

### G8 — MENÚ: cuentas, filas y vínculos (12)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `MenuAutorizado` | Pública implícita / Function | 1780 |
| `AutorizarCuentasMenu` | Privada / Function | 1788 |
| `AvisarSalidaMenu` | Pública implícita / Sub | 1819 |
| `AgregarFilaMenu` | Pública implícita / Sub | 1845 |
| `EliminarUltimaFilaMenu` | Pública implícita / Sub | 1942 |
| `FilaMenuCompleta` | Pública implícita / Function | 2093 |
| `VerificarModeloMenu` | Privada / Sub | 2119 |
| `VerificarCuentasFilaMenu` | Privada / Sub | 2128 |
| `VincularFilaMenu` | Privada / Sub | 2137 |
| `AplicarFormulasEncabezadosDesdeMenu` | Pública implícita / Sub | 2155 |
| `HojaAsociadaFilaMenu` | Privada / Function | 2171 |
| `BordeInferiorFilaMenu` | Privada / Sub | 2198 |

### G9 — Plantillas y comprobación de dependencias (9)

| Procedimiento | Visibilidad / tipo | Línea |
|---|---|---:|
| `CrearNuevaHojaDesdePlantilla` | Pública implícita / Sub | 2243 |
| `EliminarUltimaPlantilla` | Pública implícita / Sub | 2359 |
| `ContarHojasTrabajoCreadas` | Pública implícita / Function | 2455 |
| `VerificarSecuenciaPlantillas` | Privada / Function | 2459 |
| `GuardarRespaldoPlantilla` | Privada / Function | 2482 |
| `VerificarDependenciasPlantilla` | Privada / Sub | 2506 |
| `VerificarTextoDependencia` | Privada / Sub | 2592 |
| `VerificarValidacionesDependientes` | Privada / Sub | 2625 |
| `VerificarFormulasDependientes` | Privada / Sub | 2696 |
