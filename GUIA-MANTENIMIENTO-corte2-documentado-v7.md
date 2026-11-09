# Guía de mantenimiento — v7 corte 2 con base preparada hasta 1000

## Ruta rápida

**Conservar la base H12:T1000:** la limpieza automática solo puede intervenir crecimiento sobrante desde la fila 1001. La fila 1000 no es un límite de datos; se pueden usar 300.000 registros o más, dentro de los límites de Excel. Se añade una recuperación manual puntual de huecos que dejó la edición anterior. No es otra modularización.

1. Trabajar en copia íntegra. Con el código anterior completo, finalizar mantenimiento/autorización y pausar Autoguardado. Conservar las fuentes anteriores y seguir LEEME. Si cambió la contraseña desde la base, coordinar adaptación sin publicar la clave ni reponer una antigua.
2. Si está instalada la última entrega LimpiezaZonaGris, sustituir solo TODO el contenido del Módulo 4 existente por Modulo-4-v7-corte2.txt. Módulo 1 y los tres internos de este paquete son byte idénticos a esa entrega y no requieren reinstalación si ya coinciden. Mantener nombres reales, incluido acento, y no duplicar módulos.
3. Si aún usa una entrega anterior a LimpiezaZonaGris, instalar también el BorrarDatos compatible incluido en Modulo-1-limpieza-zona-gris.txt y los tres módulos internos que falten. No sobrescribir otras rutinas/cambios locales de Módulo 1: comparar e integrar su única Sub. La pareja compatible 1/4 sigue siendo necesaria para que BorrarDatos limpie crecimiento; no mezclar bases históricas.
4. Durante una instalación parcial NO ejecutar, guardar, cerrar ni cambiar de hoja. Al terminar, compilar y probar en copia. No se ha compilado ni ejecutado VBA/Excel durante la verificación de estos TXT. No cambiar ThisWorkbook, Módulos 2/3, figuras ni contraseña.
5. Si hay huecos heredados dentro de H12:T1000, ejecutar la recuperación puntual descrita abajo. No se ejecuta al abrir, guardar ni con cada edición.

### Recuperar una base que ya tiene huecos

1. Activar la hoja operativa afectada, no MENÚ, DATOS, ALCON ni el modelo PLANTILLA.
2. Ejecutar DesbloquearHoja, introducir la contraseña y elegir **No: Otra acción (mantenimiento del área funcional)**. Se reutiliza esta autorización; la recuperación no pide otra contraseña.
3. Abrir Alt+F8, elegir **RecuperarBaseZonaGris** y ejecutar. Basta una ejecución por hoja afectada; se puede repetir y solo se rellenan vacíos elegibles.
4. La macro valida todos los modelos I:T y las estructuras del rango antes de escribir, respalda solo H12:T1000, repone vacíos y sincroniza el respaldo de mantenimiento. H mantiene su fórmula especial. Respeta fórmulas/constantes existentes y vacíos expresamente conservados. No crea excepciones para las fórmulas recuperadas automáticamente.
5. Revisar, finalizar con VolverAProtegerHoja y guardar cuando las validaciones del libro lo permitan. Sin mantenimiento vigente de esa misma hoja, la macro solo indica cómo habilitarlo y no modifica contenido.

No reconstruye fórmulas de un modelo incompleto ni interviene en matrices, derrames, combinaciones o tablas: solicita revisión antes de escribir. Una fórmula existente pero incorrecta no se reemplaza. Ante un fallo manejado intenta restaurar el contenido previo desde un respaldo temporal local y volver a sincronizar el mantenimiento; no garantiza atomicidad frente a End/Reset/cierre forzado. Si no puede recuperar el estado fiable, intenta cerrar mantenimiento y deja visible el respaldo temporal para revisión. El éxito mantiene la autorización hasta proteger o finalizar según las reglas existentes. No repara vínculos externos; resolver sus avisos con los archivos autorizados.

### Criterio autorizado y límites de limpieza

- La base H12:T1000 no se borra por la limpieza automática, incluso al borrar A12, una fila intermedia o todas las entradas con BorrarDatos. No se rellenan sus 12.857 celdas con cada edición; solo se conserva lo existente. Los huecos previos se atienden con la macro manual, no con otro barrido permanente.
- Crecimiento desde 1001: limpiar solo el tramo final afectado y realmente vacío en A:G/U:V. Ejemplos 1050→1004 limpia1005:1050 y conserva H1005;1050→900 conserva la base y limpia1001:1050 afectado. Datos que crecen por encima de 1000 siguen preparando sus fórmulas normalmente.
- Solo fórmulas idénticas al modelo y sin excepción conservada se consideran automáticas; se conserva el criterio histórico ya autorizado. Las decisiones históricas idénticas sin marca no pueden reconstruirse. Desde la entrega anterior toda zona aceptada en mantenimiento recibe las excepciones existentes, sin crear marcas por fórmula automática.
- Se preservan constantes, fórmulas diferentes, decisiones/vacíos marcados y nuevos cambios aceptados aunque coincidan con modelo. No se comparan resultados de cálculo ni se normalizan literales de texto. Una fila con cualquier entrada o un hueco con registros debajo no se limpia.
- Módulo 1 se mantiene idéntico: borra sus entradas completas, captura el último registro anterior y llama al mismo auxiliar; ahora hereda la exclusión de la base sin otro cambio. Los tres componentes internos y ThisWorkbook no cambian.
- No hay limpieza inicial del sobrante histórico no afectado. No se borran formatos, validaciones, comentarios ni filas físicas ni el modelo PLANTILLA. Se mantienen las omisiones de seguridad de tablas/matrices/derrames/combinaciones o excepciones ilegibles.
- El trabajo automático sigue acotado al cambio y por bloques; la validación previa conserva su alcance sobre Target. Borrar columnas completas puede seguir siendo costoso. No se aporta medición nueva ni se garantiza una reducción concreta de tamaño o tiempo.

## 1. Mapa de responsabilidades del corte 2

| Componente | Rutinas | Responsabilidad |
|---|---:|---|
| [Módulo 4 existente](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Base1000/Modulo-4-v7-corte2.txt) | 84 | Entradas públicas, configuración central, tres objetos privados, autorizaciones, importación, MENÚ, plantillas y coordinación. |
| [NeteoValidaciones](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Base1000/NeteoValidaciones.txt) | 9 | Validación por bloques, fecha/valor, colores y listas/reglas. Archivo byte idéntico al corte 2 documentado; sin cambios en esta entrega. |
| [NeteoDependencias](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Base1000/NeteoDependencias.txt) | 5 | Inspección de referencias y resta rectangular. Archivo byte idéntico al corte 2 documentado; sin cambios en esta entrega. |
| [NeteoRespaldos](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Base1000/NeteoRespaldos.txt) | 6 | Crear respaldo temporal, capturar/restaurar contenido, comprobar compatibilidad del rango y copiar el archivo antes de eliminar. Sin estado de autorización. |

**104 rutinas en los cuatro componentes de Módulo 4 (84/9/5/6), más BorrarDatos idéntico en Módulo 1.** Se conservan las 65 firmas públicas originales, el auxiliar compartido y se añade la macro pública RecuperarBaseZonaGris. Solo cambian LimpiarFormulasFinalesVacias y el auxiliar privado CompletarColumnaCalculada entre las 83 rutinas previas; las otras 81, incluida AgregarFilaMenu, permanecen byte idénticas. Se añade una constante privada de base de 1000 filas, sin otro estado mutable.

### Fronteras internas y propiedad del estado

- NeteoRespaldos lleva `Option Explicit` y `Option Private Module`. Public dentro del proyecto: `CrearRespaldo`, `CapturarContenido`, `RestaurarContenido`, `VerificarContenidoRestaurable`, `GuardarRespaldoPlantilla`. `TextosLiterales` sigue Private y no se duplica.
- `CapturarContenido` y `RestaurarContenido` reciben tres ByVal finales: si la hoja es MENÚ, primera fila de entradas y tamaño de bloque. Módulo 4 evalúa su `EsMenu` privado sobre la misma hoja del rango y transmite sus constantes originales. Dentro se mantiene el límite 10 para MENÚ y el valor original de `PRIMERA_FILA` para operativas (12); no se copian números a cada llamada.
- La propiedad y vida del respaldo siguen en Módulo 4: `hojaMantenimiento`, `respaldoMantenimiento`, `hojaCuentasMenu` permanecen juntos y Private. No se autoriza, conserva ni termina mantenimiento desde el componente de respaldos.
- `CLAVE_HOJA`, `PRIMERA_FILA`, `TAMANO_BLOQUE` y `PREFIJO_CONSERVAR` no se mueven ni duplican. `ULTIMA_FILA_BASE` es la nueva constante privada1000: preparación conservada, nunca capacidad máxima de datos. No publicar la clave ni cambiarla aisladamente. Los componentes no usan estado de autorización global nuevo.
- Los módulos de validaciones/dependencias conservan las ocho APIs internas del corte 1. `Option Private Module` permite Public dentro del proyecto, no hacia proyectos referenciantes; no transforma una variable Public en Private ni añade seguridad a Excel.
- Sigue siendo separación gradual: importación, MENÚ y protección comparten coordinación en Módulo 4. No se presentan capas completamente independientes ni se trasladan excepciones/nombres definidos por este corte.

### Seguir un respaldo sin perder su dueño

| Inicio en Módulo 4 | Auxiliar de NeteoRespaldos | Quién conserva decisiones/salida |
|---|---|---|
| `ImportarDatosDesdeArchivoExcel` | CrearRespaldo → VerificarContenidoRestaurable → CapturarContenido; RestaurarContenido en el fallo manejado | Importación conserva el rango, la confirmación, formato/validación, cierre del respaldo, protección y banderas. |
| `DesbloquearHoja` / `GestionarCambioMantenimiento` | CrearRespaldo, CapturarContenido, RestaurarContenido | Módulo 4 conserva autorización, decisión conservar/recuperar y referencia al respaldo anterior. |
| `ActualizarRespaldoDespuesDeEntrada`, filas/vínculos de MENÚ | CapturarContenido | Módulo 4 selecciona solo las celdas afectadas y sigue siendo dueño de su snapshot. |
| `VerificarZonaMantenimiento` | VerificarContenidoRestaurable | Se conservan las comprobaciones previas de combinadas, matrices, derrames y tablas, sin omitirlas. |
| `EliminarUltimaPlantilla` | GuardarRespaldoPlantilla | Se copia el archivo completo antes de la primera mutación destructiva; confirmar/borrar y restaurar estados permanece en Módulo 4. No guarda automáticamente el original. |

`TextosLiterales` conserva el tratamiento original de textos que comienzan con `=`, otros textos, escalares/matrices, booleanos y valores de error. El corte no cambia el algoritmo de Value2/FormulaR1C1 ni corrige diferencias de representación COM. El snapshot temporal y el respaldo persistente de archivo siguen siendo conceptos distintos.

La medición del usuario **48,727 s → 2,391 s** pertenece a la optimización anterior de la revisión previa a confirmar, en un caso. La limpieza actual tampoco garantiza esos tiempos ni una mejora de tamaño o corrección end-to-end sin probar el libro.

### Botones/macros: asignaciones esperadas

Este mapa muestra los procedimientos previstos, **no certifica las asignaciones reales de las figuras**: no se inspeccionó el libro instalado. Conservar nombres de módulos, macros, argumentos y visibilidad; no asignar una función auxiliar como sustituto de un botón. **No eliminar una rutina porque no se vea llamada desde el código**: puede estar asignada a una figura o invocada externamente; falta un inventario real de OnAction.

| Acción prevista | Procedimiento | Responsabilidad y hojas |
|---|---|---|
| Validar entradas | `ValidarTodo_ReglasNegocio` | Hoja operativa activa; valida, marca y aplica las correcciones existentes. |
| Importar valores | `ImportarDatosDesdeArchivoExcel` | Rango origen y destino operativo editable; sin reorganizar columnas. |
| Desbloquear/autorizar | `DesbloquearHoja` | MENÚ: cuentas A/B o mantenimiento A:L; operativa: mantenimiento A:V. |
| Volver a proteger | `VolverAProtegerHoja` | Cierra autorización/mantenimiento y restaura bloqueos funcionales. |
| Agregar fila MENÚ | `AgregarFilaMenu` | Según última información A/B, prepara un destino vacío, conserva uno completamente formulado o detiene contenido incompleto/inesperado. No crea hoja. |
| Eliminar última fila MENÚ | `EliminarUltimaFilaMenu` | Limpia A:L, sin desplazamientos; nunca fila 10; bloquea una hoja asociada. |
| Crear plantilla | `CrearNuevaHojaDesdePlantilla` | Copia PLANTILLA, crea PLANTILLA nn y enlaza MENÚ. |
| Eliminar última plantilla | `EliminarUltimaPlantilla` | Solo última numerada permitida, con clave, confirmación y respaldo completo. Conserva su fila de MENÚ. |
| Configurar validaciones, uso técnico | `ConfigurarValidacionesEntradaHojaActiva` | Repone reglas de la operativa; no sustituye las pruebas de negocio. |
| Restablecer protección, automático/técnico | `RestablecerProteccionAlAbrir` | MENÚ/operativas y estructura; no migra por sí solo la protección del modelo. |
| Recuperar base, manual y autorizado | `RecuperarBaseZonaGris` | Solo operativa activa en mantenimiento; reponer vacíos elegibles H12:T1000, sin cambios de figuras ni ejecución automática. |
| Borrar datos, Módulo 1 | `BorrarDatos` | Borra entradas completas y llama el auxiliar compartido sin modificar este módulo; hereda conservación de base de 1000 filas y limpieza solo de crecimiento afectado. |
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

Para editar cuentas se requiere autorización; la hoja sigue protegida con las celdas permitidas desbloqueadas. `AgregarFilaMenu` reutiliza autorización de cuentas o mantenimiento vigente; si no existe, solicita y comprueba la clave antes de inspeccionar cuentas o fórmulas, sin habilitar todavía A/B. Se basa en la última información A/B, no en la última fórmula preparada C:L. Por eso repetir el botón sin llenar cuentas vuelve a la misma fila, ahora sin reemplazar sus fórmulas completas. No inserta/desplaza filas ni crea operativas. `EliminarUltimaFilaMenu` no cambia: pide clave siempre, considera contenido/fórmulas A:L, protege fila 10 y rechaza una asociación existente; conserva el borde inferior de la fila superior.

### Ajuste de AgregarFilaMenu

La comprobación se limita a las diez celdas C:L del destino, **sin interpretar sus resultados** ni comparar su semántica contable con el modelo:

| Estado de C:L en el destino | Acción |
|---|---|
| Todas vacías, sin contenido ni fórmulas | Conserva el procedimiento anterior de copia de fórmulas, formato y validaciones desde C10:L10; K/L se dejan visualmente vacías para una fila nueva. |
| Las diez celdas tienen fórmulas, sin texto #REF! en ellas | Conserva fórmulas, formatos y validaciones existentes; no pega ni pregunta por reemplazo. Autoriza A/B conforme al permiso vigente, selecciona A de esa fila y avisa dónde ingresar ambas cuentas. |
| Mezcla de fórmulas y vacíos, constantes, errores constantes u otra preparación incompleta | Avisa y sale sin sobrescribir ni habilitar una nueva autorización. El permiso de cuentas o mantenimiento que ya estuviera vigente se mantiene. |
| Una fórmula contiene #REF! | Avisa y sale para revisión, aunque IFERROR oculte el error. No repara referencias ni vínculos. La búsqueda es textual y también detiene una fórmula que incluya ese texto como literal. |

Una fórmula que devuelve vacío, cero, #N/A u otro error de cálculo no se rechaza solo por ese resultado. Esto no demuestra que la fórmula sea contablemente correcta ni que sus vínculos estén disponibles. Un valor fijo, aunque parezca un resultado correcto, no cuenta como fórmula preparada. Formatos o validaciones sin contenido no cuentan como preparación: se sigue la rama vacía y se aplica el modelo como antes.

Se conservan las comprobaciones previas de K10/L10 y su referencia a PLANTILLA 01. La fila 10 es el modelo mínimo: si C10:L10 está completo se conserva, incluidos K10/L10; no se envuelven otra vez sus fórmulas. Si falta el propio modelo, el botón no puede reconstruirlo copiándose a sí mismo: se detiene para revisión. Las fórmulas modelo existentes deben gestionar la presentación vacía si todavía no está creada PLANTILLA 01.

En mantenimiento se mantiene A:L autorizado y M en adelante bloqueado. Si la fila ya estaba completa, no se vuelve a capturar su contenido en el respaldo temporal porque no se reemplazó; si se preparó una fila vacía, se conserva la captura anterior. Los bloqueos funcionales siguen ajustándose al modo vigente, sin reemplazar fórmulas, validaciones o formato visual de una fila completa.

Ejemplo: con cuentas hasta A29/B29 y fórmulas completas en C30:L30, el botón selecciona A30 y avisa «La fila 30 ya está preparada. Ingrese las cuentas en A30 y B30. No se reemplazaron sus fórmulas». Después de ingresar cuentas en A30/B30, la siguiente ejecución apunta a la fila 31. Clave incorrecta o cancelada no reemplaza contenido ni habilita permisos. Esta lógica no agrega un cuadro de confirmación nuevo.

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
| `PREFIJO_CONSERVAR` | Prefijo de nombres definidos que representan zonas conservadas; antes solo vacíos, ahora todo contenido aceptado | No limpiar nombres definidos indiscriminadamente; pueden expresar cambios autorizados. |

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

## 6. Lista de verificación en el libro real

No basta con que compile o aparezca un mensaje correcto. Comparar con una copia anterior equivalente y registrar resultado por caso.

| Flujo | Caso mínimo y resultado a comprobar |
|---|---|
| Instalación | Cinco fuentes completas compatibles; instalar solo Módulo 4 si la pareja anterior ya coincide;65 APIs originales más el auxiliar compartido y la recuperación manual; proyecto compilado sin duplicados ni referencias inaccesibles. Probar figuras reales: su asignación no fue inventariada. |
| Entrada/validación | Fecha, NIT, valor distinto de cero, V obligatoria, Cliente/Red y catálogo; valores válidos/erróneos/pegado múltiple. Mismos mensajes y marcas rojas; H especial e I:T recuperadas sin sobrescribir excepciones. |
| Importar | Valores sin reorganizar desde origen válido; límite de zona A:G/U:V; intento de propio libro; confirmación de reemplazo con No; cancelación/fallo y recuperación del rango. No perder datos ni permitir guardado inválido. |
| MENÚ/protección | Bloqueo predeterminado total; clave incorrecta/cancelar sin cambios; cuentas solo A/B y mantenimiento A:L. Agregar: vacío copia modelo; completo conserva C:L y selecciona A; incompleto/constante/#REF! detiene sin nuevo permiso. Eliminar fila, fila 10, asociación, conservar/restaurar contenido y aviso al salir mantienen sus reglas. |
| Crear plantilla | Modelo PLANTILLA y fila MENÚ válidos; dependencia/gap/modelo inválido bloquea; clave y confirmación según permiso actual; error/cancelación restaura visibilidad, estructura y eventos. |
| Eliminar plantilla | Primera comprobación responder No: nada se borra. Luego, solo en copia desechable con respaldo confirmado: última adicional, nunca 01, dependencias bloquean; A/B y C:J conservadas, K/L provisionales y estado correcto. No prometer recuperación automática tras borrado efectivo. |
| Guardar/reabrir | Guardar desde MENÚ detecta errores operativos y bloquea; corregir y guardar. Reabrir: autorizaciones finalizadas, bloqueos/estructura y fórmulas conservadas; eventos/pantalla normales. |
| Salidas contables | Comparar Reclasificación y ALCON con las salidas esperadas, incluida Hoja1 sin datos. No se modificaron Módulos 2/3; comprobar compatibilidad real con la nueva pareja 1/4. |

### Pendientes y límites

- Falta compilación y regresión completa del proyecto realmente instalado; las verificaciones de archivos no ejecutan VBA, figuras, eventos ni libros externos.
- Siguen pendientes los vínculos a Saldos Contables GSMAP.xlsx y Cuentas Reclasificadoras Cierre Trimestral.xlsx; no se dispone de acceso/contraseñas autorizados. No romper vínculos ni intentar descifrarlos por esta entrega.
- El límite actual de 100 caracteres de NIT sigue siendo una inconsistencia pendiente de decisión; no se cambió silenciosamente.
- La limpieza adicional de H fue cancelada por el usuario y no forma parte del paquete.
- Nuevos cortes (importación, MENÚ, protección/plantillas) solo después de aprobar estas regresiones, con interfaces de estado privadas bien delimitadas; no volver públicos los estados ni repartir código por número de líneas.

### Pruebas focales añadidas para este corte

- Capturar/conservar/recuperar una celda con texto literal iniciado en =, una fórmula, un valor vacío, booleano o error, sin conversiones no deseadas; comparar con el corte 1 en el libro real.
- MENÚ: encabezados hasta fila 9 y datos desde fila 10. Operativas: encabezados hasta fila 11 y datos desde fila 12. Probar una celda, un bloque y varios bloques atravesando el límite.
- Importación: rechazar rango incompatible antes de reemplazar; cancelar sin perder contenido; probar recuperación ante fallo manejado en copia. No certificar reversión por solo ver un mensaje.
- Mantenimiento: conservar y luego recuperar una edición posterior respetando cambios autorizados previos, cerrar autorización y comprobar protección/eventos/pantalla.
- Borrar plantilla: primera revisión con No; después, solo en copia desechable, comprobar archivo de respaldo visible y válido antes de eliminación, controles de referencias y K/L de MENÚ.

En el corte 2 se comprobaron fuentes y modelos Python de límites, bloques y literales. El ajuste anterior de MENÚ añadió contratos de clasificación/autorización; esta entrega agrega limpieza acotada y conservación, con preservación de rutinas ajenas, identidad de internos y lectura del ZIP: **no ejecución del motor VBA ni de Excel**. Compilar y probar en copia los estados vacío/completo/incompleto, resultados vacíos o con error, clave cancelada/incorrecta, fila 10, mantenimiento, protección y figuras reales. No presentar estas comprobaciones como regresión funcional completa del libro.

## Pruebas específicas de esta entrega

1. Con base H12:T1000 completa, ingresar y borrar A12; repetir en 800 y1000. No debe borrar sus fórmulas existentes.
2. Datos hasta 1050: reducir a 1004 conserva base y H1005 y limpia solo crecimiento automático1005:1050 afectado. Reducir a 900 conservaH12:T1000 y limpia1001:1050 afectado.
3. BorrarDatos mantiene toda la base existente y las excepciones; no rellena los huecos heredados ni limpia sobrante histórico no afectado. Módulo 1 debe seguir idéntico.
4. Ingresar datos por encima de 1000, incluso el volumen habitual de 300.000 o más: preparar filas afectadas normalmente. La constante1000 no limita última fila ni validación/importación.
5. Recuperación sin permiso: mensaje de orientación sin cambios. Con mantenimiento de la hoja, comprobar vacíos elegibles, fórmulas existentes, constantes, vacíos conservados, fórmulaH, formato/validaciones y no creación de marcas para relleno automático.
6. Modelo incompleto o estructuras especiales: detener antes de escribir. En copia, verificar errores manejados, restauración de contenido previo, mantenimiento/snapshot, protección/eventos/pantalla y cierre del respaldo temporal.
7. Después de recuperar, una edición en mantenimiento con No debe respetar la base recuperada en el snapshot. Finalizar, guardar/reabrir y comprobar excepciones. Probar regresión de MENÚ/plantillas, incluidos permisos y AgregarFilaMenu.

Las pruebas automáticas internas son contratos de fuente y modelos de política, no compilación/ejecución del motor VBA o eventos reales. Las expectativas antiguas de limpiarI12:T12 al quedar sin entradas se sustituyen explícitamente por la conservación de la base aprobada, sin presentarlas como pruebas antiguas verdes.

## Reversión de esta entrega

Si ya tenía instalada LimpiezaZonaGris, reponer solo su Módulo 4 anterior y compilar; Módulo 1 y los tres internos son idénticos. Si necesita volver a una versión anterior a esa pareja, reponer el conjunto compatible de aquel punto, no mezclar fuentes. Restaurar código no recupera contenido borrado; conservar la copia íntegra previa. La versión anterior vuelve a permitir limpieza de la base, por lo que no debe confundirse reversión con reparación.

## Catálogo de líneas de la entrega actual

### Modulo-4-v7-corte2.txt

| Procedimiento | Línea |
|---|---:|
| `ObtenerHojaMenu` | 29 |
| `ObtenerHojaPlantilla` | 50 |
| `EsHojaTecnica` | 69 |
| `EsMenu` | 82 |
| `ExisteHoja` | 86 |
| `CopiarAnchosColumnas` | 103 |
| `CopiarAlturasFilas` | 116 |
| `CopiarFormasZonaSuperior` | 129 |
| `ValidarTodo_ReglasNegocio` | 171 |
| `ValidarHojaTrabajo` | 187 |
| `ValidarTodasLasHojasTrabajo` | 209 |
| `EsCodigoOficinaValido` | 218 |
| `EsFechaOrigenPartidaValida` | 227 |
| `EsTipoClienteValido` | 231 |
| `EsNitValido` | 245 |
| `EsNombreClienteValido` | 276 |
| `EsOficinaReclasificacionValida` | 292 |
| `EsDescripcionPartidaValida` | 313 |
| `EsValorPartidaValido` | 329 |
| `EsPartidaReclasificarValida` | 334 |
| `LimpiarTextoValidacion` | 350 |
| `LimpiarColoresZonaEditable` | 377 |
| `ExistenErroresMarcadosEnRojo` | 387 |
| `ConfigurarValidacionesEntradaHojaActiva` | 400 |
| `LimpiarValidacionesEntradaHojaActiva` | 410 |
| `ConfigurarValidacionesEntradaEnHoja` | 432 |
| `RestaurarValidacionesBasicasEnHoja` | 439 |
| `ConfigurarCodigoOficina` | 444 |
| `ConfigurarFechaOrigenPartida` | 450 |
| `ConfigurarTipoCliente` | 456 |
| `ConfigurarOficinaReclasificacion` | 462 |
| `ConfigurarValorPartida` | 468 |
| `ConfigurarPartidaReclasificar` | 474 |
| `ObtenerNitRed` | 480 |
| `ObtenerOficinaRed` | 516 |
| `EsNombreRedCatalogo` | 553 |
| `EsNitRedCatalogo` | 563 |
| `AplicarListaRedesEnCelda` | 600 |
| `QuitarListaRedesEnCelda` | 615 |
| `ActualizarListaRedSegunTipoCliente` | 619 |
| `ImportarDatosDesdeArchivoExcel` | 631 |
| `FilaTieneDatosEntrada` | 799 |
| `ObtenerUltimaFilaEntrada` | 809 |
| `ObtenerUltimaFilaParaRevision` | 821 |
| `ActualizarEntradaEnHoja` | 829 |
| `PrepararFilasEntrada` | 857 |
| `CompletarFormulasZonaGris` | 888 |
| `CompletarColumnaCalculada` | 897 |
| `LimpiarFormulasFinalesVacias` | 955 |
| `RecuperarBaseZonaGris` | 1104 |
| `ActualizarRespaldoDespuesDeEntrada` | 1196 |
| `ProtegerTodasLasHojasTrabajo` | 1220 |
| `ProtegerHojaTrabajo` | 1233 |
| `ConfigurarBloqueoHoja` | 1238 |
| `RestaurarEstadoHoja` | 1260 |
| `DesbloquearHoja` | 1284 |
| `EstaEnMantenimiento` | 1367 |
| `RestaurarBloqueoFuncional` | 1372 |
| `FinalizarMantenimiento` | 1396 |
| `VolverAProtegerHoja` | 1420 |
| `RestablecerProteccionAlAbrir` | 1439 |
| `ProtegerEstructuraLibro` | 1454 |
| `ZonaCalculada` | 1463 |
| `GestionarCambioMantenimiento` | 1475 |
| `AgregarExcepcion` | 1512 |
| `RegistrarVaciosConservados` | 1529 |
| `VerificarZonaMantenimiento` | 1555 |
| `RecuperarMantenimientoInterrumpido` | 1574 |
| `MenuAutorizado` | 1606 |
| `AutorizarCuentasMenu` | 1614 |
| `AvisarSalidaMenu` | 1645 |
| `AgregarFilaMenu` | 1671 |
| `EliminarUltimaFilaMenu` | 1783 |
| `FilaMenuCompleta` | 1934 |
| `VerificarModeloMenu` | 1960 |
| `VerificarCuentasFilaMenu` | 1969 |
| `VincularFilaMenu` | 1978 |
| `AplicarFormulasEncabezadosDesdeMenu` | 1996 |
| `HojaAsociadaFilaMenu` | 2012 |
| `BordeInferiorFilaMenu` | 2039 |
| `CrearNuevaHojaDesdePlantilla` | 2084 |
| `EliminarUltimaPlantilla` | 2200 |
| `ContarHojasTrabajoCreadas` | 2299 |
| `VerificarSecuenciaPlantillas` | 2303 |

### Modulo-1-limpieza-zona-gris.txt

| Procedimiento | Línea |
|---|---:|
| `BorrarDatos` | 3 |

### NeteoValidaciones.txt

| Procedimiento | Línea |
|---|---:|
| `RevisarFilas` | 11 |
| `FechaEntradaValida` | 87 |
| `ValorEntradaValido` | 100 |
| `HayDatosEnFila` | 124 |
| `EscribirCambios` | 136 |
| `MarcarErrores` | 159 |
| `QuitarRojoEnRango` | 180 |
| `ConfigurarReglaEnRango` | 194 |
| `PrepararListasRed` | 244 |

### NeteoDependencias.txt

| Procedimiento | Línea |
|---|---:|
| `VerificarDependenciasPlantilla` | 8 |
| `VerificarTextoDependencia` | 95 |
| `VerificarValidacionesDependientes` | 128 |
| `VerificarFormulasDependientes` | 199 |
| `ExcluirZona` | 271 |

### NeteoRespaldos.txt

| Procedimiento | Línea |
|---|---:|
| `CrearRespaldo` | 11 |
| `TextosLiterales` | 32 |
| `CapturarContenido` | 49 |
| `RestaurarContenido` | 94 |
| `VerificarContenidoRestaurable` | 140 |
| `GuardarRespaldoPlantilla` | 165 |

