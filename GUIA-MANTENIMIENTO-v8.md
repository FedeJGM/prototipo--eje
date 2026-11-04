# Guía de mantenimiento — v7 corte 2 con limpieza final de zona gris

## Ruta rápida

**Limpieza final acotada:** al reducir entradas, retirar únicamente fórmulas automáticas del tramo final afectado y vacío, conservando H de la siguiente fila y las decisiones de mantenimiento. Se modifica Módulo 4 y se entrega Módulo 1 compatible. No es otra modularización.

1. Trabajar en una copia íntegra. Con el código anterior completo, finalizar mantenimiento/autorización y pausar Autoguardado. Conservar la pareja de fuentes anterior. Si cambió CLAVE_HOJA, coordinar la adaptación sin publicar la clave ni reponer una antigua.
2. Con el corte 2 modular instalado, reemplazar TODO el contenido del Módulo 4 existente por Modulo-4-v7-corte2.txt y TODO el Módulo 1 de BorrarDatos por Modulo-1-limpieza-zona-gris.txt. Si su Módulo 1 contiene otras rutinas/cambios, no sobrescribirlo: comparar e integrar la única Sub BorrarDatos. Mantener los nombres reales, incluido acento. Instalar ambos antes de ejecutar.
3. Los tres módulos internos adjuntos son byte idénticos a la entrega anterior; no es necesario sustituirlos si ya coinciden. Si aún no están instalados, seguir LEEME para instalar el conjunto completo. No duplicar Módulo 4 ni BorrarDatos.
4. Durante la sustitución parcial NO ejecutar macros, guardar, cerrar ni cambiar de hoja. Compilar el proyecto cuando todo esté completo y probar en copia los casos de esta guía. No se ha compilado ni ejecutado VBA/Excel durante la verificación de estos TXT.
5. No sustituir ThisWorkbook ni Módulos 2/3, no cambiar figuras ni la contraseña por este ajuste.

### Criterio autorizado y límites de la limpieza

- En el tramo final afectado y realmente vacío de A:G/U:V, las fórmulas históricas idénticas al modelo y sin marca se consideran automáticas. Es el criterio aprobado: no existe información histórica para distinguir una coincidencia previamente aceptada de una fórmula generada automáticamente.
- Desde esta entrega, conservar un cambio en mantenimiento registra toda la zona aceptada, incluidas fórmulas y constantes; RegistrarVaciosConservados mantiene su nombre histórico, reutiliza las excepciones y descuenta rangos ya registrados. No crea nombres por cada fórmula automática.
- Fórmulas diferentes, constantes, excepciones/vacíos conservados y las nuevas decisiones autorizadas, aunque coincidan con el modelo, se preservan. No se comparan resultados de cálculo para decidir qué borrar.
- Comparación conservadora y exacta con FormulaR1C1 del modelo I:T y la fórmula especial H: diferencias de escritura se conservan, no se normalizan literales de texto ni se reconstruyen vínculos.
- Una fila intermedia vacía o una fila con cualquier entrada (incluido F/U/V, cero, error o una fórmula que devuelva vacío) no se limpia. Se mantiene H de la siguiente fila; si no quedan entradas, H12, respetando sus excepciones.
- Módulo 4 valida primero las filas editadas y prepara H:T solo hasta la última entrada real; no rellena las filas finales para borrarlas después. El nuevo auxiliar limita su trabajo al intervalo afectado y al contenido existente, por bloques y rangos contiguos; no recorre celda por celda toda la hoja ni todas las hojas por cada edición. La validación existente conserva su alcance sobre Target: un borrado de columnas completas puede seguir siendo costoso.
- BorrarDatos conserva sus ClearContents completos de A:G/U:V; captura el límite de entradas antes de borrar y llama al mismo auxiliar con eventos desactivados. No usa ese límite para recortar el borrado de entradas.
- No es una limpieza inicial del sobrante histórico no afectado. Al borrar hasta 1050, las fórmulas antiguas fuera de ese intervalo se conservan. Con BorrarDatos se limpia solo hasta el último registro anterior; no promete quitar todas las fórmulas históricas posteriores.
- No borra formatos, validaciones, comentarios ni filas físicas. No toca MENÚ/DATOS/ALCON/PLANTILLA. Se omiten tablas, matrices, derrames o combinaciones y se conserva el tramo si una excepción no puede resolverse con seguridad.
- Durante mantenimiento se actualiza el respaldo después de cada rango limpiado para que una recuperación no lo reviva. Ante fallo se intenta cerrar mantenimiento y restaurar protección/eventos/pantalla; no se promete deshacer un borrado manual o de BorrarDatos.
- No hay una nueva medición de tamaño ni tiempo. El formato, los vínculos y el rango usado pueden seguir influyendo en el archivo después de limpiar fórmulas.

## 1. Mapa de responsabilidades del corte 2

| Componente | Rutinas | Responsabilidad |
|---|---:|---|
| [Módulo 4 existente](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-LimpiezaZonaGris/Modulo-4-v7-corte2.txt) | 83 | Entradas públicas, configuración central, tres objetos privados, autorizaciones, importación, MENÚ, plantillas y coordinación. |
| [NeteoValidaciones](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-LimpiezaZonaGris/NeteoValidaciones.txt) | 9 | Validación por bloques, fecha/valor, colores y listas/reglas. Archivo byte idéntico al corte 2 documentado; sin cambios en esta entrega. |
| [NeteoDependencias](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-LimpiezaZonaGris/NeteoDependencias.txt) | 5 | Inspección de referencias y resta rectangular. Archivo byte idéntico al corte 2 documentado; sin cambios en esta entrega. |
| [NeteoRespaldos](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-LimpiezaZonaGris/NeteoRespaldos.txt) | 6 | Crear respaldo temporal, capturar/restaurar contenido, comprobar compatibilidad del rango y copiar el archivo antes de eliminar. Sin estado de autorización. |

**103 rutinas en los cuatro componentes de Módulo 4 (83/9/5/6), más BorrarDatos en Módulo 1.** Conserva las 65 firmas públicas originales y añade un auxiliar público con argumentos para comunicar Módulo 1 con Módulo 4. Cambian solo ActualizarEntradaEnHoja y RegistrarVaciosConservados entre las 82 rutinas anteriores; las otras 80, incluida AgregarFilaMenu, permanecen byte idénticas. No cambia el estado de módulo. Los tres componentes internos son byte idénticos.

### Fronteras internas y propiedad del estado

- NeteoRespaldos lleva `Option Explicit` y `Option Private Module`. Public dentro del proyecto: `CrearRespaldo`, `CapturarContenido`, `RestaurarContenido`, `VerificarContenidoRestaurable`, `GuardarRespaldoPlantilla`. `TextosLiterales` sigue Private y no se duplica.
- `CapturarContenido` y `RestaurarContenido` reciben tres ByVal finales: si la hoja es MENÚ, primera fila de entradas y tamaño de bloque. Módulo 4 evalúa su `EsMenu` privado sobre la misma hoja del rango y transmite sus constantes originales. Dentro se mantiene el límite 10 para MENÚ y el valor original de `PRIMERA_FILA` para operativas (12); no se copian números a cada llamada.
- La propiedad y vida del respaldo siguen en Módulo 4: `hojaMantenimiento`, `respaldoMantenimiento`, `hojaCuentasMenu` permanecen juntos y Private. No se autoriza, conserva ni termina mantenimiento desde el componente de respaldos.
- `CLAVE_HOJA`, `PRIMERA_FILA`, `TAMANO_BLOQUE` y `PREFIJO_CONSERVAR` no se mueven ni duplican. No publicar la clave ni cambiarla aisladamente. Los componentes no usan estado de autorización global nuevo.
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
| Borrar datos, Módulo 1 | `BorrarDatos` | Borra entradas completas y llama la limpieza final compartida con límite anterior; conserva la preparación mínima y los estados previos. |
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
| Instalación | Cinco fuentes completas (Módulos 1/4 y tres internos); nombres internos exactos; 65 macros/APIs originales en Módulo 4 más el auxiliar compartido; proyecto compilado sin duplicados ni referencias inaccesibles. Probar figuras reales: su asignación no fue inventariada. |
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

1. En copia, datos hasta 1005: borrar completamente A:G/U:V de 1005; conservar fila1004, H1005 y las excepciones, quitar solo fórmulas automáticas I:T elegibles de1005.
2. Datos hasta1050: borrar1005:1050; comprobar la limpieza automática del tramo, H1005 mínima, fórmulas distintas/constantes/excepciones intactas y contenido posterior no afectado intacto.
3. Dejar F, U o V con datos en1050 o borrar un campo: esa fila debe conservar H:T y validarse normalmente. Un hueco intermedio no es sobrante final.
4. Volver a escribir entradas en una fila limpiada: recuperar sus H:T elegibles según modelo, sin sobreescribir decisiones conservadas.
5. Conservar durante mantenimiento una fórmula diferente, una igual al modelo y un vacío. Borrar sus entradas; guardar/reabrir y comprobar preservación. Probar restauración inmediatamente anterior y que las fórmulas limpiadas no reaparezcan desde el respaldo.
6. Ejecutar BorrarDatos: borrar todas sus entradas habituales, limpiar solo el tramo anterior utilizado, mantener H12/excepciones y no tocar hojas técnicas. Comprobar estados, protección, cancelaciones y errores en copia.
7. Verificar una fila ya preparada de MENÚ: AgregarFilaMenu no vuelve a pedir reemplazar sus fórmulas. Probar las macros de cuentas/plantillas ya existentes como regresión.
8. Probar un pegado o borrado de varias áreas y uno grande con el volumen habitual; medir antes/después si se busca cuantificar rendimiento. No afirmar resultados económicos ni rendimiento por las pruebas de fuente.

Las pruebas automáticas entregadas internamente son contratos de fuentes y modelos de política, no un intérprete VBA, compilación ni ejecución de eventos/figuras reales. Los fallos históricos de pruebas que exigían identidad total con otra edición documental no se presentan como regresiones nuevas ni como pruebas verdes.

## Reversión de esta entrega

Reponer juntos los contenidos anteriores de Módulo 4 y Módulo 1, conservar los tres módulos internos y compilar. No quitar módulos ni aplicar una reversión antigua de modularización. Restaurar código no recupera fórmulas/datos ya borrados; conservar la copia íntegra previa.

## Catálogo de líneas de la entrega actual

### Modulo-4-v7-corte2.txt

| Procedimiento | Línea |
|---|---:|
| `ObtenerHojaMenu` | 27 |
| `ObtenerHojaPlantilla` | 48 |
| `EsHojaTecnica` | 67 |
| `EsMenu` | 80 |
| `ExisteHoja` | 84 |
| `CopiarAnchosColumnas` | 101 |
| `CopiarAlturasFilas` | 114 |
| `CopiarFormasZonaSuperior` | 127 |
| `ValidarTodo_ReglasNegocio` | 169 |
| `ValidarHojaTrabajo` | 185 |
| `ValidarTodasLasHojasTrabajo` | 207 |
| `EsCodigoOficinaValido` | 216 |
| `EsFechaOrigenPartidaValida` | 225 |
| `EsTipoClienteValido` | 229 |
| `EsNitValido` | 243 |
| `EsNombreClienteValido` | 274 |
| `EsOficinaReclasificacionValida` | 290 |
| `EsDescripcionPartidaValida` | 311 |
| `EsValorPartidaValido` | 327 |
| `EsPartidaReclasificarValida` | 332 |
| `LimpiarTextoValidacion` | 348 |
| `LimpiarColoresZonaEditable` | 375 |
| `ExistenErroresMarcadosEnRojo` | 385 |
| `ConfigurarValidacionesEntradaHojaActiva` | 398 |
| `LimpiarValidacionesEntradaHojaActiva` | 408 |
| `ConfigurarValidacionesEntradaEnHoja` | 430 |
| `RestaurarValidacionesBasicasEnHoja` | 437 |
| `ConfigurarCodigoOficina` | 442 |
| `ConfigurarFechaOrigenPartida` | 448 |
| `ConfigurarTipoCliente` | 454 |
| `ConfigurarOficinaReclasificacion` | 460 |
| `ConfigurarValorPartida` | 466 |
| `ConfigurarPartidaReclasificar` | 472 |
| `ObtenerNitRed` | 478 |
| `ObtenerOficinaRed` | 514 |
| `EsNombreRedCatalogo` | 551 |
| `EsNitRedCatalogo` | 561 |
| `AplicarListaRedesEnCelda` | 598 |
| `QuitarListaRedesEnCelda` | 613 |
| `ActualizarListaRedSegunTipoCliente` | 617 |
| `ImportarDatosDesdeArchivoExcel` | 629 |
| `FilaTieneDatosEntrada` | 797 |
| `ObtenerUltimaFilaEntrada` | 807 |
| `ObtenerUltimaFilaParaRevision` | 819 |
| `ActualizarEntradaEnHoja` | 827 |
| `PrepararFilasEntrada` | 855 |
| `CompletarFormulasZonaGris` | 886 |
| `CompletarColumnaCalculada` | 895 |
| `LimpiarFormulasFinalesVacias` | 948 |
| `ActualizarRespaldoDespuesDeEntrada` | 1092 |
| `ProtegerTodasLasHojasTrabajo` | 1116 |
| `ProtegerHojaTrabajo` | 1129 |
| `ConfigurarBloqueoHoja` | 1134 |
| `RestaurarEstadoHoja` | 1156 |
| `DesbloquearHoja` | 1180 |
| `EstaEnMantenimiento` | 1263 |
| `RestaurarBloqueoFuncional` | 1268 |
| `FinalizarMantenimiento` | 1292 |
| `VolverAProtegerHoja` | 1316 |
| `RestablecerProteccionAlAbrir` | 1335 |
| `ProtegerEstructuraLibro` | 1350 |
| `ZonaCalculada` | 1359 |
| `GestionarCambioMantenimiento` | 1371 |
| `AgregarExcepcion` | 1408 |
| `RegistrarVaciosConservados` | 1425 |
| `VerificarZonaMantenimiento` | 1451 |
| `RecuperarMantenimientoInterrumpido` | 1470 |
| `MenuAutorizado` | 1502 |
| `AutorizarCuentasMenu` | 1510 |
| `AvisarSalidaMenu` | 1541 |
| `AgregarFilaMenu` | 1567 |
| `EliminarUltimaFilaMenu` | 1679 |
| `FilaMenuCompleta` | 1830 |
| `VerificarModeloMenu` | 1856 |
| `VerificarCuentasFilaMenu` | 1865 |
| `VincularFilaMenu` | 1874 |
| `AplicarFormulasEncabezadosDesdeMenu` | 1892 |
| `HojaAsociadaFilaMenu` | 1908 |
| `BordeInferiorFilaMenu` | 1935 |
| `CrearNuevaHojaDesdePlantilla` | 1980 |
| `EliminarUltimaPlantilla` | 2096 |
| `ContarHojasTrabajoCreadas` | 2195 |
| `VerificarSecuenciaPlantillas` | 2199 |

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

