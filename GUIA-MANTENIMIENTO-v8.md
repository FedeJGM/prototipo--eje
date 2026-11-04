# Guía de mantenimiento — v7, corte 2 documentado

## Ruta rápida

**Es la misma implementación del corte 2, con comentarios selectivos de intención.** No es otro corte ni una corrección de negocio. El paquete autocontenido incluye los cuatro TXT compatibles completos, no ThisWorkbook ni fuentes históricas de otros módulos.

1. En una copia, finalizar mantenimiento con el código anterior y pausar Autoguardado. Seguir [LEEME-INSTALACION.txt](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Documentado/LEEME-INSTALACION.txt). Si cambió la contraseña desde la base, detenerse y coordinar sin compartirla.
2. Si ya está instalado el corte 2: reemplazar todo el contenido de los cuatro módulos existentes según CAMBIOS, sin crear ni renombrar módulos. Todos cambian solo por comentarios. No ejecutar, guardar ni cerrar a mitad.
3. Si se instala el paquete completo por primera vez: conservar el Módulo 4 existente y añadir únicamente los tres módulos internos que falten, con nombres exactos; no sobrescribir un módulo cuyo origen no se verificó.
4. Compilar en VBE al terminar y probar en copia. No se compiló/ejecutó VBA ni Excel en esta edición; el control de identidad de código no certifica el libro real.
5. Usar Ctrl+F por nombre; el catálogo final está recalculado para estos TXT. No cambiar asignaciones de figuras ni reinstalar ThisWorkbook/Módulos 1/2/3 instalados, incluidas sus correcciones.

## Qué cambia y qué no

| Archivo de código completo | Estado frente al corte 2 |
|---|---|
| Modulo-4-v7-corte2.txt | Solo comentarios: propiedad del estado, importación, mantenimiento, guardado, fórmulas y eliminación. |
| NeteoValidaciones.txt | Solo comentarios: revisión que cuenta/marca y también ajusta datos según reglas existentes. |
| NeteoDependencias.txt | Solo comentarios: cobertura del recorrido y excepción exacta de fórmulas de MENÚ. |
| NeteoRespaldos.txt | Solo comentarios: snapshot temporal frente a archivo persistente y responsabilidades del controlador. |

No hay archivos de código completos byte idénticos en esta edición: los cuatro reciben comentarios útiles, pero **sus líneas de código son idénticas**. Guía, LEEME y CAMBIOS se actualizan documentalmente. La implementación conserva **82 + 9 + 5 + 6 = 102 rutinas**, firmas/visibilidad/orden, clave, estados y reglas.

Reversión desde corte 2: reponer los cuatro contenidos originales compatibles del corte 2 en los mismos módulos. **No quitar módulos**: no se cambió la arquitectura. Esto revierte comentarios, no recupera hojas/datos borrados.

## 1. Mapa de responsabilidades del corte 2

| Componente | Rutinas | Responsabilidad |
|---|---:|---|
| [Módulo 4 existente](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Documentado/Modulo-4-v7-corte2.txt) | 82 | Entradas públicas, configuración central, tres objetos privados, autorizaciones, importación, MENÚ, plantillas y coordinación. |
| [NeteoValidaciones](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Documentado/NeteoValidaciones.txt) | 9 | Validación por bloques, fecha/valor, colores y listas/reglas. Misma implementación del corte 2; solo comentarios nuevos en esta edición. |
| [NeteoDependencias](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Documentado/NeteoDependencias.txt) | 5 | Inspección de referencias y resta rectangular. Misma implementación del corte 2; solo comentarios nuevos en esta edición. |
| [NeteoRespaldos](C:/Users/feder/Documents/Codex/2026-10-16/referenced-chatgpt-conversation-this-is-an/outputs/Version2/Modular-v7-Corte2-Documentado/NeteoRespaldos.txt) | 6 | Crear respaldo temporal, capturar/restaurar contenido, comprobar compatibilidad del rango y copiar el archivo antes de eliminar. Sin estado de autorización. |

**102 rutinas únicas con la misma arquitectura del corte 2.** Esta edición añade solamente comentarios completos de intención. Las líneas físicas de código, incluidas firmas, declaraciones, espacios, finales CRLF, literales y comentarios en línea, quedan idénticas al retirar únicamente comentarios completos y líneas vacías. Las 65 firmas públicas originales no cambian; no se afirma que los archivos completos tengan el mismo hash.

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

La medición del usuario **48,727 s → 2,391 s** pertenece a la optimización anterior de la revisión previa a confirmar, en un caso. Este corte no agrega una optimización ni garantiza esos tiempos o corrección end-to-end.

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

## 6. Lista de verificación en el libro real

No basta con que compile o aparezca un mensaje correcto. Comparar con una copia anterior equivalente y registrar resultado por caso.

| Flujo | Caso mínimo y resultado a comprobar |
|---|---|
| Instalación | Cuatro archivos completos; nombres internos exactos; 65 macros/APIs originales en Módulo 4; proyecto compilado sin duplicados ni referencias inaccesibles. Probar figuras reales: su asignación no fue inventariada. |
| Entrada/validación | Fecha, NIT, valor distinto de cero, V obligatoria, Cliente/Red y catálogo; valores válidos/erróneos/pegado múltiple. Mismos mensajes y marcas rojas; H especial e I:T recuperadas sin sobrescribir excepciones. |
| Importar | Valores sin reorganizar desde origen válido; límite de zona A:G/U:V; intento de propio libro; confirmación de reemplazo con No; cancelación/fallo y recuperación del rango. No perder datos ni permitir guardado inválido. |
| MENÚ/protección | Bloqueo predeterminado total; clave incorrecta/cancelar sin cambios; cuentas solo A/B y mantenimiento A:L. Agregar/eliminar fila mantienen reglas, fila 10 y asociación. Conservar/restaurar contenido; aviso al salir. |
| Crear plantilla | Modelo PLANTILLA y fila MENÚ válidos; dependencia/gap/modelo inválido bloquea; clave y confirmación según permiso actual; error/cancelación restaura visibilidad, estructura y eventos. |
| Eliminar plantilla | Primera comprobación responder No: nada se borra. Luego, solo en copia desechable con respaldo confirmado: última adicional, nunca 01, dependencias bloquean; A/B y C:J conservadas, K/L provisionales y estado correcto. No prometer recuperación automática tras borrado efectivo. |
| Guardar/reabrir | Guardar desde MENÚ detecta errores operativos y bloquea; corregir y guardar. Reabrir: autorizaciones finalizadas, bloqueos/estructura y fórmulas conservadas; eventos/pantalla normales. |
| Salidas contables | Comparar Reclasificación y ALCON con las salidas esperadas, incluida Hoja1 sin datos. No se modificaron Módulos 1/2/3; comprobar compatibilidad real. |

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

En el corte 2 se comprobaron fuentes y modelos Python de límites, bloques y literales; esta edición solo comprueba identidad de código y documentación: **no ejecución del motor VBA ni de Excel**. Compilar y probar eventos, figuras, recuperación y archivo real sigue pendiente. No pasar al siguiente corte hasta aprobar la regresión del usuario.

## 7. Catálogo completo de esta edición

Líneas verificadas sobre estos cuatro TXT; pueden variar al editar. Buscar por nombre es más estable. No se inspeccionaron asignaciones reales de figuras/OnAction.

### Modulo-4-v7-corte2.txt — 82 rutinas

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
| `FilaTieneDatosEntrada` | 796 |
| `ObtenerUltimaFilaEntrada` | 806 |
| `ObtenerUltimaFilaParaRevision` | 818 |
| `ActualizarEntradaEnHoja` | 826 |
| `PrepararFilasEntrada` | 843 |
| `CompletarFormulasZonaGris` | 874 |
| `CompletarColumnaCalculada` | 883 |
| `ActualizarRespaldoDespuesDeEntrada` | 933 |
| `ProtegerTodasLasHojasTrabajo` | 957 |
| `ProtegerHojaTrabajo` | 970 |
| `ConfigurarBloqueoHoja` | 975 |
| `RestaurarEstadoHoja` | 997 |
| `DesbloquearHoja` | 1021 |
| `EstaEnMantenimiento` | 1104 |
| `RestaurarBloqueoFuncional` | 1109 |
| `FinalizarMantenimiento` | 1133 |
| `VolverAProtegerHoja` | 1157 |
| `RestablecerProteccionAlAbrir` | 1176 |
| `ProtegerEstructuraLibro` | 1191 |
| `ZonaCalculada` | 1200 |
| `GestionarCambioMantenimiento` | 1212 |
| `AgregarExcepcion` | 1249 |
| `RegistrarVaciosConservados` | 1266 |
| `VerificarZonaMantenimiento` | 1299 |
| `RecuperarMantenimientoInterrumpido` | 1318 |
| `MenuAutorizado` | 1350 |
| `AutorizarCuentasMenu` | 1358 |
| `AvisarSalidaMenu` | 1389 |
| `AgregarFilaMenu` | 1415 |
| `EliminarUltimaFilaMenu` | 1512 |
| `FilaMenuCompleta` | 1663 |
| `VerificarModeloMenu` | 1689 |
| `VerificarCuentasFilaMenu` | 1698 |
| `VincularFilaMenu` | 1707 |
| `AplicarFormulasEncabezadosDesdeMenu` | 1725 |
| `HojaAsociadaFilaMenu` | 1741 |
| `BordeInferiorFilaMenu` | 1768 |
| `CrearNuevaHojaDesdePlantilla` | 1813 |
| `EliminarUltimaPlantilla` | 1929 |
| `ContarHojasTrabajoCreadas` | 2028 |
| `VerificarSecuenciaPlantillas` | 2032 |

### NeteoValidaciones.txt — 9 rutinas

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

### NeteoDependencias.txt — 5 rutinas

| Procedimiento | Línea |
|---|---:|
| `VerificarDependenciasPlantilla` | 8 |
| `VerificarTextoDependencia` | 95 |
| `VerificarValidacionesDependientes` | 128 |
| `VerificarFormulasDependientes` | 199 |
| `ExcluirZona` | 271 |

### NeteoRespaldos.txt — 6 rutinas

| Procedimiento | Línea |
|---|---:|
| `CrearRespaldo` | 11 |
| `TextosLiterales` | 32 |
| `CapturarContenido` | 49 |
| `RestaurarContenido` | 94 |
| `VerificarContenidoRestaurable` | 140 |
| `GuardarRespaldoPlantilla` | 165 |

