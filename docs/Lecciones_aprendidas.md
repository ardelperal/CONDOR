Lecciones Aprendidas - Proyecto CONDOR
Este documento centraliza las lecciones de arquitectura y flujo de trabajo aprendidas durante el desarrollo del proyecto CONDOR. Su propósito es servir como guía para mantener la calidad, consistencia y mantenibilidad del código.

Lección 1: La Estricta Naturaleza de las Interfaces en VBA
Observación: A diferencia del resto de VBA, la implementación de interfaces (Implements) es estrictamente sensible a los detalles de la firma del procedimiento.
Regla Inquebrantable: La firma de un método en una clase implementadora debe ser una copia idéntica, carácter por carácter, de la firma en la interfaz.
Acción Correctiva: Ante errores de "declaración no coincide", se debe usar un prompt de sincronización forzada, tratando la interfaz como la única fuente de verdad y reescribiendo las firmas en las clases implementadoras.

Lección 2: El Principio de "Programar Contra la Interfaz" en los Tests
Observación: Los errores de "método no encontrado" en los tests ocurren cuando se declara una variable del tipo de la clase concreta en lugar de la interfaz.
Regla Inquebrantable: Dentro de cualquier módulo de pruebas (Test_*.bas), las variables que referencian a nuestros servicios deben ser declaradas del tipo de su interfaz.
Acción Correctiva: Usar periódicamente prompts de auditoría de calidad de pruebas para verificar que todos los tests cumplen con esta y otras reglas de estructura (AAA, manejo de errores, etc.).

Lección 3: Estructura de Módulos y Clases en VBA
Observación: El compilador de VBA es estricto con el orden de las declaraciones dentro de un fichero.
Regla Inquebrantable: Todas las declaraciones a nivel de módulo (Public/Private/Dim para variables, Type, Enum, Declare) deben estar agrupadas en la sección de declaraciones, en la parte superior del fichero, antes de la primera definición de Sub, Function o Property.
Acción Correctiva: Ante errores de "comentario solo puede aparecer después de End Sub..." o similares, la causa raíz suele ser una declaración fuera de lugar. Se debe mover a la parte superior del fichero.

Lección 4: El Flujo de Trabajo rebuild -> Compilación Manual -> test
Observación: El comando rebuild del CLI sincroniza los ficheros, pero no garantiza la compilación en tiempo de ejecución. Muchos errores solo se manifiestan al compilar dentro de Access.
Regla Inquebrantable: El flujo de trabajo estándar es el Ciclo de Trabajo Asistido. Ninguna prueba se ejecuta hasta que el Supervisor haya confirmado que el proyecto compila exitosamente de forma manual (Depuración -> Compilar Proyecto).
Acción Correctiva: El prompt para Copilot siempre debe finalizar con una pausa para la verificación manual del Supervisor antes de proceder con los tests o el commit.

Lección 5: Conversión Explícita de Tipos desde Arrays Variant
Observación: Al iterar sobre un array de tipo Variant (creado con Array(...)) y pasar sus elementos a una función que espera un tipo de dato específico (Long, String, etc.), VBA puede fallar al realizar la conversión de tipo implícita.
Regla Inquebrantable: Para garantizar la robustez, siempre se debe realizar una conversión de tipo explícita al pasar un elemento de un array Variant a una función que espera un tipo específico.
Acción Correctiva: Ante este error, se debe añadir la función de conversión apropiada (CLng, CStr, CInt, CBool, etc.) en la llamada al procedimiento.

Lección 6: Usar los Tests y Módulos de Acceso a Datos como Especificación para Clases de Datos
Observación: Errores de "método o dato miembro no encontrado" ocurren frecuentemente en los tests y módulos de acceso a datos al usar clases de tipo de datos (T_*.cls) que están incompletas.
Regla Inquebrantable: Los tests que construyen objetos de datos y los módulos de acceso a datos que asignan valores a propiedades de objetos actúan como la especificación funcional para esas clases de datos. La clase debe contener todas las propiedades públicas que tanto los tests como los módulos de datos utilizan.
Acción Correctiva: Ante este error, se debe realizar una auditoría proactiva completa de todos los módulos que utilizan la clase de datos para añadir todas las propiedades faltantes.

Lección 7: La Batería de Pruebas Debe Ser Exhaustiva
Observación: Se han detectado nuevos módulos de prueba que no se han añadido a la función principal de ejecución de pruebas en modTestRunner.bas.
Regla Inquebrantable: Cada vez que se cree un nuevo módulo de pruebas (Test_...bas), su función de ejecución principal (Test_*_RunAll) debe ser registrada inmediatamente dentro de la función RegisterTestSuites en modTestRunner.bas.
Acción Correctiva: Todos los prompts que impliquen la creación de un nuevo módulo de pruebas deben incluir explícitamente un paso final para modificar modTestRunner.bas y añadir la llamada a la nueva suite.

Lección 8: La Centralización del Manejo de Errores es Obligatoria
Observación: Se ha detectado código que, si bien utiliza On Error GoTo ErrorHandler, no registra el error capturado en nuestro servicio central modErrorHandler.
Regla Inquebrantable: Un bloque ErrorHandler sin una llamada a modErrorHandler.LogError (o LogCriticalError) se considera una implementación incompleta y errónea.
Acción Correctiva: Todos los prompts que impliquen la creación o modificación de código deben incluir la directiva de implementar el manejo de errores centralizado.

Lección 9: La Auditoría de Operaciones es un Requisito, no una Opción
Observación: La trazabilidad de las acciones de negocio (quién hizo qué y cuándo) es tan crítica como el registro de errores para la seguridad y el mantenimiento del sistema.
Regla Inquebrantable: Toda función o procedimiento que represente una acción de negocio significativa debe registrar dicha acción a través del servicio IOperationLogger.
Acción Correctiva: Todos los prompts para la creación de nuevas funcionalidades deben incluir explícitamente el requisito de implementar el logging de auditoría en los puntos clave del negocio.

Lección 10: El Aislamiento de las Pruebas Unitarias con Mocks no es Negociable
Observación: Se han detectado pruebas unitarias que interactúan con dependencias reales (conexiones a bases de datos, acceso al sistema de ficheros), lo que las hace lentas y frágiles.
Regla Inquebrantable: Una prueba unitaria debe probar una única "unidad" de código de forma aislada. Todas las dependencias externas deben ser reemplazadas por Mocks.
Acción Correctiva: Todos los prompts para la creación de pruebas unitarias deben incluir explícitamente el requisito de usar Mocks para todas las dependencias externas.

Lección 11: La Plantilla Estándar para Clases de Servicio
Observación: Para asegurar la consistencia y la aplicación de todos los principios aprendidos, toda nueva clase de servicio debe seguir una estructura estándar.
Regla Inquebrantable: Toda nueva clase de servicio debe tener la siguiente estructura mínima: Implements INombreDelServicio, dependencias privadas, un método Initialize, implementación de la interfaz, manejo de errores centralizado y logging de auditoría.
Acción Correctiva: Utilizar la plantilla estándar para la creación de nuevas clases de servicio.

Lección 12: La Separación del Frontend y el Backend es Crítica
Observación: El uso de CurrentDb en módulos que acceden a datos rompe la arquitectura cliente-servidor. El código del frontend no debe interactuar directamente con la base de datos de datos.
Regla Inquebrantable: La base de datos de datos (_datos.accdb) debe ser accedida exclusivamente a través de conexiones DAO, utilizando la ruta y la contraseña almacenadas en el servicio de configuración (modConfig). El uso de CurrentDb solo es válido para operaciones en la base de datos del frontend (código VBA, formularios, etc.).
Acción Correctiva: Reemplazar todas las instancias de CurrentDb en los repositorios por DBEngine.OpenDatabase(modConfig.GetInstance().GetDataPath()).

Lección 13: La Centralización de la Configuración y la Seguridad
Observación: Las configuraciones sensibles, como las contraseñas, no deben estar hardcodeadas directamente en el código de la aplicación.
Regla Inquebrantable: Todos los valores de configuración, especialmente los sensibles, deben ser gestionados por una única fuente de verdad: el servicio de configuración (CConfig). El acceso a estos valores debe realizarse a través de los métodos de la clase de configuración, como GetDatabasePassword().
Acción Correctiva: Eliminar todas las cadenas de contraseña hardcodeadas del código de la aplicación y reemplazarlas por llamadas al servicio de configuración.

Lección 14: La Eliminación Proactiva de Código Muerto es un Requisito de Mantenibilidad
Observación: Con la refactorización, es común que funciones, variables o clases enteras queden sin uso, aumentando la complejidad del código.
Regla Inquebrantable: El código que no se utiliza, se elimina. No se debe mantener código comentado o sin referencias "por si acaso".
Acción Correctiva: Cualquier prompt que implique una refactorización o la eliminación de una funcionalidad debe incluir explícitamente un paso final: "Busca y elimina cualquier función, variable o clase de ayuda que haya quedado sin referencias tras este cambio".

Lección 15: No Documentar Cambios en el Desarrollo
Observación: En esta fase del proyecto, documentar cada cambio en el desarrollo crea una sobrecarga innecesaria y ralentiza el progreso. Regla Inquebrantable: No se debe documentar el estado de los cambios durante el desarrollo. La única documentación requerida en esta etapa es el estado final del proyecto una vez que se considera un producto mínimo viable. Acción Correctiva: Los prompts para Copilot no deben incluir instrucciones para actualizar la documentación de desarrollo, como PLAN_DE_ACCION.md o README.md, o CONDOR_MASTER_PLAN.md hasta que el producto mínimo viable esté completo. La documentación se actualizará de forma integral al finalizar la fase actual.