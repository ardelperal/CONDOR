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
Regla Inquebrantable: Los tests que construyen objetos de datos y los módulos de acceso a datos que asignan valores a propiedades de objetos actúan como la especificación funcional para esas clases de datos. La clase debe contener todas las propiedades públicas que tanto los tests como los módulos de datos utilizan. Por cada tabla tb... en la base de datos, debe existir su correspondiente clase de datos T_... en el código. Esta correspondencia 1:1 garantiza que los repositorios puedan mapear correctamente los datos de un Recordset a un objeto.
Acción Correctiva: Ante este error, se debe realizar una auditoría proactiva completa de todos los módulos que utilizan la clase de datos para añadir todas las propiedades faltantes.

Lección 7: La Batería de Pruebas Debe Ser Exhaustiva
Observación: Se han detectado nuevos módulos de prueba que no se han añadido a la función principal de ejecución de pruebas en modTestRunner.bas.
Regla Inquebrantable: Cada vez que se cree un nuevo módulo de pruebas (Test_...bas), su función de ejecución principal (Test_*_RunAll) debe ser registrada inmediatamente dentro de la función RegisterTestSuites en modTestRunner.bas.
Acción Correctiva: Todos los prompts que impliquen la creación de un nuevo módulo de pruebas deben incluir explícitamente un paso final para modificar modTestRunner.bas y añadir la llamada a la nueva suite.

Lección 8: La Centralización del Manejo de Errores es Obligatoria
Observación: Se ha detectado código que, si bien utiliza On Error GoTo ErrorHandler, no registra el error capturado en nuestro servicio central modErrorHandler.
Regla Inquebrantable: Un bloque ErrorHandler sin una llamada a modErrorHandler.LogError (o LogCriticalError) se considera una implementación incompleta y errónea.
Acción Correctiva: Todos los prompts que impliquen la creación o modificación de código deben incluir la directiva de implementar el manejo de errores centralizado.

Regla de Estilo Crítica: La variable del servicio de errores (Dim ErrorHandler As IErrorHandlerService) debe declararse una única vez al inicio del procedimiento, nunca repetirse dentro del bloque ErrorHandler:. Las declaraciones duplicadas causan errores de compilación que impiden el registro del error original, dejando el sistema en un estado inconsistente e inaceptable.

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

Ejemplo de Violación (CAuthRepository antes de refactorización):
```vba
Private m_Database As DAO.Database

Private Sub Class_Initialize()
    Set m_Database = CurrentDb  ' ❌ INCORRECTO: Viola arquitectura cliente-servidor
End Sub
```

Ejemplo Correcto (CSolicitudRepository):
```vba
Private m_configService As IConfig
Private m_operationLogger As IOperationLogger

Public Sub Initialize(ByVal configService As IConfig, ByVal operationLogger As IOperationLogger)
    Set m_configService = configService
    Set m_operationLogger = operationLogger
End Sub

Private Function GetData() As DAO.Recordset
    ' ✅ CORRECTO: Conexión explícita al Backend
    Dim backendPath As String
    backendPath = m_configService.GetDataPath()
    
    Dim backendPassword As String
    backendPassword = m_configService.GetDatabasePassword()
    
    Set db = DBEngine.OpenDatabase(backendPath, False, False, ";PWD=" & backendPassword)
End Function
```

Acción Correctiva: Reemplazar todas las instancias de CurrentDb en los repositorios por DBEngine.OpenDatabase con configuración inyectada.

Lección 13: La Centralización de la Configuración y la Seguridad
Observación: Las configuraciones sensibles, como las contraseñas, no deben estar hardcodeadas directamente en el código de la aplicación. Además, el uso de strings genéricos con GetValue() es propenso a errores tipográficos y dificulta el mantenimiento.
Regla Inquebrantable: Todos los valores de configuración, especialmente los sensibles, deben ser gestionados por una única fuente de verdad: el servicio de configuración (CConfig). Para configuraciones críticas y frecuentemente utilizadas, se deben crear métodos específicos en la interfaz IConfig (como GetDataPath(), GetDatabasePassword()) en lugar de usar strings genéricos con GetValue(). Esto proporciona seguridad de tipos, detección temprana de errores y facilita la refactorización.

Sistema de Configuración de Dos Niveles (Frontend/Backend): La aplicación implementa una arquitectura de configuración de dos niveles. TbLocalConfig (ubicada en el Frontend) actúa como tabla de arranque (bootstrap) que contiene únicamente el indicador de entorno ('LOCAL' o 'OFICINA'). tbConfiguracion (ubicada en el Backend) contiene todos los parámetros globales de la aplicación. El sistema lee el entorno desde TbLocalConfig y utiliza constantes de ruta base definidas en modConfig.bas para construir dinámicamente la ruta del backend, accediendo luego a tbConfiguracion para obtener la configuración completa. Esta separación hace la aplicación completamente portable entre entornos sin necesidad de modificar datos o código, cumpliendo el principio de "configuración sobre convención" y eliminando errores de despliegue.

Acción Correctiva: Eliminar todas las cadenas de contraseña hardcodeadas del código de la aplicación y reemplazarlas por llamadas al servicio de configuración. Para configuraciones críticas, crear métodos específicos en IConfig y refactorizar el código existente para usar estos métodos en lugar de GetValue() con strings. Implementar validación robusta de entornos con mensajes de error descriptivos para configuraciones incorrectas.

Lección 14: La Eliminación Proactiva de Código Muerto es un Requisito de Mantenibilidad
Observación: Con la refactorización, es común que funciones, variables o clases enteras queden sin uso, aumentando la complejidad del código.
Regla Inquebrantable: El código que no se utiliza, se elimina. No se debe mantener código comentado o sin referencias "por si acaso".
Acción Correctiva: Cualquier prompt que implique una refactorización o la eliminación de una funcionalidad debe incluir explícitamente un paso final: "Busca y elimina cualquier función, variable o clase de ayuda que haya quedado sin referencias tras este cambio".

Nota Adicional: Este principio se aplica especialmente a las propiedades de las clases de datos (T_*) cuando se eliminan columnas de las tablas correspondientes. Al eliminar un campo de una tabla, la variable miembro privada y las propiedades públicas (Get/Let) asociadas en la clase T_* correspondiente se convierten automáticamente en código muerto y deben ser eliminadas para mantener la consistencia entre el esquema de datos y las clases de datos.

Lección 15: No Documentar Cambios en el Desarrollo
Observación: En esta fase del proyecto, documentar cada cambio en el desarrollo crea una sobrecarga innecesaria y ralentiza el progreso. Regla Inquebrantable: No se debe documentar el estado de los cambios durante el desarrollo. La única documentación requerida en esta etapa es el estado final del proyecto una vez que se considera un producto mínimo viable. Acción Correctiva: Los prompts para Copilot no deben incluir instrucciones para actualizar la documentación de desarrollo, como PLAN_DE_ACCION.md o README.md, o CONDOR_MASTER_PLAN.md hasta que el producto mínimo viable esté completo. La documentación se actualizará de forma integral al finalizar la fase actual.

Lección 16: El Patrón Infalible para la Limpieza de Recursos en Error Handlers
Observación: Los bloques ErrorHandler que contienen errores de compilación son inaceptables, ya que impiden el registro del error original y dejan el sistema en un estado inconsistente. Un patrón común de error es verificar el estado de objetos de base de datos (como rs.State) antes de cerrarlos, lo cual puede fallar si el error ocurrió antes de que el objeto fuera completamente inicializado.
Regla Inquebrantable: La única forma aceptada para cerrar objetos de base de datos (Recordsets, QueryDefs, Database) dentro de un ErrorHandler es verificando primero que no sean Nothing. El patrón estándar es: `If Not [objeto] Is Nothing Then [objeto].Close`. Nunca se debe verificar el estado del objeto (.State) ya que esto puede generar errores adicionales.
Acción Correctiva: Todos los bloques ErrorHandler deben usar exclusivamente el patrón `If Not rs Is Nothing Then rs.Close` para la limpieza de Recordsets, y patrones similares para otros objetos de base de datos. Cualquier verificación de .State debe ser eliminada.

Lección 17: Principio de Responsabilidad Única para Repositorios
Observación: Se ha detectado que servicios como CExpedienteService dependían incorrectamente de ISolicitudRepository para obtener datos de expedientes, violando el principio de responsabilidad única y creando un acoplamiento inadecuado entre entidades de negocio diferentes.
Regla Inquebrantable: Cada repositorio debe gestionar una única entidad de negocio y sus datos relacionados. ISolicitudRepository solo debe manejar operaciones sobre la entidad Solicitud (T_Solicitud), IExpedienteRepository solo debe manejar operaciones sobre la entidad Expediente (T_Expediente), etc. Los servicios deben depender de los repositorios apropiados para cumplir sus contratos de interfaz.
Acción Correctiva: Cuando un servicio no puede cumplir su contrato de interfaz con las dependencias actuales, se debe crear el repositorio específico para la entidad que necesita y refactorizar el servicio para usar la dependencia correcta. Nunca se debe "reutilizar" un repositorio de una entidad diferente para acceder a datos de otra entidad.

Caso Crítico - CExpedienteService (2024): Se detectó un fallo de diseño arquitectónico crítico donde CExpedienteService estaba dependiendo de ISolicitudRepository en lugar de IExpedienteRepository. Este error causaba fallos de compilación al intentar llamar métodos como ObtenerExpedientePorNemotecnico que no existen en ISolicitudRepository. La corrección requirió:
- Verificar que CExpedienteService dependa exclusivamente de IExpedienteRepository
- Confirmar que modExpedienteServiceFactory inyecte la dependencia correcta
- Validar que los tests usen CMockExpedienteRepository

Ejemplo Práctico - Refactorización del AuthService:
Antes: CAuthService dependía incorrectamente de ISolicitudRepository para consultar datos de usuarios y permisos, violando el principio de responsabilidad única.
Después: Se creó IAuthRepository con CAuthRepository para gestionar exclusivamente las consultas de autenticación (TbUsuariosAplicaciones, TbUsuariosAplicacionesPermisos). CAuthService ahora depende de IAuthRepository, separando claramente las responsabilidades:
- ISolicitudRepository: Gestiona T_Solicitud y datos relacionados
- IAuthRepository: Gestiona consultas de usuarios y permisos de aplicación
- IExpedienteRepository: Gestiona T_Expediente y datos relacionados
Esta separación mejora la mantenibilidad, testabilidad y cumple el principio de responsabilidad única. NUNCA un servicio debe depender de un repositorio de otra entidad de negocio.

Lección 18: La Prevención de Inyección de SQL con Consultas Parametrizadas es Obligatoria
Observación: Se han detectado repositorios que construyen consultas SQL mediante la concatenación de strings, lo cual introduce una vulnerabilidad de seguridad crítica de Inyección de SQL.
Regla Inquebrantable: Nunca se debe construir una consulta SQL concatenando directamente valores de entrada del usuario o del programa. Es una práctica de seguridad inaceptable que puede comprometer toda la base de datos.
Acción Correctiva: Todo acceso a la base de datos que implique una cláusula WHERE debe realizarse exclusivamente a través de consultas parametrizadas, utilizando un objeto `DAO.QueryDef` y su colección `Parameters`. Este patrón es el estándar del proyecto y su cumplimiento no es negociable.

Lección 19: La Comunicación entre el CLI y VBA Debe Basarse en Valores de Retorno, no en Ficheros
Observación: Los sistemas basados en ficheros de log para la comunicación entre el CLI (condor_cli.vbs) y el motor de pruebas VBA son frágiles y propensos a errores. Esta arquitectura introduce "efectos secundarios" que violan los principios de código limpio y bajo acoplamiento.
Regla Inquebrantable: Las funciones VBA llamadas desde el CLI deben devolver resultados directamente a través de valores de retorno. La comunicación debe ser síncrona y directa, eliminando la dependencia de ficheros intermedios que pueden fallar, corromperse o no generarse.
Acción Correctiva: Refactorizar toda comunicación CLI-VBA para usar `objAccess.Application.Run("FuncionVBA")` capturando el valor de retorno directamente. Las funciones VBA deben devolver strings estructurados que incluyan tanto el reporte legible como indicadores parseables (ej: "RESULT: SUCCESS" o "RESULT: FAILED") para facilitar la automatización.

Lección 20: La Automatización sobre la Configuración Manual (Principio de Cero Mantenimiento)
Observación: El registro manual de suites de pruebas en la función RegisterAllSuites de modTestRunner.bas es una deuda técnica crítica. Cada vez que se añade un nuevo fichero de pruebas (Test_*.bas o IntegrationTest_*.bas), existe el riesgo de olvidar registrarlo manualmente, lo que resulta en pruebas que no se ejecutan y una falsa sensación de seguridad en la calidad del código.
Regla Inquebrantable: Los sistemas deben auto-configurarse basándose en convenciones de nomenclatura, eliminando completamente la intervención manual. El descubrimiento automático de suites de pruebas debe basarse en la inspección dinámica del proyecto VBA, identificando módulos que cumplan con las convenciones establecidas (nombres que comiencen con "Test_" o "IntegrationTest_").

Lección 25: Un Framework de Pruebas Robusto Requiere una Librería de Aserciones Completa
Observación: Los errores de compilación como "AssertNotNull no existe" revelan que nuestra librería de aserciones (modAssert.bas) está incompleta. Una librería de aserciones incompleta limita la expresividad y legibilidad de nuestras pruebas, forzando a los desarrolladores a usar workarounds o aserciones menos específicas.
Regla Inquebrantable: La librería de aserciones debe proporcionar funciones específicas para todos los casos de validación comunes en las pruebas. Esto incluye verificaciones de nulidad (AssertNotNull, AssertIsNull), igualdad, desigualdad, rangos, tipos de datos, y estados de objetos. Cada función de aserción debe proporcionar mensajes de error claros y específicos.
Acción Correctiva: Ante errores de "método no encontrado" en aserciones, se debe implementar inmediatamente la función faltante en modAssert.bas siguiendo el patrón estándar: validación de la condición, lanzamiento de error con código específico (vbObjectError + número único) y mensaje descriptivo. La completitud de la librería de aserciones es un prerequisito para pruebas de calidad.

Lección 27: Pruebas Automatizadas Desatendidas
Observación: Las pruebas automatizadas fallaban en entornos desatendidos debido a diálogos de confirmación y alertas de Microsoft Access que requerían interacción del usuario.
Regla Inquebrantable: Para pruebas automatizadas en Access, siempre configurar DisplayAlerts = False antes de ejecutar operaciones, usar dbFailOnError en lugar de False para conexiones de base de datos, diseñar funciones principales con parámetros opcionales para testing, y evitar cualquier dependencia de interacción del usuario en el flujo de pruebas.
Acción Correctiva: Configurar Access en modo silencioso (objAccess.Application.DisplayAlerts = False, objAccess.Application.Echo False, objAccess.AutomationSecurity = 1), reemplazar False por dbFailOnError en todas las llamadas a DBEngine.OpenDatabase para evitar diálogos de confirmación, y refactorizar funciones críticas como App_Start para aceptar parámetros opcionales que permitan inyección de datos de prueba sin requerir autenticación interactiva.

Ejemplo de Implementación Correcta:
```vba
' En condor_cli.vbs - Configuración para modo desatendido
objAccess.Application.DisplayAlerts = False
objAccess.Application.Echo False
objAccess.AutomationSecurity = 1

' En repositorios - Uso de dbFailOnError
Set db = DBEngine.OpenDatabase(backendPath, dbFailOnError, False, ";PWD=" & backendPassword)

' En modAppManager.bas - Parámetro opcional para testing
Public Function App_Start(Optional ByVal testUserEmail As String = "") As Boolean
    If testUserEmail <> "" Then
        ' Usar email de prueba sin autenticación
        g_CurrentUserEmail = testUserEmail
    Else
        ' Flujo normal de autenticación
        g_CurrentUserEmail = authService.GetCurrentUserEmail()
    End If
End Function
```

Impacto: Permite la ejecución completamente desatendida de pruebas, integración exitosa con pipelines de CI/CD, y eliminación de falsos positivos por diálogos de confirmación.
Acción Correctiva: Implementar un sistema de descubrimiento automático que utilice la librería "Microsoft Visual Basic for Applications Extensibility 5.3" para inspeccionar dinámicamente todos los vbComponents del proyecto, identificar módulos de prueba por convención de nomenclatura, y registrarlos automáticamente sin intervención manual. Esto garantiza que todas las pruebas se ejecuten siempre, independientemente de errores humanos en el mantenimiento.

Lección 21: Las Clases del Framework de Pruebas Deben Ser Robustas
Observación: Los errores de "método no encontrado" en el motor de pruebas ocurren cuando las clases de resultados (CTestResult, CTestSuiteResult) no tienen los métodos necesarios para ser instanciadas y configuradas de forma fiable, especialmente dentro de los bloques de manejo de errores del propio TestRunner.
Regla Inquebrantable: Las clases de resultados del framework de pruebas deben tener métodos Initialize para poder ser configuradas de forma consistente y robusta. Esto incluye CTestResult.Initialize(testName), CTestSuiteResult.Initialize(suiteName), y métodos auxiliares como Fail() y Pass() para CTestResult. Estos métodos son críticos para el manejo de errores en el TestRunner.
Acción Correctiva: Implementar métodos Initialize en todas las clases de resultados del framework de pruebas, asegurando que puedan ser instanciadas y configuradas de forma fiable incluso en situaciones de error. Esto garantiza que el sistema de pruebas pueda reportar fallos de ejecución de suites completas sin generar errores adicionales de compilación.

Lección 22: Los Mocks Deben Ser Clases Completas y Reutilizables
Observación: Los errores de "tipo no definido" en las pruebas ocurren cuando se intenta usar mocks que no están completamente implementados. Un mock incompleto rompe el principio de aislamiento de las pruebas unitarias y hace que las pruebas dependan de implementaciones reales, violando la Lección 10.
Regla Inquebrantable: Cada mock (CMock...) debe ser una clase completa que implemente su interfaz correspondiente, permita la configuración de sus valores de retorno (mediante métodos como .AddSetting) y registre las llamadas (..._WasCalled, ..._CallCount) para facilitar las aserciones en las pruebas. Los mocks deben ser reutilizables entre diferentes pruebas y suites.
NOTA CRÍTICA: Nuestros mocks inteligentes (como CMockConfig) se configuran usando métodos de preparación (ej. .AddSetting), no asignando valores a propiedades ..._ReturnValue. El patrón correcto es mockConfig.AddSetting "CLAVE", valor en lugar de mockConfig.Propiedad_ReturnValue = valor.
Acción Correctiva: Implementar clases mock completas para todas las interfaces del sistema, incluyendo variables privadas para almacenar configuraciones, métodos públicos para configurar valores de retorno, propiedades de seguimiento de llamadas, métodos Reset para limpiar el estado entre pruebas, y la implementación completa de todos los métodos de la interfaz correspondiente.

Lección 23: Las Auditorías de Calidad de Pruebas Verifican el Aislamiento
Observación: Con el tiempo, las pruebas unitarias pueden degradarse y comenzar a depender de clases reales en lugar de mocks, convirtiéndose en "falsos unitarios" que pueden ocultar errores o fallar por razones equivocadas. Esta degradación compromete la fiabilidad de la red de seguridad de pruebas.
Regla Inquebrantable: Se deben realizar auditorías periódicas de calidad para verificar que todas las pruebas unitarias (módulos Test_*.bas, excluyendo IntegrationTest_*) usen exclusivamente mocks para sus dependencias externas. Cualquier instanciación de clases concretas (New C...) en lugar de mocks (New CMock...) debe ser corregida inmediatamente.
Acción Correctiva: Implementar un proceso de auditoría sistemática que revise todos los módulos de prueba unitaria, identifique instanciaciones incorrectas de clases concretas, las reemplace por sus mocks correspondientes, y verifique que las variables se declaren con el tipo de interfaz apropiado. Esta auditoría debe ejecutarse antes de cada release y después de refactorizaciones significativas.

Lección 24: Las Clases Concretas Deben Exponer Métodos Públicos de Conveniencia
Observación: Los errores de "método no encontrado" en pruebas de integración ocurren cuando las clases concretas solo implementan métodos privados de interfaz, impidiendo el acceso directo desde las pruebas o el uso de la clase sin declarar variables del tipo de interfaz.
Regla Inquebrantable: Toda clase que implemente una interfaz debe exponer también métodos Public con la misma firma que deleguen la llamada a la implementación Private de la interfaz. Esto facilita las pruebas de integración y el uso directo de la clase sin romper el encapsulamiento ni el principio de "Programar Contra la Interfaz".
Acción Correctiva: Para cada método Private Function IInterfaz_Metodo implementado en una clase, crear un método Public Function Metodo correspondiente que simplemente delegue la llamada: Public Function Metodo(...) As TipoRetorno / Set Metodo = IInterfaz_Metodo(...) / End Function. Aplicar este patrón consistentemente en todas las clases de repositorio y servicios del proyecto.

Lección 25: Un Framework de Pruebas Robusto Requiere una Librería de Aserciones Completa
Observación: Los errores de "método no encontrado" en el framework de pruebas (como AssertNotNull no existe) demuestran que nuestra librería modAssert.bas está incompleta. No podemos escribir buenas pruebas si nos faltan las herramientas básicas de verificación.
Regla Inquebrantable: El módulo modAssert.bas debe proporcionar un conjunto completo y simétrico de funciones de aserción que cubran todos los casos de uso comunes en las pruebas: verificación de valores booleanos, nulos/no nulos, igualdad, desigualdad, y rangos. Cada aserción debe lanzar errores específicos con códigos únicos para facilitar la depuración.
Acción Correctiva: Implementar sistemáticamente todas las funciones de aserción necesarias en modAssert.bas, incluyendo pares simétricos (AssertNotNull/AssertIsNull, AssertEquals/AssertNotEquals, etc.), con documentación completa de parámetros y códigos de error. Mantener una lista de verificación de aserciones implementadas para garantizar la completitud del framework.

Funciones de Aserción Estándar Implementadas:
- AssertTrue(condition As Boolean, message As String) - Error: vbObjectError + 510
- AssertFalse(condition As Boolean, message As String) - Error: vbObjectError + 511
- AssertEquals(expected As Variant, actual As Variant, message As String) - Error: vbObjectError + 512
- AssertNotNull(obj As Object, message As String) - Error: vbObjectError + 513
- AssertIsNull(obj As Object, message As String) - Error: vbObjectError + 514
- Fail(message As String) - Error: vbObjectError + 515
- IsTrue(condition As Boolean) - Función de compatibilidad que no lanza errores

Meta-Testing: El módulo Test_modAssert.bas contiene pruebas unitarias para cada función de aserción, verificando tanto casos de éxito como de fallo, garantizando que el framework se pruebe a sí mismo.

Lección 26: Encapsulación de Lógica Común en Clases que Implementan Interfaces
Observación: Los errores de compilación "IConfig_GetValue no se puede llamar con Me" ocurren cuando los métodos de implementación de interfaz intentan llamarse entre ellos usando Me.IInterfaz_Metodo. Esto viola la naturaleza estricta de las interfaces en VBA, donde los métodos Private Function IInterfaz_Metodo no pueden ser invocados directamente.
Regla Inquebrantable: Los métodos de implementación de interfaz (Private Function IInterfaz_Metodo) no deben llamarse entre ellos. La lógica común debe ser encapsulada en funciones auxiliares Private que puedan ser invocadas por todos los métodos de la interfaz.
Acción Correctiva: Crear funciones auxiliares Private para la lógica compartida (ej: GetSettingValue, GetMockSettingValue) y hacer que todos los métodos de implementación de interfaz las utilicen en lugar de intentar llamarse entre ellos. Esto elimina las dependencias circulares y respeta la arquitectura estricta de interfaces de VBA.

Lección 27: Las Pruebas Automatizadas Deben Ser 100% Desatendidas
Observación: Los pop-ups de base de datos y diálogos interactivos durante la ejecución de pruebas desde el CLI rompen la automatización y hacen que el sistema de CI/CD falle. Cualquier interacción con la interfaz de usuario durante las pruebas automatizadas es una violación crítica del principio de ejecución desatendida.
Regla Inquebrantable: Todo código ejecutado en modo de prueba no debe, bajo ninguna circunstancia, generar una interacción con la interfaz de usuario. Las conexiones a datos deben ser explícitas y usar opciones como dbFailOnError para forzar un error programático en lugar de mostrar diálogos. El CLI debe configurar DisplayAlerts = False para suprimir diálogos inesperados.
Acción Correctiva: Auditar todas las llamadas a DBEngine.OpenDatabase para incluir el parámetro dbFailOnError, configurar objAccess.Application.DisplayAlerts = False en el CLI antes de ejecutar pruebas, y establecer conexiones de base de datos con parámetros explícitos que fuercen errores programáticos en lugar de diálogos interactivos. Cualquier código que pueda generar pop-ups debe ser refactorizado para manejar errores de forma programática.

Lección 28: La Construcción Dinámica de Rutas Elimina el Acoplamiento Fuerte con el Sistema de Archivos
Observación: El almacenamiento de rutas absolutas en la base de datos (como rutaArchivo en tbAdjuntos) crea un acoplamiento fuerte con la estructura del sistema de archivos, haciendo que la aplicación sea frágil ante cambios de ubicación, movimientos de servidores o despliegues en diferentes entornos.
Regla Inquebrantable: Las rutas de archivos nunca deben almacenarse como rutas absolutas en la base de datos. En su lugar, se debe almacenar únicamente el nombre del archivo y construir la ruta completa dinámicamente en tiempo de ejecución combinando una ruta base obtenida de la configuración (IConfig.GetAttachmentsPath()) con el nombre del archivo.
Acción Correctiva: Eliminar campos de ruta absoluta (rutaArchivo) de las tablas de datos y reemplazarlos con campos de nombre de archivo (nombreArchivo). Implementar métodos en IConfig para obtener rutas base configurables (GetAttachmentsPath, GetTemplatesPath, etc.) y construir rutas completas dinámicamente usando Path.Combine o concatenación de strings. Esto hace la aplicación completamente portable entre entornos sin necesidad de actualizar datos en la base de datos.
NOTA ADICIONAL: De forma similar, otros metadatos de los ficheros (como el tipo de archivo) deben derivarse dinámicamente del nombre del fichero usando la extensión, en lugar de almacenarse de forma redundante en campos separados. Esto elimina la posibilidad de inconsistencias entre el nombre del archivo y sus metadatos almacenados, y reduce la complejidad del esquema de datos siguiendo el principio de eliminación proactiva de código muerto.

Lección 29: La Normalización de Datos Previene Inconsistencias y Mejora la Integridad Referencial
Observación: El almacenamiento de valores descriptivos como cadenas de texto (ej: estadoInterno = "Borrador") en lugar de claves foráneas numéricas crea múltiples problemas: inconsistencias por variaciones tipográficas, dificultad para mantener integridad referencial, imposibilidad de cambiar descripciones centralizadamente, y mayor consumo de espacio de almacenamiento.
Regla Inquebrantable: Todos los campos que representen valores de un conjunto finito y controlado deben normalizarse usando claves foráneas numéricas (Long) que referencien tablas de configuración. Los valores descriptivos deben almacenarse únicamente en las tablas de referencia (ej: tbEstados) y obtenerse mediante JOINs cuando sea necesario mostrarlos al usuario.
Acción Correctiva: Identificar todos los campos de texto que contengan valores de conjuntos controlados (estados, tipos, categorías, etc.) y refactorizarlos a campos Long con claves foráneas. Crear o actualizar las tablas de referencia correspondientes (tbEstados, tbTiposSolicitud, etc.) con campos Clave (Long, PK) y Descripcion (Text). Actualizar todas las clases T_* para usar tipos Long en lugar de String para estos campos, y modificar los repositorios para realizar JOINs cuando sea necesario obtener las descripciones. Esto garantiza integridad referencial, elimina inconsistencias, centraliza el mantenimiento de valores de dominio y optimiza el almacenamiento.

## Lección 30: DAO y ADO no son Intercambiables

Observación: Las librerías DAO (Data Access Objects) y ADO (ActiveX Data Objects) proporcionan tipos de objeto incompatibles entre sí. Un DAO.Recordset no puede ser asignado a una variable Object que espera un ADODB.Recordset, y viceversa. Esta incompatibilidad genera errores de tiempo de ejecución difíciles de diagnosticar cuando se mezclan ambas tecnologías en la misma cadena de llamadas.
Regla Inquebrantable: Se debe mantener consistencia tecnológica en toda la cadena de acceso a datos. Si los repositorios utilizan DAO, todas las clases de servicio, pruebas unitarias y funciones auxiliares deben trabajar exclusivamente con tipos DAO (DAO.Database, DAO.Recordset). No se permite mezclar DAO y ADO en el mismo flujo de datos.
Acción Correctiva: Auditar todo el código que maneja recordsets para identificar inconsistencias entre DAO y ADO. Estandarizar en DAO para todo el proyecto CONDOR. Refactorizar las funciones auxiliares de pruebas que crean recordsets mock para que generen DAO.Recordset utilizando bases de datos temporales en disco en lugar de ADODB.Recordset en memoria. Actualizar todas las declaraciones de variables y parámetros de función para usar tipos DAO específicos (DAO.Recordset) en lugar de Object genérico.

## Lección 31: El Ciclo de Vida de los Objetos Debe Respetarse (Declarar, Instanciar, Usar, Liberar)

Observación: En VBA, una variable de objeto declarada con Dim solo reserva memoria para la referencia, pero no crea el objeto. Intentar acceder a propiedades o métodos de una variable de objeto no instanciada genera el error "Object variable or With block variable not set" (Error 91). Este error es especialmente común al trabajar con clases personalizadas como T_Expediente.
Regla Inquebrantable: Toda variable de objeto debe seguir el ciclo completo: 1) Declaración con Dim, 2) Instanciación con Set ... = New ..., 3) Uso de propiedades y métodos, 4) Liberación con Set ... = Nothing. La instanciación debe ocurrir inmediatamente después de la declaración y antes de cualquier uso del objeto.
Acción Correctiva: Revisar todo el código que declara variables de objeto para verificar que cada Dim vaya seguido de su correspondiente Set ... = New ... antes del primer uso. Implementar un patrón estándar en los métodos de servicio: declarar la variable, instanciarla inmediatamente, poblar sus propiedades, y devolverla. En caso de error, asegurar que la liberación (Set ... = Nothing) ocurra en la sección de limpieza del manejador de errores para evitar memory leaks.

## Lección 32: La Base de Datos como Código (Database as Code)

Observación: La gestión manual de cambios en el esquema y datos de la base de datos genera inconsistencias entre entornos, pérdida de trazabilidad de cambios, y dificultades para revertir modificaciones problemáticas. Los cambios manuales directos en la base de datos rompen el principio de versionado y automatización que rige el resto del proyecto.
Regla Inquebrantable: La base de datos debe tratarse como un artefacto más del proyecto que puede ser versionado, automatizado y gestionado a través de scripts. Todos los cambios en el esquema (CREATE, ALTER, DROP) y en los datos de configuración (INSERT, UPDATE, DELETE) deben realizarse exclusivamente a través de scripts SQL versionados y ejecutados mediante herramientas automatizadas.
Acción Correctiva: Implementar herramientas CLI que permitan ejecutar scripts SQL de forma automatizada y transaccional. Crear un repositorio de scripts SQL organizados por versiones o funcionalidades. Establecer un flujo de trabajo donde los cambios de base de datos se propongan, revisen y apliquen a través de scripts, no mediante modificaciones manuales. Esto garantiza la reproducibilidad, trazabilidad y consistencia de los cambios de base de datos entre todos los entornos del proyecto.

## Lección 33: Identificar y Replicar el "Estándar de Oro"
Observación: A lo largo del desarrollo, han surgido implementaciones que no solo son correctas, sino que representan la expresión ideal de nuestra arquitectura. Estos módulos, como CSolicitudService y su suite de pruebas, se convierten en el "Estándar de Oro".

Regla Inquebrantable: Cuando se identifica un módulo o patrón como "Estándar de Oro", debe ser utilizado como la plantilla de referencia obligatoria para el desarrollo de todas las funcionalidades futuras de naturaleza similar. Ignorar estos patrones y reinventar la rueda introduce inconsistencias y reduce la calidad general del proyecto.

Acción Correctiva: El Arquitecto (CONDOR-Expert) es responsable de identificar estos patrones de excelencia. Todos los prompts proactivos de refactorización deben instruir a Copilot para que tome estos módulos "Estándar de Oro" como base, asegurando que las mejores prácticas se repliquen de manera consistente en toda la base de código.

## Lección 35: El Principio de Responsabilidad Única se Aplica a Todas las Clases

Observación: Los errores de compilación como "ShowReport no existe" en modTestRunner revelan violaciones del Principio de Responsabilidad Única. Cuando una clase intenta realizar tareas que no son su responsabilidad principal, se generan dependencias incorrectas y métodos inexistentes. CTestReporter es un ejemplo perfecto: su responsabilidad es Generar reportes, no Mostrarlos.

Regla Inquebrantable: Cada clase debe tener un único propósito bien definido y claramente delimitado. CTestReporter debe limitarse a generar reportes como strings a través de su método GenerateReport(). Mostrar esos reportes al usuario (mediante MsgBox, formularios, o archivos) es responsabilidad de la capa de presentación, no del reportero. Esta separación permite reutilizar el reportero en diferentes contextos (CLI, UI, archivos) sin modificar su código.

Acción Correctiva: Auditar todas las clases del proyecto para verificar que cada una tenga una responsabilidad única y bien definida. Refactorizar cualquier clase que intente realizar múltiples responsabilidades, separando la lógica en clases especializadas. En el caso específico de CTestReporter, eliminar cualquier método de presentación (ShowReport) y mantener únicamente la generación de reportes (GenerateReport). La capa de presentación debe ser responsable de tomar el string generado y mostrarlo según el contexto apropiado.

Caso Crítico - CTestReporter (2024): La clase CTestReporter violaba el Principio de Responsabilidad Única al intentar tanto generar como mostrar reportes. La corrección implicó: 1) Eliminar el método ShowReport inexistente, 2) Usar únicamente GenerateReport() para obtener el reporte como string, 3) Delegar la presentación a la capa apropiada (MsgBox en RunTestFramework). Esta separación permite que CTestReporter sea reutilizable en CLI, UI y otros contextos sin modificaciones.